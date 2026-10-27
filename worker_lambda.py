"""
Worker Lambda — triggered by SQS (ReportBatchItemFailures must be enabled on the mapping).

GPT's mistake: it generated a separate FastAPI "Fluid" server (app.py) and had this
Lambda make an HTTP call to it.  That server never existed in your architecture.
All EC2 provisioning is done here directly with boto3.

Required Lambda environment variables:
  TABLE_NAME            — DynamoDB table name
  AMI_ID_LINUX          — AMI for Linux instances
  AMI_ID_WINDOWS        — AMI for Windows instances
  SUBNET_ID             — Subnet to launch into
  SECURITY_GROUP_IDS    — Comma-separated security group IDs
  INSTANCE_PROFILE      — IAM instance profile name (must allow SSM)
  SPLUNK_REALM          — e.g. us0
  SPLUNK_ACCESS_TOKEN   — Splunk ingest token (store in Secrets Manager and inject at deploy)
Optional:
  INSTANCE_TYPE         — default t3.small
  KEY_NAME              — EC2 key pair name (omit for SSM-only access)

Lambda timeout: set to 900 s (15 min) in the function configuration.
"""

import json
import os
import time
from datetime import datetime, timezone

import boto3

# ── AWS clients ───────────────────────────────────────────────────────────────
_region = os.environ.get("AWS_REGION")
table = boto3.resource("dynamodb").Table(os.environ["TABLE_NAME"])
ec2 = boto3.client("ec2", region_name=_region)
ssm = boto3.client("ssm", region_name=_region)

# ── Config ────────────────────────────────────────────────────────────────────
AMI_LINUX             = os.environ["AMI_ID_LINUX"]
AMI_WINDOWS           = os.environ["AMI_ID_WINDOWS"]
SUBNET_ID             = os.environ["SUBNET_ID"]
SECURITY_GROUP_IDS    = [x.strip() for x in os.environ["SECURITY_GROUP_IDS"].split(",") if x.strip()]
INSTANCE_PROFILE      = os.environ["INSTANCE_PROFILE"]
INSTANCE_TYPE         = os.environ.get("INSTANCE_TYPE", "t3.small")
KEY_NAME              = os.environ.get("KEY_NAME", "")
SPLUNK_REALM          = os.environ.get("SPLUNK_REALM", "")
SPLUNK_ACCESS_TOKEN   = os.environ.get("SPLUNK_ACCESS_TOKEN", "")


# ── DynamoDB helper ───────────────────────────────────────────────────────────
def set_status(gear_id: str, status: str, extra: dict | None = None) -> None:
    extra = extra or {}
    names      = {"#s": "status", "#u": "updated_at"}
    values     = {":s": status, ":u": datetime.now(timezone.utc).isoformat()}
    assignments = ["#s = :s", "#u = :u"]
    for i, (k, v) in enumerate(extra.items()):
        names[f"#n{i}"]  = k
        values[f":v{i}"] = v
        assignments.append(f"#n{i} = :v{i}")
    table.update_item(
        Key={"gear_id": gear_id},
        UpdateExpression="SET " + ", ".join(assignments),
        ExpressionAttributeNames=names,
        ExpressionAttributeValues=values,
    )


# ── User-data scripts ─────────────────────────────────────────────────────────
def linux_user_data() -> str:
    """Installs Apache only. Splunk OTel Collector is installed post-boot via SSM Run Command."""
    return """#!/bin/bash
set -euxo pipefail
if command -v dnf >/dev/null 2>&1; then dnf install -y httpd; else yum install -y httpd; fi
cat > /etc/httpd/conf.d/server-status.conf <<'EOF'
<Location "/server-status">
    SetHandler server-status
    Require ip 127.0.0.1
</Location>
ExtendedStatus On
EOF
systemctl enable --now httpd
"""


def windows_user_data() -> str:
    return r"""<powershell>
Install-WindowsFeature -Name Web-Server -IncludeManagementTools
Start-Service W3SVC
Set-Service W3SVC -StartupType Automatic
</powershell>"""


# ── EC2 launch ────────────────────────────────────────────────────────────────
def launch_ec2(payload: dict) -> str:
    """Launch the instance, wait until it passes EC2 status checks, return instance_id."""
    gear_id     = payload["gear_id"]
    os_type     = payload.get("os", "linux").lower()
    environment = payload.get("environment", "DEV")
    technology  = payload.get("technology", "apache")

    image_id  = AMI_LINUX  if os_type == "linux" else AMI_WINDOWS
    user_data = linux_user_data() if os_type == "linux" else windows_user_data()

    tags = [
        {"Key": "GearID",      "Value": gear_id},
        {"Key": "Environment", "Value": environment},
        {"Key": "Technology",  "Value": technology},
        {"Key": "ManagedBy",   "Value": "FluidProvisioning"},
    ]
    run_args: dict = {
        "ImageId":             image_id,
        "InstanceType":        INSTANCE_TYPE,
        "MinCount":            1,
        "MaxCount":            1,
        "SubnetId":            SUBNET_ID,
        "SecurityGroupIds":    SECURITY_GROUP_IDS,
        "IamInstanceProfile":  {"Name": INSTANCE_PROFILE},
        "UserData":            user_data,
        "TagSpecifications":   [{"ResourceType": "instance", "Tags": tags}],
    }
    if KEY_NAME:
        run_args["KeyName"] = KEY_NAME

    result      = ec2.run_instances(**run_args)
    instance_id = result["Instances"][0]["InstanceId"]

    # instance_status_ok means both EC2 + system checks passed — SSM agent is reliably up by then.
    # Delay=15 × MaxAttempts=26 ≈ 6.5 min max.
    ec2.get_waiter("instance_status_ok").wait(
        InstanceIds=[instance_id],
        WaiterConfig={"Delay": 15, "MaxAttempts": 26},
    )
    return instance_id


# ── SSM helpers ───────────────────────────────────────────────────────────────
def _wait_for_ssm_agent(instance_id: str, max_wait_s: int = 180) -> None:
    """
    After instance_status_ok the SSM agent should already be registered, but poll
    for up to max_wait_s seconds as a safety net before sending the command.
    """
    deadline = time.time() + max_wait_s
    while time.time() < deadline:
        info = ssm.describe_instance_information(
            Filters=[{"Key": "InstanceIds", "Values": [instance_id]}]
        )
        if info.get("InstanceInformationList"):
            return
        time.sleep(15)
    raise RuntimeError(
        f"SSM agent on {instance_id} did not register within {max_wait_s} s. "
        "Check the instance profile has AmazonSSMManagedInstanceCore."
    )


def install_otel_linux(instance_id: str) -> None:
    """Install the Splunk OTel Collector on a Linux instance via SSM Run Command."""
    if not SPLUNK_REALM or not SPLUNK_ACCESS_TOKEN:
        raise RuntimeError(
            "SPLUNK_REALM and SPLUNK_ACCESS_TOKEN must be set as Lambda environment variables."
        )

    _wait_for_ssm_agent(instance_id)

    commands = [
        "set -e",
        "curl -sSL https://dl.observability.splunkcloud.com/splunk-otel-collector.sh "
        "-o /tmp/splunk-otel-collector.sh",
        f"sh /tmp/splunk-otel-collector.sh --realm '{SPLUNK_REALM}' -- '{SPLUNK_ACCESS_TOKEN}'",
        "systemctl enable --now splunk-otel-collector",
    ]
    sent = ssm.send_command(
        InstanceIds=[instance_id],
        DocumentName="AWS-RunShellScript",
        Parameters={"commands": commands},
        TimeoutSeconds=300,
    )
    command_id = sent["Command"]["CommandId"]

    # Delay=10 × MaxAttempts=50 ≈ 8.3 min max.
    ssm.get_waiter("command_executed").wait(
        CommandId=command_id,
        InstanceId=instance_id,
        WaiterConfig={"Delay": 10, "MaxAttempts": 50},
    )
    invocation = ssm.get_command_invocation(CommandId=command_id, InstanceId=instance_id)
    if invocation.get("Status") != "Success":
        raise RuntimeError(
            f"Splunk OTel installation failed (SSM Status={invocation.get('Status')}). "
            "Inspect AWS Systems Manager → Run Command for stdout/stderr."
        )


# ── Lambda handler ────────────────────────────────────────────────────────────
def lambda_handler(event, context):
    failures = []
    for record in event.get("Records", []):
        payload = json.loads(record["body"])
        gear_id = str(payload["gear_id"])
        try:
            set_status(gear_id, "PROCESSING")
            instance_id = launch_ec2(payload)

            if payload.get("os", "linux").lower() == "linux":
                install_otel_linux(instance_id)

            set_status(gear_id, "CREATED", {"instance_id": instance_id, "error": ""})
        except Exception as exc:
            print(json.dumps({"level": "ERROR", "gear_id": gear_id, "error": str(exc)}))
            set_status(gear_id, "FAILED", {"error": str(exc)[:1000]})
            failures.append({"itemIdentifier": record["messageId"]})

    # SQS event source mapping must have ReportBatchItemFailures enabled.
    return {"batchItemFailures": failures}
