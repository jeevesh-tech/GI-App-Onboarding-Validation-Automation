"""
Validator Lambda — triggered by API Gateway (Lambda proxy integration).

Validates the incoming provisioning request, writes a NEW record to DynamoDB
(conditional put prevents duplicate gear IDs, including concurrent races),
then enqueues the job to SQS for the Worker Lambda to pick up.

Required Lambda environment variables:
  TABLE_NAME   — DynamoDB table name
  QUEUE_URL    — SQS queue URL
"""

import json
import os
from datetime import datetime, timezone

import boto3
from botocore.exceptions import ClientError

table    = boto3.resource("dynamodb").Table(os.environ["TABLE_NAME"])
sqs      = boto3.client("sqs")
QUEUE_URL = os.environ["QUEUE_URL"]

VALID_OS          = {"linux", "windows"}
VALID_ENVIRONMENTS = {"DEV", "QA", "PROD"}
VALID_TECHNOLOGIES = {"apache", "iis", "java", ".net", "python"}


def reply(code: int, payload: dict) -> dict:
    return {
        "statusCode": code,
        "headers": {"Content-Type": "application/json"},
        "body": json.dumps(payload),
    }


def lambda_handler(event, context):
    try:
        body = event.get("body") or {}
        if isinstance(body, str):
            body = json.loads(body)

        gear_id     = str(body.get("gear_id", "")).strip()
        os_name     = str(body.get("os", body.get("operating_system", "linux"))).lower().strip()
        region      = str(body.get("region", os.environ.get("AWS_REGION", ""))).strip()
        environment = str(body.get("environment", "DEV")).upper().strip()
        technology  = str(body.get("technology", "apache")).lower().strip()

        # ── Validation ────────────────────────────────────────────────────────
        if not gear_id:
            return reply(400, {"error": "gear_id is required"})
        if os_name not in VALID_OS:
            return reply(400, {"error": f"os must be one of: {', '.join(sorted(VALID_OS))}"})
        if environment not in VALID_ENVIRONMENTS:
            return reply(400, {"error": f"environment must be one of: {', '.join(sorted(VALID_ENVIRONMENTS))}"})
        if technology not in VALID_TECHNOLOGIES:
            return reply(400, {"error": f"technology must be one of: {', '.join(sorted(VALID_TECHNOLOGIES))}"})

        # ── Write to DynamoDB (conditional — prevents duplicate gear IDs) ─────
        now  = datetime.now(timezone.utc).isoformat()
        item = {
            "gear_id":     gear_id,
            "os":          os_name,
            "region":      region,
            "environment": environment,
            "technology":  technology,
            "status":      "NEW",
            "flag":        "",          # populated by the macro or admin when a problem is flagged
            "instance_id": "",
            "error":       "",
            "created_at":  now,
            "updated_at":  now,
        }
        table.put_item(
            Item=item,
            ConditionExpression="attribute_not_exists(gear_id)",
        )

        # ── Enqueue to SQS ────────────────────────────────────────────────────
        sqs.send_message(
            QueueUrl=QUEUE_URL,
            MessageBody=json.dumps({
                "gear_id":     gear_id,
                "os":          os_name,
                "region":      region,
                "environment": environment,
                "technology":  technology,
            }),
        )

        return reply(202, {
            "gear_id": gear_id,
            "status":  "NEW",
            "message": "Provisioning request accepted",
        })

    except ClientError as exc:
        if exc.response.get("Error", {}).get("Code") == "ConditionalCheckFailedException":
            return reply(409, {"error": "gear_id already exists"})
        print(json.dumps({"level": "ERROR", "error": str(exc)}))
        return reply(500, {"error": "AWS request failed"})
    except (ValueError, TypeError, json.JSONDecodeError) as exc:
        return reply(400, {"error": f"Invalid request: {exc}"})
    except Exception as exc:
        print(json.dumps({"level": "ERROR", "error": str(exc)}))
        return reply(500, {"error": "Internal server error"})
