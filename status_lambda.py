"""
Status Lambda — triggered by API Gateway GET /status?gear_id=GEAR001

Returns the current DynamoDB record for the given gear_id.
Used by the Excel VBA macro (RefreshGearStatus) to poll provisioning progress.

Required Lambda environment variables:
  TABLE_NAME — DynamoDB table name
"""

import json
import os
from decimal import Decimal

import boto3

table = boto3.resource("dynamodb").Table(os.environ["TABLE_NAME"])


def _default(value):
    """JSON-serialise Decimal values that DynamoDB returns."""
    if isinstance(value, Decimal):
        return int(value) if value % 1 == 0 else float(value)
    raise TypeError(f"Unsupported type: {type(value).__name__}")


def lambda_handler(event, context):
    params  = event.get("queryStringParameters") or {}
    gear_id = str(params.get("gear_id", "")).strip()

    if not gear_id:
        return {
            "statusCode": 400,
            "headers": {"Content-Type": "application/json"},
            "body": json.dumps({"error": "gear_id is required"}),
        }

    result = table.get_item(Key={"gear_id": gear_id}, ConsistentRead=True)
    item   = result.get("Item")

    if not item:
        return {
            "statusCode": 404,
            "headers": {"Content-Type": "application/json"},
            "body": json.dumps({"error": "gear_id not found"}),
        }

    return {
        "statusCode": 200,
        "headers": {"Content-Type": "application/json"},
        "body": json.dumps(item, default=_default),
    }
