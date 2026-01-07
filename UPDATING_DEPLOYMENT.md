# Updating AgentCore MCP Server Deployment

This guide explains how to deploy a new version of your AgentCore MCP server after the initial deployment.

---

## Overview

When you update an AgentCore Runtime, AWS automatically:
- Creates a new immutable version (V1, V2, V3, etc.)
- Updates the `DEFAULT` endpoint to point to the latest version
- Preserves previous versions for rollback

**Key Points:**
- ✅ Updates are **non-destructive** - previous versions remain available
- ✅ `DEFAULT` endpoint automatically points to the latest version
- ✅ You can create custom endpoints pointing to specific versions
- ✅ Zero-downtime updates (new version is ready before traffic switches)

---

## Update Methods

Choose the method based on how you originally deployed:

### Method 1: Using AgentCore CLI (If You Used CLI for Initial Deployment)

If you used the AgentCore Starter Toolkit CLI for your initial deployment:

#### Step 1: Update Your Code

```bash
# Make your code changes
# Edit your server files, dependencies, etc.

# Update version in pyproject.toml (optional, for tracking)
# version = "0.2.0"
```

#### Step 2: Update Configuration (if needed)

```bash
# Edit agentcore.yaml if you need to change:
# - Environment variables
# - Request headers
# - OAuth settings
# - Network configuration

vim agentcore.yaml
```

#### Step 3: Deploy the Update

```bash
# The CLI will detect changes and create a new version
agentcore update

# Or if using launch command:
agentcore launch
```

**What happens:**
- CLI packages your updated code
- Uploads to S3/ECR (depending on your setup)
- Creates a new version of the runtime
- Updates the DEFAULT endpoint

#### Step 4: Verify the Update

```bash
# List runtimes to see current version
agentcore list

# Describe runtime to see version details
agentcore describe --runtime-name ms365-email-mcp-runtime

# View logs to ensure new version is working
agentcore logs --runtime-name ms365-email-mcp-runtime
```

---

### Method 2: Manual AWS CLI Update (S3-Based Deployment)

If you used manual AWS CLI deployment with S3:

#### Step 1: Prepare Updated Package

```bash
cd /path/to/email-mcp-server

# Make your code changes
# Edit server files, update dependencies, etc.

# Install dependencies and create updated package
uv pip install --target ./package .

# Copy your updated server code
cp -r ms365_email_mcp_server ./package/

# Create updated zip file
cd package
zip -r ../ms365-email-mcp-server-v2.zip .
cd ..
```

#### Step 2: Upload Updated Package to S3

```bash
# Set variables
export S3_BUCKET="my-agentcore-artifacts"
export AWS_REGION="us-gov-west-1"  # or your region
export VERSION="v2"  # Optional: for version tracking

# Upload new version
aws s3 cp ms365-email-mcp-server-v2.zip \
  s3://${S3_BUCKET}/mcp-servers/ms365-email-mcp-server-${VERSION}.zip \
  --region ${AWS_REGION}

# Or overwrite existing (if you want to keep same name)
aws s3 cp ms365-email-mcp-server-v2.zip \
  s3://${S3_BUCKET}/mcp-servers/ms365-email-mcp-server.zip \
  --region ${AWS_REGION}
```

#### Step 3: Update Runtime Configuration

Create an update configuration file `agentcore-runtime-update.json`:

```json
{
  "runtimeConfig": {
    "mcpRuntimeConfig": {
      "serverConfig": {
        "s3CodeConfig": {
          "s3Uri": "s3://my-agentcore-artifacts/mcp-servers/ms365-email-mcp-server-v2.zip",
          "handler": "ms365_email_mcp_server.server_token_only:main"
        },
        "runtime": "python3.11",
        "environmentVariables": {
          "MS365_CLOUD_TYPE": "gov",
          "LOG_LEVEL": "INFO",
          "HOST": "0.0.0.0",
          "PORT": "8000",
          "STATELESS_HTTP": "true"
        },
        "timeout": 300,
        "memorySize": 512
      },
      "transportConfig": {
        "type": "HTTP",
        "httpConfig": {
          "port": 8000
        }
      }
    }
  }
}
```

**Note:** Only include fields you want to update. Omitted fields remain unchanged.

#### Step 4: Update the Runtime

```bash
# Get your runtime ARN
export RUNTIME_ARN=$(aws bedrock-agentcore list-runtimes \
  --query "runtimes[?name=='ms365-email-mcp-runtime'].arn | [0]" \
  --output text \
  --region ${AWS_REGION})

# Update the runtime (creates new version)
aws bedrock-agentcore update-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --cli-input-json file://agentcore-runtime-update.json \
  --region ${AWS_REGION}
```

#### Step 5: Verify the Update

```bash
# List runtime versions
aws bedrock-agentcore list-agent-runtime-versions \
  --runtime-arn ${RUNTIME_ARN} \
  --region ${AWS_REGION}

# Get latest version info
aws bedrock-agentcore get-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --region ${AWS_REGION} \
  --query 'runtime.{Version:version,State:state,Endpoint:endpoint}'

# Test the updated runtime
export RUNTIME_ENDPOINT=$(aws bedrock-agentcore get-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --query 'runtime.endpoint' \
  --output text \
  --region ${AWS_REGION})

# Inbound auth is IAM (SigV4), so use an AWS SDK / SigV4-capable HTTP client.
# Ensure you pass the Graph token via the custom header:
# - X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization: Bearer <graph_token>
#
# Example payload:
# {"tool": "list-mail-messages", "parameters": {"top": 1}}
```

---

### Method 3: Update Only Configuration (No Code Changes)

If you only need to update configuration (environment variables, headers, OAuth settings) without changing code:

#### Update Environment Variables

```bash
export RUNTIME_ARN="arn:aws:bedrock-agentcore:us-gov-west-1:123456789012:runtime/ms365-email-mcp-runtime"

aws bedrock-agentcore update-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --runtime-config '{
    "mcpRuntimeConfig": {
      "serverConfig": {
        "environmentVariables": {
          "MS365_CLOUD_TYPE": "gov",
          "LOG_LEVEL": "DEBUG",  # Changed from INFO
          "HOST": "0.0.0.0",
          "PORT": "8000",
          "STATELESS_HTTP": "true"
        }
      }
    }
  }' \
  --region ${AWS_REGION}
```

#### Update Request Header Allowlist

```bash
aws bedrock-agentcore update-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --request-header-configuration '{
    "requestHeaderAllowlist": [
      "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization",
      "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier",
      "X-Custom-Header"  # Added new header
    ]
  }' \
  --region ${AWS_REGION}
```

#### Update OAuth/JWT Authorizer Configuration (Not used in IAM model)

If your runtime uses **IAM inbound authentication**, you do not configure an OAuth/JWT authorizer on AgentCore Runtime, so there is nothing to update here.

---

## Version Management

### List All Versions

```bash
aws bedrock-agentcore list-agent-runtime-versions \
  --runtime-arn ${RUNTIME_ARN} \
  --region ${AWS_REGION} \
  --query 'agentRuntimeVersions[*].{Version:version,CreatedAt:createdAt,State:state}'
```

### View Specific Version

```bash
# Get details of a specific version
aws bedrock-agentcore get-agent-runtime-version \
  --runtime-arn ${RUNTIME_ARN} \
  --version "v2" \
  --region ${AWS_REGION}
```

### Create Custom Endpoints for Version Control

You can create endpoints pointing to specific versions for staging/production:

```bash
# Create a production endpoint pointing to a specific version
aws bedrock-agentcore create-agent-runtime-endpoint \
  --runtime-arn ${RUNTIME_ARN} \
  --endpoint-name "production" \
  --agent-runtime-version "v2" \
  --description "Production endpoint on version 2" \
  --region ${AWS_REGION}

# Create a staging endpoint pointing to latest
aws bedrock-agentcore create-agent-runtime-endpoint \
  --runtime-arn ${RUNTIME_ARN} \
  --endpoint-name "staging" \
  --agent-runtime-version "latest" \
  --description "Staging endpoint always on latest" \
  --region ${AWS_REGION}
```

### Update Endpoint to New Version

```bash
# Update production endpoint to version 3
aws bedrock-agentcore update-agent-runtime-endpoint \
  --runtime-arn ${RUNTIME_ARN} \
  --endpoint-name "production" \
  --agent-runtime-version "v3" \
  --description "Updated production to version 3" \
  --region ${AWS_REGION}
```

---

## Rollback to Previous Version

If you need to rollback to a previous version:

### Option 1: Update DEFAULT Endpoint to Previous Version

```bash
# List versions to find the version you want
aws bedrock-agentcore list-agent-runtime-versions \
  --runtime-arn ${RUNTIME_ARN} \
  --region ${AWS_REGION}

# Update DEFAULT endpoint to previous version (e.g., v2)
aws bedrock-agentcore update-agent-runtime-endpoint \
  --runtime-arn ${RUNTIME_ARN} \
  --endpoint-name "DEFAULT" \
  --agent-runtime-version "v2" \
  --region ${AWS_REGION}
```

### Option 2: Create New Endpoint Pointing to Previous Version

```bash
# Create rollback endpoint
aws bedrock-agentcore create-agent-runtime-endpoint \
  --runtime-arn ${RUNTIME_ARN} \
  --endpoint-name "rollback" \
  --agent-runtime-version "v2" \
  --description "Rollback endpoint" \
  --region ${AWS_REGION}
```

---

## Best Practices

### 1. Version Your Code Packages

```bash
# Use versioned S3 keys
s3://bucket/mcp-servers/ms365-email-mcp-server-v1.0.0.zip
s3://bucket/mcp-servers/ms365-email-mcp-server-v1.1.0.zip
s3://bucket/mcp-servers/ms365-email-mcp-server-v2.0.0.zip
```

### 2. Test Before Production Update

```bash
# 1. Deploy to a test/staging endpoint first
# 2. Test thoroughly
# 3. Update production endpoint after validation
```

### 3. Monitor After Update

```bash
# Check CloudWatch logs
aws logs tail /aws/bedrock-agentcore/runtime/${RUNTIME_NAME} --follow

# Check metrics
aws cloudwatch get-metric-statistics \
  --namespace AWS/BedrockAgentCore \
  --metric-name Invocations \
  --dimensions Name=RuntimeArn,Value=${RUNTIME_ARN} \
  --start-time $(date -u -d '1 hour ago' +%Y-%m-%dT%H:%M:%S) \
  --end-time $(date -u +%Y-%m-%dT%H:%M:%S) \
  --period 300 \
  --statistics Sum
```

### 4. Keep Previous Versions

- Don't delete old S3 packages immediately
- Keep at least 2-3 previous versions for rollback
- Document what changed in each version

### 5. Use Infrastructure as Code

Consider using CloudFormation, Terraform, or CDK to manage updates:

```yaml
# Example CloudFormation update
Resources:
  AgentCoreRuntime:
    Type: AWS::Bedrock::AgentCore::Runtime
    Properties:
      RuntimeName: ms365-email-mcp-runtime
      RuntimeConfig:
        MCPRuntimeConfig:
          ServerConfig:
            S3CodeConfig:
              S3Uri: !Sub "s3://${S3Bucket}/mcp-servers/ms365-email-mcp-server-v2.zip"
```

---

## Common Update Scenarios

### Scenario 1: Update Dependencies

```bash
# 1. Update pyproject.toml with new dependency versions
# 2. Rebuild package
uv pip install --target ./package .

# 3. Upload and update runtime
aws s3 cp ms365-email-mcp-server.zip s3://${S3_BUCKET}/mcp-servers/
aws bedrock-agentcore update-runtime --runtime-arn ${RUNTIME_ARN} ...
```

### Scenario 2: Fix a Bug

```bash
# 1. Fix the bug in your code
# 2. Test locally
python -m ms365_email_mcp_server.server_token_only

# 3. Package and deploy
# (follow steps from Method 2 above)
```

### Scenario 3: Add New Environment Variable

```bash
# Update runtime with new environment variable
aws bedrock-agentcore update-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --runtime-config '{
    "mcpRuntimeConfig": {
      "serverConfig": {
        "environmentVariables": {
          "MS365_CLOUD_TYPE": "gov",
          "LOG_LEVEL": "INFO",
          "HOST": "0.0.0.0",
          "PORT": "8000",
          "STATELESS_HTTP": "true",
          "NEW_FEATURE_FLAG": "enabled"  # New variable
        }
      }
    }
  }' \
  --region ${AWS_REGION}
```

### Scenario 4: Add a New Agent/Caller (IAM model)

If inbound authentication is **IAM**, adding a new agent/caller means granting it IAM permissions to invoke the runtime (and ensuring it can mint/provide a Graph token to pass via the custom header).

---

## Troubleshooting Updates

### Issue: Update Fails

**Symptoms:**
- `update-runtime` command fails
- Runtime state shows `UPDATE_FAILED`

**Solutions:**
```bash
# Check error details
aws bedrock-agentcore get-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --region ${AWS_REGION} \
  --query 'runtime.stateReason'

# Verify S3 package is accessible
aws s3 ls s3://${S3_BUCKET}/mcp-servers/ms365-email-mcp-server.zip

# Check IAM permissions
aws iam simulate-principal-policy \
  --policy-source-arn ${EXECUTION_ROLE_ARN} \
  --action-names bedrock-agentcore:UpdateRuntime \
  --resource-arns ${RUNTIME_ARN}
```

### Issue: New Version Not Working

**Symptoms:**
- Update succeeds but runtime returns errors
- Health checks failing

**Solutions:**
```bash
# Check CloudWatch logs
aws logs tail /aws/bedrock-agentcore/runtime/${RUNTIME_NAME} --follow

# Verify code package integrity
unzip -t ms365-email-mcp-server.zip

# Test locally first
python -m ms365_email_mcp_server.server_token_only

# Rollback if needed
aws bedrock-agentcore update-agent-runtime-endpoint \
  --runtime-arn ${RUNTIME_ARN} \
  --endpoint-name "DEFAULT" \
  --agent-runtime-version "v2"  # Previous working version
```

### Issue: DEFAULT Endpoint Not Updating

**Symptoms:**
- Update creates new version but DEFAULT still points to old version

**Solutions:**
```bash
# Check endpoint state
aws bedrock-agentcore list-agent-runtime-endpoints \
  --runtime-arn ${RUNTIME_ARN} \
  --region ${AWS_REGION}

# Manually update DEFAULT endpoint
aws bedrock-agentcore update-agent-runtime-endpoint \
  --runtime-arn ${RUNTIME_ARN} \
  --endpoint-name "DEFAULT" \
  --agent-runtime-version "latest" \
  --region ${AWS_REGION}
```

---

## Quick Reference

### Update Commands Summary

```bash
# Method 1: AgentCore CLI
agentcore update

# Method 2: AWS CLI - Full Update
aws bedrock-agentcore update-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --cli-input-json file://update-config.json

# Method 3: AWS CLI - Configuration Only
aws bedrock-agentcore update-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --runtime-config '{"mcpRuntimeConfig": {...}}'

# List Versions
aws bedrock-agentcore list-agent-runtime-versions \
  --runtime-arn ${RUNTIME_ARN}

# Rollback
aws bedrock-agentcore update-agent-runtime-endpoint \
  --runtime-arn ${RUNTIME_ARN} \
  --endpoint-name "DEFAULT" \
  --agent-runtime-version "v2"
```

---

## Summary

✅ **Updates create new immutable versions**  
✅ **DEFAULT endpoint automatically points to latest**  
✅ **Previous versions remain available for rollback**  
✅ **Zero-downtime updates**  
✅ **Support for staging/production endpoints**

Choose the update method that matches your deployment approach, test thoroughly, and keep previous versions for rollback capability.

