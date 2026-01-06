# AWS AgentCore Runtime Deployment Guide

This guide walks through deploying the MS365 Email MCP Server as an AWS AgentCore Runtime using **S3-based code upload** with **OAuth 2.0 JWT authentication** for the US Government cloud endpoint.

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Deployment Methods](#deployment-methods)
   - [Option A: Interactive AgentCore CLI (Recommended for First-Time Setup)](#option-a-interactive-agentcore-cli-recommended-for-first-time-setup)
   - [Option B: Manual AWS CLI Deployment with S3](#option-b-manual-aws-cli-deployment-with-s3)
3. [Architecture Overview](#architecture-overview)
4. [Step 1: Prepare Your MCP Server Package](#step-1-prepare-your-mcp-server-package)
5. [Step 2: Upload to S3](#step-2-upload-to-s3)
6. [Step 3: Create IAM Roles](#step-3-create-iam-roles)
7. [Step 4: Configure OAuth 2.0 JWT Authorizer](#step-4-configure-oauth-20-jwt-authorizer)
8. [Step 5: Create AgentCore Runtime](#step-5-create-agentcore-runtime)
9. [Step 6: Configure Request Header Allowlist](#step-6-configure-request-header-allowlist)
10. [Step 7: Test Your Deployment](#step-7-test-your-deployment)
11. [Troubleshooting](#troubleshooting)

---

## Prerequisites

- **AWS CLI** configured with appropriate credentials
- **Python 3.11+** and **UV** package manager installed
- **Microsoft Entra ID** application registration with:
  - Client ID
  - Tenant ID
  - Configured scopes for Microsoft Graph API
- **S3 bucket** for code artifacts (for Option B)
- **AWS IAM permissions** to create AgentCore runtimes, IAM roles, and S3 operations

---

## Deployment Methods

Choose one of the following deployment approaches based on your preferences:

### Option A: Interactive AgentCore CLI (Recommended for First-Time Setup)

The **AgentCore Starter Toolkit** provides an interactive CLI that guides you through the deployment process with prompts for all required configuration. This is the easiest way to get started.

#### A.1 Install the AgentCore CLI

```bash
pip install bedrock-agentcore-starter-toolkit
```

**Note:** The AgentCore CLI typically uses **Amazon ECR** for containerized deployments. If you specifically need S3-based deployment, use **Option B** below.

#### A.2 Initialize Your Project

Navigate to your project directory and initialize the AgentCore configuration:

```bash
cd /path/to/email-mcp-server

# Initialize with your entry point file
agentcore init --entry-point ms365_email_mcp_server/server_token_only.py --protocol MCP
```

This creates an `agentcore.yaml` configuration file in your project.

#### A.3 Configure Interactively

Run the interactive configuration wizard:

```bash
agentcore configure
```

The CLI will prompt you for:

1. **Runtime Name**: 
   ```
   Enter runtime name [ms365-email-mcp-runtime]:
   ```

2. **Execution Role**: 
   ```
   Enter IAM execution role ARN (or press Enter to create one):
   ```
   - Press `Enter` to have the CLI create the role automatically, or
   - Provide your existing role ARN

3. **OAuth 2.0 Configuration**:
   ```
   Enable OAuth 2.0 JWT authentication? (yes/no) [yes]: yes
   
   Enter discovery URL: https://login.microsoftonline.us/<YOUR_TENANT_ID>/v2.0/.well-known/openid-configuration
   
   Enter allowed audiences (comma-separated): api://<YOUR_CLIENT_ID>,<YOUR_CLIENT_ID>
   
   Enter allowed client IDs (comma-separated): <CLIENT_ID_1>,<CLIENT_ID_2>
   
   Enter allowed scopes (comma-separated): Mail.Read,Mail.ReadWrite,Mail.Send
   ```

4. **Environment Variables**:
   ```
   Enter environment variables (key=value, comma-separated):
   MS365_CLOUD_TYPE=gov,LOG_LEVEL=INFO,HOST=0.0.0.0,PORT=8100,STATELESS_HTTP=true
   ```

5. **Request Header Allowlist**:
   ```
   Enter request headers to allowlist (comma-separated):
   Authorization,X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier
   ```

6. **Cloud Configuration** (if prompted):
   ```
   Select cloud type [commercial/gov]: gov
   ```

#### A.4 Review Configuration

The CLI creates an `agentcore.yaml` file. Review it:

```bash
cat agentcore.yaml
```

You should see something like:

```yaml
runtime:
  name: ms365-email-mcp-runtime
  protocol: MCP
  entryPoint: ms365_email_mcp_server/server_token_only.py
  runtime: python3.11
  
authorizer:
  type: JWT_BEARER_TOKEN
  discoveryUrl: https://login.microsoftonline.us/<TENANT_ID>/v2.0/.well-known/openid-configuration
  allowedAudiences:
    - api://<CLIENT_ID>
    - <CLIENT_ID>
  allowedClients:
    - <CLIENT_ID_1>
    - <CLIENT_ID_2>
  allowedScopes:
    - Mail.Read
    - Mail.ReadWrite
    - Mail.Send

requestHeaders:
  allowlist:
    - Authorization
    - X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier

environment:
  MS365_CLOUD_TYPE: gov
  LOG_LEVEL: INFO
  HOST: 0.0.0.0
  PORT: "8100"
  STATELESS_HTTP: "true"
```

#### A.5 Deploy

Deploy your MCP server interactively:

```bash
agentcore launch
```

This command will:
- Build your MCP server (package dependencies)
- Create or use an ECR repository (if using container deployment)
- Create the IAM execution role (if you didn't provide one)
- Create the AgentCore runtime with all configurations
- Deploy the runtime

**Output:**
```
✓ Building deployment package...
✓ Uploading to S3/ECR...
✓ Creating AgentCore runtime...
✓ Deployment complete!

Runtime ARN: arn:aws:bedrock-agentcore:us-gov-west-1:123456789012:runtime/ms365-email-mcp-runtime
Runtime Endpoint: https://runtime.bedrock-agentcore.us-gov-west-1.amazonaws.com/...
```

#### A.6 Verify Deployment

```bash
# List your runtimes
agentcore list

# Get runtime details
agentcore describe --runtime-name ms365-email-mcp-runtime
```

#### A.7 Update Configuration (if needed)

To update settings after deployment:

```bash
# Edit agentcore.yaml manually, then:
agentcore update

# Or run configure again:
agentcore configure
agentcore update
```

**Note:** The AgentCore CLI may use ECR by default. If you specifically need **S3-based deployment** (no ECR), see **Option B** below.

---

### Option B: Manual AWS CLI Deployment with S3

This approach gives you full control over the deployment process and uses **S3 for code storage** (no ECR required). Follow the detailed steps below starting from [Step 1](#step-1-prepare-your-mcp-server-package).

**Use this option if:**
- You prefer S3-based deployment over ECR
- You need more control over the deployment process
- You want to script/automate the deployment
- You're familiar with AWS CLI and JSON configuration

**Continue reading below for the complete manual deployment guide.**

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                         Agent/Client                            │
│  (Strands Agent, MCP Client with JWT token from Entra ID)      │
└────────────────────────┬────────────────────────────────────────┘
                         │ Authorization: Bearer <jwt_token>
                         │ X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│              AWS AgentCore Runtime (JWT Authorizer)             │
│  - Validates JWT against login.microsoftonline.us (GCC)        │
│  - Passes Authorization header through to MCP server            │
│  - Passes custom headers (user_identifier)                      │
└────────────────────────┬────────────────────────────────────────┘
                         │ Authorization: Bearer <jwt_token>
                         │ X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│           MS365 Email MCP Server (server_token_only.py)         │
│  - Extracts token from Authorization header                     │
│  - Extracts user_identifier from custom header                  │
│  - Uses token directly for Microsoft Graph API calls            │
└────────────────────────┬────────────────────────────────────────┘
                         │ Authorization: Bearer <jwt_token>
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│              Microsoft Graph API (graph.microsoft.us)           │
│                    US Government Cloud (GCC)                    │
└─────────────────────────────────────────────────────────────────┘
```

**Key Points:**
- The same JWT token is used for **both** AgentCore authentication **and** Graph API access
- AgentCore validates the JWT against Microsoft Entra ID (GCC endpoint)
- The token is passed through to the MCP server via the `Authorization` header
- No client secrets are passed in headers (security best practice)

---

## Step 1: Prepare Your MCP Server Package

### 1.1 Package the Server

```bash
cd /path/to/email-mcp-server

# Install dependencies and create a deployment package
uv pip install --target ./package .

# Copy your server code
cp -r ms365_email_mcp_server ./package/

# Create a zip file
cd package
zip -r ../ms365-email-mcp-server.zip .
cd ..
```

### 1.2 Verify Package Contents

```bash
unzip -l ms365-email-mcp-server.zip | head -20
```

You should see:
- `ms365_email_mcp_server/server_token_only.py`
- `fastmcp/`
- `httpx/`
- `loguru/`
- Other dependencies

---

## Step 2: Upload to S3

### 2.1 Create S3 Bucket (if needed)

```bash
# Replace with your desired bucket name and region
export S3_BUCKET="my-agentcore-artifacts"
export AWS_REGION="us-gov-west-1"  # or us-gov-east-1 for GCC

aws s3 mb s3://${S3_BUCKET} --region ${AWS_REGION}
```

### 2.2 Upload the Package

```bash
aws s3 cp ms365-email-mcp-server.zip \
  s3://${S3_BUCKET}/mcp-servers/ms365-email-mcp-server.zip \
  --region ${AWS_REGION}
```

### 2.3 Note the S3 URI

```bash
export S3_CODE_URI="s3://${S3_BUCKET}/mcp-servers/ms365-email-mcp-server.zip"
echo "S3 Code URI: ${S3_CODE_URI}"
```

---

## Step 3: Create IAM Roles

### 3.1 Create Execution Role

Create a file `agentcore-execution-role-trust-policy.json`:

```json
{
  "Version": "2012-10-17",
  "Statement": [
    {
      "Effect": "Allow",
      "Principal": {
        "Service": "bedrock-agentcore.amazonaws.com"
      },
      "Action": "sts:AssumeRole"
    }
  ]
}
```

Create the role:

```bash
aws iam create-role \
  --role-name MS365EmailMCPExecutionRole \
  --assume-role-policy-document file://agentcore-execution-role-trust-policy.json \
  --region ${AWS_REGION}
```

### 3.2 Attach Policies to Execution Role

```bash
# Basic execution permissions
aws iam attach-role-policy \
  --role-name MS365EmailMCPExecutionRole \
  --policy-arn arn:aws:iam::aws:policy/service-role/AWSLambdaBasicExecutionRole \
  --region ${AWS_REGION}

# S3 read access for code
aws iam put-role-policy \
  --role-name MS365EmailMCPExecutionRole \
  --policy-name S3CodeAccess \
  --policy-document '{
    "Version": "2012-10-17",
    "Statement": [
      {
        "Effect": "Allow",
        "Action": [
          "s3:GetObject",
          "s3:GetObjectVersion"
        ],
        "Resource": "arn:aws:s3:::'"${S3_BUCKET}"'/mcp-servers/*"
      }
    ]
  }' \
  --region ${AWS_REGION}
```

### 3.3 Get Role ARN

```bash
export EXECUTION_ROLE_ARN=$(aws iam get-role \
  --role-name MS365EmailMCPExecutionRole \
  --query 'Role.Arn' \
  --output text \
  --region ${AWS_REGION})

echo "Execution Role ARN: ${EXECUTION_ROLE_ARN}"
```

---

## Step 4: Configure OAuth 2.0 JWT Authorizer

### 4.1 Set Microsoft Entra ID Variables

```bash
# For US Government Cloud (GCC)
export TENANT_ID="your-tenant-id"
export CLIENT_ID="your-client-id"  # Your agent's client ID
export DISCOVERY_URL="https://login.microsoftonline.us/${TENANT_ID}/v2.0/.well-known/openid-configuration"

# For Commercial Cloud (use this instead if not using GCC)
# export DISCOVERY_URL="https://login.microsoftonline.com/${TENANT_ID}/v2.0/.well-known/openid-configuration"
```

**Important:** The discovery URL for US Government cloud uses `login.microsoftonline.us` as documented in [Microsoft's GCC documentation](https://learn.microsoft.com/en-us/azure/active-directory/develop/authentication-national-cloud).

### 4.2 Create JWT Authorizer Configuration

Create a file `jwt-authorizer-config.json`:

```json
{
  "type": "JWT_BEARER_TOKEN",
  "jwtBearerTokenAuthorizerConfig": {
    "discoveryUrl": "https://login.microsoftonline.us/<TENANT_ID>/v2.0/.well-known/openid-configuration",
    "allowedAudiences": [
      "api://<CLIENT_ID>",
      "<CLIENT_ID>"
    ],
    "allowedClients": [
      "<CLIENT_ID_1>",
      "<CLIENT_ID_2>"
    ],
    "allowedScopes": [
      "Mail.Read",
      "Mail.ReadWrite",
      "Mail.Send"
    ],
    "customClaims": {
      "upn": "upn",
      "oid": "oid"
    }
  }
}
```

**Replace placeholders:**
- `<TENANT_ID>`: Your Azure AD tenant ID
- `<CLIENT_ID>`: Your primary client ID (used in audience validation)
- `<CLIENT_ID_1>`, `<CLIENT_ID_2>`: Client IDs of all agents that will consume this MCP server

**Configuration Details:**

| Field | Description |
|-------|-------------|
| `discoveryUrl` | OpenID Connect discovery endpoint for JWT validation. Use `login.microsoftonline.us` for GCC, `login.microsoftonline.com` for commercial. |
| `allowedAudiences` | List of valid `aud` claims in the JWT. Typically `api://{clientId}` or the client ID itself. |
| `allowedClients` | List of client IDs (`azp` or `appid` claims) allowed to access this runtime. Add all agent client IDs here. |
| `allowedScopes` | Required OAuth scopes in the JWT token. Match these to your Microsoft Graph API permissions. |
| `customClaims` | Map JWT claims to custom attributes (optional). Useful for extracting user principal name or object ID. |

**Reference:** [AWS AgentCore JWT Authorizer Documentation](https://docs.aws.amazon.com/bedrock-agentcore-control/latest/APIReference/API_AuthorizerConfiguration.html)

---

## Step 5: Create AgentCore Runtime

### 5.1 Create Runtime Configuration

Create a file `agentcore-runtime-config.json`:

```json
{
  "name": "ms365-email-mcp-runtime",
  "description": "MS365 Email MCP Server with OAuth 2.0 JWT authentication for US Government Cloud",
  "runtimeType": "MCP",
  "runtimeConfig": {
    "mcpRuntimeConfig": {
      "serverConfig": {
        "s3CodeConfig": {
          "s3Uri": "s3://my-agentcore-artifacts/mcp-servers/ms365-email-mcp-server.zip",
          "handler": "ms365_email_mcp_server.server_token_only:main"
        },
        "runtime": "python3.11",
        "environmentVariables": {
          "MS365_CLOUD_TYPE": "gov",
          "LOG_LEVEL": "INFO",
          "HOST": "0.0.0.0",
          "PORT": "8100",
          "STATELESS_HTTP": "true"
        },
        "timeout": 300,
        "memorySize": 512
      },
      "transportConfig": {
        "type": "HTTP",
        "httpConfig": {
          "port": 8100
        }
      }
    }
  },
  "authorizerConfiguration": {
    "type": "JWT_BEARER_TOKEN",
    "jwtBearerTokenAuthorizerConfig": {
      "discoveryUrl": "https://login.microsoftonline.us/<TENANT_ID>/v2.0/.well-known/openid-configuration",
      "allowedAudiences": [
        "api://<CLIENT_ID>",
        "<CLIENT_ID>"
      ],
      "allowedClients": [
        "<CLIENT_ID_1>",
        "<CLIENT_ID_2>"
      ],
      "allowedScopes": [
        "Mail.Read",
        "Mail.ReadWrite",
        "Mail.Send"
      ]
    }
  },
  "requestHeaderConfiguration": {
    "requestHeaderAllowlist": [
      "Authorization",
      "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier"
    ]
  },
  "executionRoleArn": "arn:aws:iam::123456789012:role/MS365EmailMCPExecutionRole"
}
```

**Replace placeholders:**
- `s3Uri`: Your S3 URI from Step 2.3
- `<TENANT_ID>`: Your Azure AD tenant ID
- `<CLIENT_ID>`: Your primary client ID
- `<CLIENT_ID_1>`, `<CLIENT_ID_2>`: Client IDs of agents
- `executionRoleArn`: Your execution role ARN from Step 3.3

**Key Configuration Sections:**

#### S3 Code Configuration
```json
"s3CodeConfig": {
  "s3Uri": "s3://bucket/path/to/code.zip",
  "handler": "ms365_email_mcp_server.server_token_only:main"
}
```
- `s3Uri`: S3 location of your packaged code
- `handler`: Python module path to the entry point function

#### Environment Variables
```json
"environmentVariables": {
  "MS365_CLOUD_TYPE": "gov",  // "gov" for GCC, "commercial" for standard
  "LOG_LEVEL": "INFO",
  "HOST": "0.0.0.0",
  "PORT": "8100",
  "STATELESS_HTTP": "true"
}
```

#### Request Header Allowlist (Critical!)
```json
"requestHeaderAllowlist": [
  "Authorization",
  "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier"
]
```

**Why this matters:**
- `Authorization`: Passes the JWT token through to the MCP server for Graph API calls
- `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier`: Passes the user identifier for shared mailbox access

**Reference:** [RequestHeaderConfiguration API Documentation](https://docs.aws.amazon.com/bedrock-agentcore-control/latest/APIReference/API_RequestHeaderConfiguration.html)

### 5.2 Create the Runtime

```bash
aws bedrock-agentcore create-runtime \
  --cli-input-json file://agentcore-runtime-config.json \
  --region ${AWS_REGION}
```

### 5.3 Get Runtime ARN

```bash
export RUNTIME_ARN=$(aws bedrock-agentcore list-runtimes \
  --query "runtimes[?name=='ms365-email-mcp-runtime'].arn | [0]" \
  --output text \
  --region ${AWS_REGION})

echo "Runtime ARN: ${RUNTIME_ARN}"
```

---

## Step 6: Configure Request Header Allowlist

The `requestHeaderAllowlist` is **critical** for the token-only authentication flow. It tells AgentCore Runtime which HTTP headers to pass through to your MCP server.

### 6.1 Understanding Header Pass-Through

```
Agent Request Headers:
├── Authorization: Bearer <jwt_token>
├── X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier: user@domain.com
└── Other headers...

                    ↓ (AgentCore validates JWT)

AgentCore passes through (based on allowlist):
├── Authorization: Bearer <jwt_token>  ✅ (in allowlist)
├── X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier: user@domain.com  ✅ (in allowlist)
└── Other headers... ❌ (not in allowlist, dropped)

                    ↓

MCP Server receives:
├── Authorization: Bearer <jwt_token>
└── X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier: user@domain.com
```

### 6.2 Header Configuration in Runtime

The configuration from Step 5.1 includes:

```json
"requestHeaderConfiguration": {
  "requestHeaderAllowlist": [
    "Authorization",
    "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier"
  ]
}
```

**Header Descriptions:**

| Header | Purpose | Required |
|--------|---------|----------|
| `Authorization` | Contains `Bearer <jwt_token>` for Microsoft Graph API access | ✅ Yes |
| `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier` | UserPrincipalName or Graph ID for shared mailbox access | Optional |

### 6.3 Update Existing Runtime (if needed)

If you need to update the header allowlist on an existing runtime:

```bash
aws bedrock-agentcore update-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --request-header-configuration '{
    "requestHeaderAllowlist": [
      "Authorization",
      "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier"
    ]
  }' \
  --region ${AWS_REGION}
```

---

## Step 7: Test Your Deployment

### 7.1 Obtain a JWT Token

Use MSAL or Azure CLI to get a token:

```bash
# Using Azure CLI (for testing)
az login --use-device-code --cloud AzureUSGovernment  # For GCC
# or
az login --use-device-code  # For commercial

# Get token for Microsoft Graph
export JWT_TOKEN=$(az account get-access-token \
  --resource https://graph.microsoft.us \
  --query accessToken \
  --output tsv)

echo "JWT Token (first 50 chars): ${JWT_TOKEN:0:50}..."
```

**For production agents**, use MSAL Python or JavaScript:

```python
from msal import ConfidentialClientApplication

app = ConfidentialClientApplication(
    client_id="your-client-id",
    client_credential="your-client-secret",
    authority="https://login.microsoftonline.us/your-tenant-id"
)

result = app.acquire_token_for_client(scopes=["https://graph.microsoft.us/.default"])
jwt_token = result["access_token"]
```

### 7.2 Test MCP Server Directly

```bash
# Get runtime endpoint
export RUNTIME_ENDPOINT=$(aws bedrock-agentcore get-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --query 'runtime.endpoint' \
  --output text \
  --region ${AWS_REGION})

# Test list-mail-messages tool
curl -X POST "${RUNTIME_ENDPOINT}/invoke" \
  -H "Authorization: Bearer ${JWT_TOKEN}" \
  -H "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier: user@domain.com" \
  -H "Content-Type: application/json" \
  -d '{
    "tool": "list-mail-messages",
    "parameters": {
      "top": 10,
      "unread_only": true
    }
  }'
```

### 7.3 Expected Response

```json
{
  "messages": [
    {
      "id": "AAMkAGI...",
      "subject": "Test Email",
      "sender": {
        "emailAddress": {
          "name": "John Doe",
          "address": "john@example.com"
        }
      },
      "receivedDateTime": "2024-01-06T10:30:00Z",
      "isRead": false,
      "hasAttachments": false,
      "bodyPreview": "This is a preview..."
    }
  ],
  "count": 1
}
```

### 7.4 Test with Strands Agent

```python
from strands_agents import Agent
from strands_agents.tools import MCPTool

# Configure MCP tool with AgentCore Runtime
mcp_tool = MCPTool(
    name="ms365-email",
    runtime_arn=RUNTIME_ARN,
    headers={
        "Authorization": f"Bearer {jwt_token}",
        "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier": "user@domain.com"
    }
)

# Create agent
agent = Agent(
    name="Email Assistant",
    tools=[mcp_tool],
    model="anthropic.claude-3-5-sonnet-20241022-v2:0"
)

# Test
response = agent.run("List my unread emails")
print(response)
```

---

## Troubleshooting

### Issue: "Unauthorized" or 401 Error

**Possible Causes:**
1. JWT token is expired or invalid
2. JWT `aud` claim doesn't match `allowedAudiences`
3. JWT `azp` or `appid` claim not in `allowedClients`
4. JWT scopes don't match `allowedScopes`

**Solution:**
```bash
# Decode your JWT token to inspect claims
echo ${JWT_TOKEN} | cut -d'.' -f2 | base64 -d | jq .

# Verify:
# - "aud" matches one of your allowedAudiences
# - "azp" or "appid" matches one of your allowedClients
# - "scp" or "roles" contains required scopes
```

### Issue: "Authorization header not found"

**Possible Causes:**
1. `Authorization` not in `requestHeaderAllowlist`
2. Agent not sending `Authorization` header

**Solution:**
```bash
# Update runtime to include Authorization in allowlist
aws bedrock-agentcore update-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --request-header-configuration '{
    "requestHeaderAllowlist": ["Authorization"]
  }' \
  --region ${AWS_REGION}
```

### Issue: "Invalid discovery URL"

**Possible Causes:**
1. Wrong cloud endpoint (commercial vs government)
2. Tenant ID is incorrect

**Solution:**
```bash
# Test discovery URL manually
curl https://login.microsoftonline.us/${TENANT_ID}/v2.0/.well-known/openid-configuration

# Should return JSON with "issuer", "jwks_uri", etc.
# If 404, verify tenant ID and cloud type
```

### Issue: "S3 access denied"

**Possible Causes:**
1. Execution role doesn't have S3 read permissions
2. S3 bucket policy blocks access

**Solution:**
```bash
# Verify execution role has S3 policy
aws iam get-role-policy \
  --role-name MS365EmailMCPExecutionRole \
  --policy-name S3CodeAccess \
  --region ${AWS_REGION}

# Test S3 access
aws s3 cp ${S3_CODE_URI} /tmp/test.zip --region ${AWS_REGION}
```

### Issue: "Module not found" or Import Errors

**Possible Causes:**
1. Dependencies not included in zip package
2. Incorrect handler path

**Solution:**
```bash
# Verify package contents
unzip -l ms365-email-mcp-server.zip | grep server_token_only

# Ensure handler matches: "ms365_email_mcp_server.server_token_only:main"
# Directory structure should be:
# - ms365_email_mcp_server/
#   - __init__.py
#   - server_token_only.py
# - fastmcp/
# - httpx/
# - etc.
```

### Issue: "Graph API returns 401"

**Possible Causes:**
1. JWT token doesn't have Graph API permissions
2. Token is for wrong cloud (commercial vs government)
3. Token audience is wrong

**Solution:**
```bash
# Verify token audience and resource
echo ${JWT_TOKEN} | cut -d'.' -f2 | base64 -d | jq .aud

# For GCC, should be: "https://graph.microsoft.us"
# For commercial, should be: "https://graph.microsoft.com"

# Request new token with correct resource
az account get-access-token \
  --resource https://graph.microsoft.us \
  --query accessToken \
  --output tsv
```

### Issue: "agentcore: command not found"

**Possible Causes:**
1. AgentCore CLI not installed
2. Not in PATH

**Solution:**
```bash
# Install the toolkit
pip install bedrock-agentcore-starter-toolkit

# Verify installation
agentcore --version

# If using virtual environment, ensure it's activated
source venv/bin/activate  # or your venv path
```

### Issue: "ECR repository not found" (when using CLI)

**Possible Causes:**
1. CLI defaults to ECR deployment
2. ECR permissions missing

**Solution:**
```bash
# Option 1: Let CLI create ECR repository automatically
# (Press Enter when prompted for ECR repository)

# Option 2: Use manual S3 deployment (Option B in this guide)
# Switch to manual deployment approach

# Option 3: Ensure IAM role has ECR permissions
aws iam attach-role-policy \
  --role-name <your-execution-role> \
  --policy-arn arn:aws:iam::aws:policy/AmazonEC2ContainerRegistryFullAccess
```

### Issue: "agentcore.yaml not found"

**Possible Causes:**
1. `agentcore configure` not run
2. In wrong directory

**Solution:**
```bash
# Initialize configuration first
agentcore init --entry-point ms365_email_mcp_server/server_token_only.py --protocol MCP

# Or navigate to project root
cd /path/to/email-mcp-server
agentcore configure
```

### Issue: "Configuration validation failed"

**Possible Causes:**
1. Invalid YAML syntax
2. Missing required fields
3. Invalid OAuth discovery URL

**Solution:**
```bash
# Validate configuration
agentcore validate

# Check YAML syntax
yamllint agentcore.yaml  # if installed

# Test discovery URL manually
curl https://login.microsoftonline.us/<TENANT_ID>/v2.0/.well-known/openid-configuration
```

---

## Quick Reference: AgentCore CLI Commands

```bash
# Install
pip install bedrock-agentcore-starter-toolkit

# Initialize project
agentcore init --entry-point <entry_point> --protocol MCP

# Interactive configuration
agentcore configure

# Deploy
agentcore launch

# List runtimes
agentcore list

# Describe runtime
agentcore describe --runtime-name <name>

# Update runtime
agentcore update

# Delete runtime
agentcore delete --runtime-name <name>

# View logs
agentcore logs --runtime-name <name>

# Validate configuration
agentcore validate

# Help
agentcore --help
agentcore <command> --help
```

---

## Additional Resources

### AWS AgentCore Documentation
- [AWS AgentCore Runtime API Reference](https://docs.aws.amazon.com/bedrock-agentcore-control/latest/APIReference/Welcome.html)
- [AWS AgentCore JWT Authorizer Configuration](https://docs.aws.amazon.com/bedrock-agentcore-control/latest/APIReference/API_AuthorizerConfiguration.html)
- [AWS Request Header Configuration](https://docs.aws.amazon.com/bedrock-agentcore-control/latest/APIReference/API_RequestHeaderConfiguration.html)
- [AgentCore Starter Toolkit Documentation](https://aws.github.io/bedrock-agentcore-starter-toolkit/)
- [AgentCore CLI Quickstart Guide](https://aws.github.io/bedrock-agentcore-starter-toolkit/user-guide/runtime/quickstart.html)
- [AgentCore CLI GitHub Repository](https://github.com/aws/bedrock-agentcore-starter-toolkit)

### Microsoft Documentation
- [Microsoft Entra ID US Government Cloud](https://learn.microsoft.com/en-us/azure/active-directory/develop/authentication-national-cloud)
- [Microsoft Graph API for US Government](https://learn.microsoft.com/en-us/graph/deployments#microsoft-graph-and-graph-explorer-service-root-endpoints)
- [MSAL Python Documentation](https://msal-python.readthedocs.io/)
- [OAuth 2.0 Client Credentials Flow](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow)

---

## Summary

You've successfully deployed the MS365 Email MCP Server as an AWS AgentCore Runtime with:

✅ **S3-based code deployment** (no ECR required)  
✅ **OAuth 2.0 JWT authentication** using Microsoft Entra ID  
✅ **US Government Cloud support** (`login.microsoftonline.us`, `graph.microsoft.us`)  
✅ **Authorization header pass-through** for Graph API access  
✅ **Custom header support** for shared mailbox access  
✅ **Multi-agent support** via `allowedClients` configuration  

Your agents can now securely access Microsoft 365 email operations using their own JWT tokens, with no client secrets passed in headers.

