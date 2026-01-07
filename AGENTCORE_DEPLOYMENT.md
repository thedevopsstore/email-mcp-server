# AWS AgentCore Runtime Deployment Guide

This guide walks through deploying the MS365 Email MCP Server as an AWS AgentCore Runtime using **S3-based code upload** with **IAM (SigV4) inbound authentication**, while passing the **Microsoft Graph access token** via AgentCore Runtime custom headers.

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
7. [Step 4: Configure Request Header Allowlist](#step-4-configure-request-header-allowlist)
8. [Step 5: Create AgentCore Runtime (IAM Auth)](#step-5-create-agentcore-runtime-iam-auth)
9. [Step 6: Test Your Deployment](#step-6-test-your-deployment)
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

**⚠️ Important Note:** The AgentCore CLI commands and structure are evolving. The commands shown below are examples based on typical CLI patterns. **Always verify the actual commands** by running `agentcore --help` after installation, and refer to the [official documentation](https://aws.github.io/bedrock-agentcore-starter-toolkit/) for the most current commands.

#### A.1 Install the AgentCore CLI

```bash
pip install bedrock-agentcore-starter-toolkit

# Verify installation and see available commands
agentcore --help
```

**Note:** The AgentCore CLI typically uses **Amazon ECR** for containerized deployments. If you specifically need S3-based deployment, use **Option B** below.

#### A.2 Create Configuration File

Navigate to your project directory and create an `agentcore.yaml` configuration file manually, or use the CLI's interactive configuration:

```bash
cd /path/to/email-mcp-server

# Option 1: Create agentcore.yaml manually (see A.4 for template)
# Option 2: Use CLI configure command (if available)
agentcore configure
```

**Note:** The exact CLI commands may vary. Check available commands with:
```bash
agentcore --help
# or
pip show bedrock-agentcore-starter-toolkit
```

#### A.3 Configure Interactively (if supported)

If the CLI supports interactive configuration, run:

```bash
agentcore configure
```

**Note:** If `agentcore configure` is not available, create the `agentcore.yaml` file manually (see A.4 below).

If the interactive wizard is available, it will prompt you for:

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

3. **Inbound Authentication (IAM)**:
   - No OAuth/JWT authorizer is configured in this model.
   - Calls to the runtime are authorized using AWS IAM (SigV4).

4. **Environment Variables**:
   ```
   Enter environment variables (key=value, comma-separated):
   MS365_CLOUD_TYPE=gov,LOG_LEVEL=INFO,HOST=0.0.0.0,PORT=8000,STATELESS_HTTP=true
   ```
   
   **Important:** AgentCore Runtime requires MCP servers to listen on **port 8000** internally. The external port is managed by AgentCore Runtime itself.

5. **Request Header Allowlist**:
   ```
   Enter request headers to allowlist (comma-separated):
   X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization,X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier
   ```

6. **Cloud Configuration** (if prompted):
   ```
   Select cloud type [commercial/gov]: gov
   ```

#### A.4 Create/Review Configuration File

Create or review the `agentcore.yaml` configuration file:

If the CLI created the file automatically, review it:

```bash
cat agentcore.yaml
```

**Or create it manually** with the following template:

```yaml
runtime:
  name: ms365-email-mcp-runtime
  protocol: MCP
  entryPoint: ms365_email_mcp_server/server_token_only.py
  runtime: python3.11

# Inbound auth is IAM (SigV4) in this model (no JWT authorizer configured)

requestHeaders:
  allowlist:
    - X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization
    - X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier

environment:
  MS365_CLOUD_TYPE: gov
  LOG_LEVEL: INFO
  HOST: 0.0.0.0
  PORT: "8000"  # Required: AgentCore Runtime expects port 8000
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

## Port Configuration for AgentCore Runtime

### ⚠️ Important: Port Requirements

When deploying to **AgentCore Runtime**, your MCP server **must** be configured as follows:

| Setting | Value | Reason |
|---------|-------|--------|
| **Host** | `0.0.0.0` | Required - server must listen on all interfaces |
| **Port** | `8000` | **Required** - AgentCore Runtime expects MCP servers on port 8000 |
| **External Port** | Managed by AgentCore | Clients connect to the runtime endpoint, not the container port |

### How Ports Work in AgentCore Runtime

```
┌─────────────────────────────────────────────────────────┐
│                    Client/Agent                         │
│  Connects to: https://runtime.bedrock-agentcore...     │
│  (AgentCore-managed endpoint, NOT port 8000 directly)  │
└────────────────────────┬────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────┐
│            AgentCore Runtime Load Balancer              │
│  Routes to container's /mcp endpoint on port 8000      │
└────────────────────────┬────────────────────────────────┘
                         │
                         ▼
┌─────────────────────────────────────────────────────────┐
│         MCP Server Container (Your Code)                │
│  Must listen on: 0.0.0.0:8000                          │
│  FastMCP handles /mcp endpoint automatically           │
└─────────────────────────────────────────────────────────┘
```

**Key Points:**
- **Internal Port**: Your MCP server **must** listen on port `8000` inside the container
- **External Port**: AgentCore Runtime manages the external endpoint; clients never connect directly to port 8000
- **Port Mapping**: There's no port mapping like Docker (`-p 8000:8000`) - AgentCore handles routing automatically
- **Local Development**: For local testing (not AgentCore), you can use any port (e.g., 8100)

### Setting the Port for AgentCore

Set the port via environment variable in your runtime configuration:

```bash
# In agentcore.yaml or runtime config
environment:
  PORT: "8000"  # Required for AgentCore Runtime
  HOST: "0.0.0.0"  # Required for AgentCore Runtime
```

---

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                         Agent/Client                            │
│  (Strands Agent, MCP Client with AWS creds + Graph token)      │
│  Connects to: https://runtime.bedrock-agentcore...             │
└────────────────────────┬────────────────────────────────────────┘
                         │ X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization: Bearer <graph_token>
                         │ X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│              AWS AgentCore Runtime (IAM SigV4)                  │
│  - Authorizes inbound via IAM                                  │
│  - Routes to container on port 8000 (/mcp endpoint)            │
│  - Passes allowlisted custom headers to MCP server             │
└────────────────────────┬────────────────────────────────────────┘
                         │ Internal: 0.0.0.0:8000/mcp
                         │ X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization: Bearer <graph_token>
                         │ X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│           MS365 Email MCP Server (server_token_only.py)         │
│  - Listens on 0.0.0.0:8000 (required)                          │
│  - Extracts token from X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization header
│  - Extracts user_identifier from custom header                  │
│  - Uses token directly for Microsoft Graph API calls            │
└────────────────────────┬────────────────────────────────────────┘
                         │ Authorization: Bearer <graph_token>
                         ▼
┌─────────────────────────────────────────────────────────────────┐
│              Microsoft Graph API (graph.microsoft.us)           │
│                    US Government Cloud (GCC)                    │
└─────────────────────────────────────────────────────────────────┘
```

**Key Points:**
- Inbound access to AgentCore Runtime uses **IAM (SigV4)** (no OAuth/JWT authorizer in this model)
- The Microsoft Graph token is passed through to the MCP server via `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization`
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

## Step 4: Configure Request Header Allowlist

The `requestHeaderAllowlist` is **critical** for the token-only Graph flow. It tells AgentCore Runtime which HTTP headers to pass through to your MCP server.

Allowlist these headers:

- `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization` (Graph access token)
- `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier` (shared mailbox selector, optional)

Reference: [RequestHeaderConfiguration API Documentation](https://docs.aws.amazon.com/bedrock-agentcore-control/latest/APIReference/API_RequestHeaderConfiguration.html)

---

## Step 5: Create AgentCore Runtime (IAM Auth)

Inbound authentication is IAM (SigV4) in this model—no JWT authorizer configuration is required.

### 5.1 Create Runtime Configuration

Create a file `agentcore-runtime-config.json`:

```json
{
  "name": "ms365-email-mcp-runtime",
  "description": "MS365 Email MCP Server (IAM inbound auth) with Graph token passed via AgentCore custom header",
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
  },
  "requestHeaderConfiguration": {
    "requestHeaderAllowlist": [
      "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization",
      "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier"
    ]
  },
  "executionRoleArn": "arn:aws:iam::123456789012:role/MS365EmailMCPExecutionRole"
}
```

**Replace placeholders:**
- `s3Uri`: Your S3 URI from Step 2.3
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
  "PORT": "8000",
  "STATELESS_HTTP": "true"
}
```

#### Request Header Allowlist (Critical!)
```json
"requestHeaderAllowlist": [
  "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization",
  "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier"
]
```

**Why this matters:**
- `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization`: Passes the Microsoft Graph access token through to the MCP server for Graph API calls
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

## Appendix: Header Allowlist Deep Dive

The `requestHeaderAllowlist` is **critical** for the token-only authentication flow. It tells AgentCore Runtime which HTTP headers to pass through to your MCP server.

### 6.1 Understanding Header Pass-Through

```
Agent Request Headers:
├── X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization: Bearer <graph_token>
├── X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier: user@domain.com
└── Other headers...

                    ↓ (AgentCore authorizes inbound via IAM; passes allowlisted headers)

AgentCore passes through (based on allowlist):
├── X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization: Bearer <graph_token>  ✅ (in allowlist)
├── X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier: user@domain.com  ✅ (in allowlist)
└── Other headers... ❌ (not in allowlist, dropped)

                    ↓

MCP Server receives:
├── X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization: Bearer <graph_token>
└── X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier: user@domain.com
```

### 6.2 Header Configuration in Runtime

The configuration from Step 5.1 includes:

```json
"requestHeaderConfiguration": {
  "requestHeaderAllowlist": [
    "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization",
    "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier"
  ]
}
```

**Header Descriptions:**

| Header | Purpose | Required |
|--------|---------|----------|
| `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization` | Contains `Bearer <graph_token>` for Microsoft Graph API access | ✅ Yes |
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

## Step 6: Test Your Deployment

### 7.1 Obtain a JWT Token

Use MSAL or Azure CLI to get a token:

```bash
# Using Azure CLI (for testing)
az login --use-device-code --cloud AzureUSGovernment  # For GCC
# or
az login --use-device-code  # For commercial

# Get token for Microsoft Graph
export GRAPH_TOKEN=$(az account get-access-token \
  --resource https://graph.microsoft.us \
  --query accessToken \
  --output tsv)

echo "Graph token (first 50 chars): ${GRAPH_TOKEN:0:50}..."
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
graph_token = result["access_token"]
```

### 7.2 Test MCP Server Directly

```bash
# Get runtime endpoint
export RUNTIME_ENDPOINT=$(aws bedrock-agentcore get-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --query 'runtime.endpoint' \
  --output text \
  --region ${AWS_REGION})

# Inbound auth is IAM (SigV4), so you must use an AWS SDK or a SigV4-capable HTTP client.
# The important part is the headers you pass through to the MCP server:
#
# - X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization: Bearer <graph_token>
# - X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier: <optional>
#
# Example payload:
# {
#   "tool": "list-mail-messages",
#   "parameters": {
#     "top": 10,
#     "unread_only": true
#   }
# }
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
        "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization": f"Bearer {graph_token}",
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

### Issue: "AccessDenied" / 403 from AgentCore Runtime

**Possible Causes:**
1. Caller doesn't have AWS credentials (or is using the wrong profile/role)
2. IAM policy doesn't allow invoking the runtime

**Solution:**
- Verify the AWS identity you are using (role/user/profile)
- Update IAM policy to allow invoking the AgentCore runtime

### Issue: "access_token is required" (from MCP server)

**Possible Causes:**
1. `X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization` header not allowlisted
2. Caller didn't send the header

**Solution:**
```bash
# Update runtime to include the custom header(s) in the allowlist
aws bedrock-agentcore update-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --request-header-configuration '{
    "requestHeaderAllowlist": [
      "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization",
      "X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-UserIdentifier"
    ]
  }' \
  --region ${AWS_REGION}
```

### Issue: "Graph API returns 401"

**Possible Causes:**
1. Graph token expired/invalid
2. Token minted for the wrong Graph cloud (`graph.microsoft.com` vs `graph.microsoft.us`)
3. Missing Graph permissions (`Mail.Read`, `Mail.Send`, etc.)

**Solution:**
- Re-mint the Graph token from the correct authority and resource
- Ensure correct scopes/permissions are granted and consented

### Issue: "Token acquisition fails" (Graph token)

**Possible Causes:**
1. Wrong authority for your cloud (commercial vs gov)
2. Wrong Graph resource (`graph.microsoft.com` vs `graph.microsoft.us`)
3. Missing admin consent / permissions for Graph

**Solution:**
- For US Gov cloud, use authority base `https://login.microsoftonline.us` (and Graph resource `https://graph.microsoft.us`)
- Re-check the app registration permissions and admin consent

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
1. Graph token doesn't have Graph API permissions
2. Token is for wrong cloud (commercial vs government)
3. Token audience is wrong

**Solution:**
```bash
# Verify token audience and resource
echo ${GRAPH_TOKEN} | cut -d'.' -f2 | base64 -d | jq .aud

# For GCC, should be: "https://graph.microsoft.us"
# For commercial, should be: "https://graph.microsoft.com"

# Request new token with correct resource
az account get-access-token \
  --resource https://graph.microsoft.us \
  --query accessToken \
  --output tsv
```

### Issue: "Connection refused" or "Cannot connect to MCP server"

**Possible Causes:**
1. Wrong port configured (not 8000)
2. Server not listening on 0.0.0.0
3. Container not starting properly

**Solution:**
```bash
# Verify runtime environment variables include PORT=8000
aws bedrock-agentcore get-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --query 'runtime.runtimeConfig.mcpRuntimeConfig.serverConfig.environmentVariables' \
  --region ${AWS_REGION}

# Should show:
# {
#   "PORT": "8000",
#   "HOST": "0.0.0.0",
#   ...
# }

# Check runtime logs for port binding errors
agentcore logs --runtime-name ms365-email-mcp-runtime

# Look for messages like:
# "Listening on 0.0.0.0:8000"
# NOT "Listening on 0.0.0.0:8100" or "127.0.0.1:8000"
```

### Issue: "Port already in use" or "Address already in use"

**Possible Causes:**
1. Multiple instances trying to bind to port 8000
2. Port conflict in container

**Solution:**
```bash
# This shouldn't happen in AgentCore Runtime (each container is isolated)
# If you see this locally, ensure you're using the correct port:

# For local development (NOT AgentCore):
PORT=8100 python -m ms365_email_mcp_server.server_token_only

# For AgentCore Runtime:
PORT=8000  # Must be 8000
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
1. Configuration file not created
2. In wrong directory
3. CLI doesn't auto-create the file

**Solution:**
```bash
# Option 1: Create agentcore.yaml manually (see template in A.4)
# Option 2: Navigate to project root and use CLI (if available)
cd /path/to/email-mcp-server
agentcore configure  # If this command exists

# Option 3: Check what commands are available
agentcore --help
```

### Issue: "Configuration validation failed"

**Possible Causes:**
1. Invalid YAML syntax
2. Missing required fields
3. Missing/invalid required fields

**Solution:**
```bash
# Validate configuration
agentcore validate

# Check YAML syntax
yamllint agentcore.yaml  # if installed

# If you are not using an OAuth/JWT authorizer for inbound auth (IAM model), there is no discovery URL to validate.
```

---

## Quick Reference: AgentCore CLI Commands

**Note:** The exact CLI commands may vary. Always verify with `agentcore --help` after installation.

```bash
# Install
pip install bedrock-agentcore-starter-toolkit

# Check available commands
agentcore --help

# Common commands (verify these exist in your version):
# - agentcore configure (if available)
# - agentcore create-runtime (or similar)
# - agentcore deploy (or agentcore launch)
# - agentcore list-runtimes (or agentcore list)
# - agentcore describe-runtime
# - agentcore update-runtime
# - agentcore delete-runtime

# View logs (if available)
agentcore logs --runtime-name <name>

# Validate configuration (if available)
agentcore validate
```

**Important:** The AgentCore CLI commands are evolving. For the most up-to-date commands, refer to:
- [AgentCore Starter Toolkit Documentation](https://aws.github.io/bedrock-agentcore-starter-toolkit/)
- [AWS AgentCore Developer Guide](https://docs.aws.amazon.com/bedrock-agentcore/latest/devguide/)

---

## Updating Your Deployment

After your initial deployment, see **[UPDATING_DEPLOYMENT.md](./UPDATING_DEPLOYMENT.md)** for detailed instructions on:
- Deploying new code versions
- Updating configuration (environment variables, headers, OAuth)
- Version management and rollback procedures
- Best practices for updates

**Quick Update Commands:**

```bash
# Using AgentCore CLI
agentcore update

# Using AWS CLI (S3-based)
aws bedrock-agentcore update-runtime \
  --runtime-arn ${RUNTIME_ARN} \
  --cli-input-json file://update-config.json
```

---

## Additional Resources

### AWS AgentCore Documentation
- [AWS AgentCore Runtime API Reference](https://docs.aws.amazon.com/bedrock-agentcore-control/latest/APIReference/Welcome.html)
- [AWS Request Header Configuration](https://docs.aws.amazon.com/bedrock-agentcore-control/latest/APIReference/API_RequestHeaderConfiguration.html)
- [AgentCore Runtime Versioning](https://docs.aws.amazon.com/bedrock-agentcore/latest/devguide/agent-runtime-versioning.html)
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
✅ **IAM (SigV4) inbound authentication** to AgentCore Runtime  
✅ **US Government Cloud support** (`login.microsoftonline.us`, `graph.microsoft.us`)  
✅ **Custom header pass-through** for Graph API access (`X-Amzn-Bedrock-AgentCore-Runtime-Custom-Ms365-Authorization`)  
✅ **Custom header support** for shared mailbox access  
✅ **Multi-agent support** via IAM policies + per-agent Graph tokens
✅ **Correct port configuration**: Server listening on `0.0.0.0:8000` (required by AgentCore Runtime)

### Key Port Configuration Notes:

- **Internal Port**: Your MCP server **must** listen on port `8000` inside the container
- **Host**: Server **must** bind to `0.0.0.0` (not `127.0.0.1` or `localhost`)
- **External Port**: AgentCore Runtime manages the external endpoint automatically
- **No Port Mapping Needed**: AgentCore routes requests to your container's port 8000 internally

Your agents can now securely access Microsoft 365 email operations using their own JWT tokens, with no client secrets passed in headers. Agents connect to the AgentCore Runtime endpoint (not directly to port 8000), and the runtime routes requests to your MCP server.

