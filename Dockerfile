FROM python:3.11-slim

# Install UV
COPY --from=ghcr.io/astral-sh/uv:latest /uv /usr/local/bin/uv

WORKDIR /app

# Copy dependency files
COPY pyproject.toml ./

# Install dependencies using UV
RUN uv pip install --system -e .

# Copy application code
COPY ms365_email_mcp_server ./ms365_email_mcp_server

# Expose port
EXPOSE 8100

# Run the server using the entry point
CMD ["ms365-email-mcp-server"]
