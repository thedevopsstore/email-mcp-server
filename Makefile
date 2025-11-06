.PHONY: install dev-install run test clean format lint

# Install dependencies
install:
	uv sync

# Install with dev dependencies
dev-install:
	uv sync --dev

# Run the server
run:
	uv run ms365-email-mcp-server

# Run tests (if you add tests)
test:
	uv run pytest

# Format code
format:
	uv run black ms365_email_mcp_server
	uv run ruff format ms365_email_mcp_server

# Lint code
lint:
	uv run ruff check ms365_email_mcp_server
	uv run black --check ms365_email_mcp_server

# Clean build artifacts
clean:
	rm -rf dist/
	rm -rf build/
	rm -rf *.egg-info
	rm -rf .uv/
	find . -type d -name __pycache__ -exec rm -r {} +
	find . -type f -name "*.pyc" -delete

