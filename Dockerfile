# Microsoft Teams MCP Server Dockerfile
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    git \
    curl \
    && rm -rf /var/lib/apt/lists/*

# Install uv package manager
RUN pip install uv

# Copy project files
COPY pyproject.toml uv.lock ./
COPY . .

# Create virtual environment and install dependencies
RUN uv sync --frozen

# Create credentials directory
RUN mkdir -p /app/.microsoft_teams_mcp/credentials

# Set environment variables
ENV PYTHONPATH=/app
ENV PYTHONUNBUFFERED=1
ENV PORT=8003
ENV TEAMS_MCP_PORT=8003
ENV TEAMS_MCP_BASE_URI=http://localhost
ENV MCP_ENABLE_OAUTH21=true
ENV LOG_LEVEL=INFO

# Expose port
EXPOSE 8003

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
  CMD curl -f http://localhost:${PORT}/health || exit 1

# Default command (can be overridden)
CMD ["uv", "run", "main.py", "--transport", "streamable-http", "--port", "8003"]
