#!/bin/bash
# Development setup script for Teams MCP Server

set -e

echo "🚀 Setting up Teams MCP Server development environment..."

# Check if uv is installed
if ! command -v uv &> /dev/null; then
    echo "❌ uv is not installed. Please install it first:"
    echo "   curl -LsSf https://astral.sh/uv/install.sh | sh"
    exit 1
fi

# Create virtual environment
echo "📦 Creating virtual environment..."
uv venv

# Activate virtual environment
echo "🔧 Activating virtual environment..."
source .venv/bin/activate

# Install dependencies
echo "📥 Installing dependencies..."
uv pip install -e ".[dev]"

# Sync dependencies
echo "🔄 Syncing dependencies..."
uv sync

# Copy environment file if it doesn't exist
if [ ! -f .env ]; then
    echo "📝 Creating .env file from template..."
    cp .env.oauth21 .env
    echo "⚠️  Please edit .env file with your Microsoft OAuth credentials"
fi

# Create logs directory
mkdir -p logs

echo "✅ Development environment setup complete!"
echo ""
echo "Next steps:"
echo "1. Activate the virtual environment: source .venv/bin/activate"
echo "2. Edit .env file with your Microsoft OAuth credentials"
echo "3. Run the server: make run-http"
echo ""
echo "Available make commands:"
echo "  make help        - Show available commands"
echo "  make run-http    - Run HTTP server"
echo "  make test        - Run tests"
echo "  make format      - Format code"
echo "  make env-check   - Check environment variables"
