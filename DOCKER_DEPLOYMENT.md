# Microsoft Teams MCP Server - Docker Deployment

This guide explains how to deploy the Microsoft Teams MCP Server using Docker.

## Prerequisites

1. **Docker & Docker Compose** installed
2. **Microsoft Azure App Registration** with required permissions
3. **Environment variables** configured

## Quick Start

### 1. Clone and Setup

```bash
git clone <repository-url>
cd teams-mcp-server-python
```

### 2. Configure Environment

```bash
# Copy the example environment file
cp .env.example .env

# Edit .env with your Microsoft OAuth credentials
nano .env
```

Required environment variables:
- `MICROSOFT_OAUTH_CLIENT_ID`: Your Azure App Client ID
- `MICROSOFT_OAUTH_CLIENT_SECRET`: Your Azure App Client Secret
- `MICROSOFT_TENANT_ID`: Your tenant ID (or "common" for multi-tenant)

### 3. Deploy with Docker Compose

```bash
# Build and start the service
docker-compose up -d

# View logs
docker-compose logs -f teams-mcp-server

# Stop the service
docker-compose down
```

### 4. Deploy with Docker

```bash
# Build the image
docker build -t teams-mcp-server .

# Run the container
docker run -d \
  --name teams-mcp-server \
  -p 8000:8000 \
  -e MICROSOFT_OAUTH_CLIENT_ID=your-client-id \
  -e MICROSOFT_OAUTH_CLIENT_SECRET=your-client-secret \
  -e MICROSOFT_TENANT_ID=your-tenant-id \
  -v teams_mcp_credentials:/app/.microsoft_teams_mcp/credentials \
  teams-mcp-server
```

## Configuration

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `MICROSOFT_OAUTH_CLIENT_ID` | Azure App Client ID | Required |
| `MICROSOFT_OAUTH_CLIENT_SECRET` | Azure App Client Secret | Required |
| `MICROSOFT_TENANT_ID` | Tenant ID or "common" | `common` |
| `PORT` | Server port | `8000` |
| `TEAMS_MCP_BASE_URI` | Base URI for callbacks | `http://localhost` |
| `MCP_ENABLE_OAUTH21` | Enable OAuth 2.1 | `true` |
| `MCP_SINGLE_USER_MODE` | Single user mode | `false` |
| `LOG_LEVEL` | Logging level | `INFO` |

### Custom Port

To run on a different port:

```bash
# Docker Compose
PORT=3000 docker-compose up -d

# Docker
docker run -d \
  --name teams-mcp-server \
  -p 3000:8000 \
  -e PORT=8000 \
  teams-mcp-server uv run main.py --transport streamable-http --port 8000
```

### Production Deployment

For production, consider:

1. **Use HTTPS**: Set `TEAMS_MCP_BASE_URI=https://yourdomain.com`
2. **Secure secrets**: Use Docker secrets or external secret management
3. **Persistent storage**: Mount credentials volume to persistent storage
4. **Resource limits**: Set memory and CPU limits
5. **Health checks**: Monitor the `/health` endpoint

Example production docker-compose.yml:

```yaml
version: '3.8'

services:
  teams-mcp-server:
    build: .
    ports:
      - "8000:8000"
    environment:
      - TEAMS_MCP_BASE_URI=https://yourdomain.com
      - MICROSOFT_OAUTH_REDIRECT_URI=https://yourdomain.com/callback
      - OAUTH2_ALLOW_INSECURE_TRANSPORT=false
    volumes:
      - /path/to/persistent/credentials:/app/.microsoft_teams_mcp/credentials
    deploy:
      resources:
        limits:
          memory: 512M
          cpus: '0.5'
    restart: unless-stopped
```

## Azure App Registration

Register your app in Azure with these settings:

1. **Redirect URIs**: 
   - `http://localhost:8000/callback` (development)
   - `https://yourdomain.com/callback` (production)

2. **API Permissions**:
   - `User.Read`
   - `Team.ReadBasic.All`
   - `Channel.ReadBasic.All`
   - `ChannelMessage.Read.All`
   - `Chat.Read`
   - `TeamMember.Read.All`

## Health Check

The server provides a health check endpoint at `/health`:

```bash
curl http://localhost:8000/health
```

## Troubleshooting

### Common Issues

1. **Authentication Failed**: Check OAuth credentials and redirect URI
2. **Port Conflicts**: Change the host port in docker-compose.yml
3. **Permission Denied**: Verify Azure app permissions and admin consent

### Logs

View detailed logs:

```bash
# Docker Compose
docker-compose logs -f teams-mcp-server

# Docker
docker logs -f teams-mcp-server
```

### Debug Mode

Enable debug logging:

```bash
# Add to .env
LOG_LEVEL=DEBUG
```

## Support

For issues and questions, please check:
- Server logs for error details
- Azure app registration configuration
- Network connectivity and firewall settings
