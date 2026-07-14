# Pipedrive MCP Server

This is a Model Context Protocol (MCP) server that connects to the Pipedrive API v2. It allows you to expose Pipedrive data and functionality to LLM applications like Claude.

## Features

- Read and write access to Pipedrive data
- Exposes deals, persons, organizations, leads, activities, and notes
- Includes all fields including custom fields
- Create and update records in Pipedrive
- Predefined prompts for common operations
- Docker support with multi-stage builds
- JWT authentication support
- Built-in rate limiting for API requests
- Advanced deal filtering (by owner, status, date range, value, etc.)

## Setup

### Standard Setup

1. Clone this repository
2. Install dependencies:
   ```
   npm install
   ```
3. Create a `.env` file in the root directory with your configuration:
   ```
   PIPEDRIVE_API_TOKEN=your_api_token_here
   PIPEDRIVE_DOMAIN=your-company.pipedrive.com
   ```
4. Build the project:
   ```
   npm run build
   ```
5. Start the server:
   ```
   npm start
   ```

### Docker Setup

#### Option 1: Using Docker Compose (standalone)

1. Copy `.env.example` to `.env` and configure your settings:
   ```bash
   PIPEDRIVE_API_TOKEN=your_api_token_here
   PIPEDRIVE_DOMAIN=your-company.pipedrive.com
   MCP_TRANSPORT=sse  # Use SSE transport for Docker
   MCP_PORT=3000
   ```
2. Build and run with Docker Compose:
   ```bash
   docker-compose up -d
   ```
3. The server will be available at `http://localhost:3000`
   - SSE endpoint: `http://localhost:3000/sse`
   - Health check: `http://localhost:3000/health`

#### Option 2: Using Pre-built Docker Image

Pull and run the pre-built image from GitHub Container Registry:

**For SSE transport (HTTP access):**
```bash
docker run -d \
  -p 3000:3000 \
  -e PIPEDRIVE_API_TOKEN=your_api_token_here \
  -e PIPEDRIVE_DOMAIN=your-company.pipedrive.com \
  -e MCP_TRANSPORT=sse \
  -e MCP_PORT=3000 \
  ghcr.io/juhokoskela/pipedrive-mcp-server:main
```

**For stdio transport (local use):**
```bash
docker run -i \
  -e PIPEDRIVE_API_TOKEN=your_api_token_here \
  -e PIPEDRIVE_DOMAIN=your-company.pipedrive.com \
  ghcr.io/juhokoskela/pipedrive-mcp-server:main
```

#### Option 3: Integrating into Existing Project

Add the MCP server to your existing application's `docker-compose.yml`:

```yaml
services:
  # Your existing services...

  pipedrive-mcp-server:
    image: ghcr.io/juhokoskela/pipedrive-mcp-server:main
    container_name: pipedrive-mcp-server
    restart: unless-stopped
    ports:
      - "3000:3000"
    environment:
      - PIPEDRIVE_API_TOKEN=${PIPEDRIVE_API_TOKEN}
      - PIPEDRIVE_DOMAIN=${PIPEDRIVE_DOMAIN}
      - MCP_TRANSPORT=sse
      - MCP_PORT=3000
      - PIPEDRIVE_RATE_LIMIT_MIN_TIME_MS=${PIPEDRIVE_RATE_LIMIT_MIN_TIME_MS:-250}
      - PIPEDRIVE_RATE_LIMIT_MAX_CONCURRENT=${PIPEDRIVE_RATE_LIMIT_MAX_CONCURRENT:-2}
    healthcheck:
      test: ["CMD", "wget", "--no-verbose", "--tries=1", "--spider", "http://localhost:3000/health", "||", "exit", "1"]
      interval: 30s
      timeout: 10s
      retries: 3
      start_period: 10s
    logging:
      driver: "json-file"
      options:
        max-size: "10m"
        max-file: "3"
```

Then add the required environment variables to your `.env` file.

### Environment Variables

Required:
- `PIPEDRIVE_API_TOKEN` - Your Pipedrive API token
- `PIPEDRIVE_DOMAIN` - Your Pipedrive domain (e.g., `your-company.pipedrive.com`)

Optional (JWT Authentication):
- `MCP_JWT_SECRET` - JWT secret for authentication
- `MCP_JWT_TOKEN` - JWT token for authentication
- `MCP_JWT_ALGORITHM` - JWT algorithm (default: HS256)
- `MCP_JWT_AUDIENCE` - JWT audience
- `MCP_JWT_ISSUER` - JWT issuer

When JWT authentication is enabled, all SSE requests (`/sse` and the message endpoint) must include an `Authorization: Bearer <token>` header signed with the configured secret.

Optional (Rate Limiting):
- `PIPEDRIVE_RATE_LIMIT_MIN_TIME_MS` - Minimum time between requests in milliseconds (default: 250)
- `PIPEDRIVE_RATE_LIMIT_MAX_CONCURRENT` - Maximum concurrent requests (default: 2)

Optional (Transport Configuration):
- `MCP_TRANSPORT` - Transport type: `stdio` (default, for local use) or `sse` (for Docker/HTTP access)
- `MCP_PORT` - Port for SSE transport (default: 3000, only used when `MCP_TRANSPORT=sse`)
- `MCP_ENDPOINT` - Message endpoint path for SSE (default: /message, only used when `MCP_TRANSPORT=sse`)

## Using with Claude

To use this server with Claude for Desktop:

1. Configure Claude for Desktop by editing your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "pipedrive": {
      "command": "node",
      "args": ["/path/to/pipedrive-mcp-server/build/index.js"],
      "env": {
        "PIPEDRIVE_API_TOKEN": "your_api_token_here",
        "PIPEDRIVE_DOMAIN": "your-company.pipedrive.com"
      }
    }
  }
}
```

2. Restart Claude for Desktop
3. In the Claude application, you should now see the Pipedrive tools available

## Available Tools

### Read Tools

- `get-users`: Get all users/owners from Pipedrive to identify owner IDs for filtering
- `get-deals`: Get deals with flexible filtering options (search by title, date range, owner, stage, status, value range, etc.)
- `get-deal`: Get a specific deal by ID (including custom fields)
- `get-deal-notes`: Get detailed notes and custom booking details for a specific deal
- `search-deals`: Search deals by term
- `get-persons`: Get all persons from Pipedrive (including custom fields)
- `get-person`: Get a specific person by ID (including custom fields)
- `search-persons`: Search persons by term
- `get-organizations`: Get all organizations from Pipedrive (including custom fields)
- `get-organization`: Get a specific organization by ID (including custom fields)
- `search-organizations`: Search organizations by term
- `get-pipelines`: Get all pipelines from Pipedrive
- `get-pipeline`: Get a specific pipeline by ID
- `get-stages`: Get all stages from all pipelines
- `search-leads`: Search leads by term
- `search-all`: Search across all item types (deals, persons, organizations, etc.)

### Write Tools

- `add-deal`: Create a new deal in Pipedrive with customizable fields (title, value, status, linked contacts, owner, label, etc.)
- `update-deal`: Update an existing deal with new information
- `delete-deal`: Delete a deal
- `merge-deals`: Merge a deal into another deal
- `add-deal-follower` / `remove-deal-follower`: Manage deal followers
- `add-deal-product` / `update-deal-product` / `delete-deal-product`: Manage products (line items) attached to a deal
- `add-person`: Create a new person/contact in Pipedrive with email, phone, organization, owner, and label
- `update-person`: Update an existing person's information
- `delete-person`: Delete a person
- `merge-persons`: Merge a person into another person
- `add-person-follower` / `remove-person-follower`: Manage person followers
- `add-organization`: Create a new organization in Pipedrive
- `update-organization`: Update an existing organization's information
- `delete-organization`: Delete an organization
- `merge-organizations`: Merge an organization into another organization
- `add-organization-follower` / `remove-organization-follower`: Manage organization followers
- `add-note`: Add a note to a deal, person, organization, or lead
- `update-note`: Update an existing note
- `delete-note`: Delete a note
- `add-note-comment` / `get-note-comments` / `update-note-comment` / `delete-note-comment`: Manage comments on a note
- `add-activity`: Create a new activity/task in Pipedrive
- `update-activity`: Update an existing activity
- `delete-activity`: Delete an activity
- `add-lead`: Create a new lead in Pipedrive (to attach a note to a new lead, call `add-note` afterwards with the returned lead ID)
- `update-lead`: Update an existing lead
- `delete-lead`: Delete a lead
- `get-products` / `get-product`: List or fetch products
- `add-product` / `update-product` / `delete-product`: Manage products
- `upload-file`: Upload a file and associate it with a deal, person, organization, product, activity, or lead
- `link-file-to-item`: Link a remote (Google Drive) file to a deal, person, or organization
- `delete-file`: Delete a file
- `get-webhooks` / `add-webhook` / `delete-webhook`: Manage webhooks
- `add-pipeline` / `update-pipeline` / `delete-pipeline`: Manage pipelines
- `add-stage` / `update-stage` / `delete-stage`: Manage pipeline stages

## Available Prompts

- `list-all-deals`: List all deals in Pipedrive
- `list-all-persons`: List all persons in Pipedrive
- `list-all-pipelines`: List all pipelines in Pipedrive
- `analyze-deals`: Analyze deals by stage
- `analyze-contacts`: Analyze contacts by organization
- `analyze-leads`: Analyze leads by status
- `compare-pipelines`: Compare different pipelines and their stages
- `find-high-value-deals`: Find high-value deals

## License

MIT
