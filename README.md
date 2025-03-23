# Microsoft Todo MCP Service for Claude

This project provides a Model Context Protocol (MCP) service for Claude that allows you to interact with your Microsoft Todo tasks using natural language.

## Features

- View all your Microsoft Todo task lists
- See tasks in a specific list
- Create new tasks with optional due dates and priority
- View checklist items for a task

## Setup

1. Clone this repository
2. Install dependencies:
   ```
   npm install
   ```
3. Create a `.env` file with your Microsoft Graph API credentials:
   ```
   CLIENT_ID=your-client-id
   CLIENT_SECRET=your-client-secret
   TENANT_ID=your-tenant-id
   REDIRECT_URI=http://localhost:3000/callback
   ```
4. Run the authentication server to get your token:
   ```
   npm run auth
   ```
5. Build the MCP service:
   ```
   npm run build
   ```
6. Update your Claude Desktop configuration to include this MCP service:
   ```json
   {
       "mcpServers": {
           "mstodo": {
               "command": "node",
               "args": [
                   "/path/to/your/build/todo-index.js"
               ]
           }
       }
   }
   ```

## Usage

Once set up, you can use the following commands in Claude:

- `@mstodo auth-status` - Check your authentication status
- `@mstodo get-task-lists` - Get all your task lists
- `@mstodo get-tasks [list-id]` - Get tasks from a specific list
- `@mstodo create-task [list-id] [title] [due-date] [importance]` - Create a new task
- `@mstodo get-checklist-items [list-id] [task-id]` - Get checklist items for a task

## Authentication

The service uses Microsoft's OAuth 2.0 for authentication. The token is stored locally in a `tokens.json` file and will be refreshed automatically when needed.

To re-authenticate, run:
```
npm run auth
``` 