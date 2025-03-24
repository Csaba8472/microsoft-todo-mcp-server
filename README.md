# Microsoft Todo MCP Service for Claude

This project provides a Model Context Protocol (MCP) service for Claude that allows you to interact with your Microsoft Todo tasks using natural language.

## Features

### Task List Management (Top-level containers that organize tasks into categories)
- View all your Microsoft Todo task lists
- Create new task lists for better organization
- Update existing task lists
- Delete task lists you no longer need

### Task Management
- Get tasks from specific lists with filtering and sorting options
- Create new tasks with rich details (due dates, priority, body text, etc.)
- Update existing tasks to change any property
- Delete tasks that are no longer needed

### Checklist Item Management (Subtasks)
- View checklist items (subtasks) for a task with completion status
- Create new checklist items to break down tasks
- Update checklist items to mark as complete or edit text
- Delete checklist items when no longer needed

## Setup

1. Clone this repository
2. Install dependencies:
   ```
   npm install
   ```
   
### Setting up Azure App Registration

To use this MCP service, you need to register an application in the Azure portal to get the required credentials:

1. Go to [Azure App Registration Portal](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click on "New registration"
3. Enter a name for your application (e.g., "Todo MCP for Claude")
4. Under "Supported account types", select "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"
5. Set the Redirect URI to "Web" and enter `http://localhost:3000/callback`
6. Click "Register"
7. Once registered, copy the "Application (client) ID" value - this will be your `CLIENT_ID`
8. From the left menu, click on "Certificates & secrets"
9. Under "Client secrets", click "New client secret"
10. Add a description (e.g., "MCP Access") and select an expiration period
11. Click "Add" and immediately copy the "Value" - this will be your `CLIENT_SECRET`
12. From the left menu, click on "API permissions"
13. Click "Add a permission" and select "Microsoft Graph"
14. Select "Delegated permissions"
15. Search for and select the following permissions:
    - Tasks.ReadWrite
    - Tasks.Read
16. Click "Add permissions"
17. Click "Grant admin consent" (if you have admin rights) or have your admin approve the permissions

3. Create a `.env` file using the provided `.env.example` template:
   ```
   # Microsoft Graph API Credentials
   CLIENT_ID=your-client-id
   CLIENT_SECRET=your-client-secret
   TENANT_ID=consumers  # Use 'consumers' for personal Microsoft accounts
   REDIRECT_URI=http://localhost:3000/callback
   
   # Token Storage Path (optional)
   # TOKEN_FILE_PATH=/custom/path/to/tokens.json
   ```
4. Run the authentication server to get your token:
   ```
   npm run auth
   ```
   If your browser doesn't open automatically, manually navigate to:
   ```
   http://localhost:3000
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

   The Claude Desktop configuration file is located at:
   - **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - **Windows**: `%APPDATA%\Claude\claude_desktop_config.json` 
   - **Linux**: `~/.config/Claude/claude_desktop_config.json`

   You can open it with the following commands:

   **macOS**:
   ```
   open ~/Library/Application\ Support/Claude/claude_desktop_config.json
   ```

   **Windows** (PowerShell):
   ```
   notepad $env:APPDATA\Claude\claude_desktop_config.json
   ```

   **Linux**:
   ```
   xdg-open ~/.config/Claude/claude_desktop_config.json
   ```

## Usage

You can interact with the Microsoft Todo MCP service using natural language in Claude. Here are some examples:

### Task Management

**Viewing Tasks**
- "Show me my todo items due this week"
- "What tasks do I have in my Work list?"
- "List all my high priority tasks"
- "Show me tasks that are past due"

**Creating Tasks**
- "Add a new task to buy groceries this weekend"
- "Create a todo item to finish the quarterly report by next Friday"
- "Add 'Call dentist to schedule appointment' to my Personal list"
- "Create a task with high importance to submit project proposal by Tuesday"

**Updating Tasks**
- "Mark the 'Send email to client' task as complete"
- "Change the due date of my report task to next Monday"
- "Update the 'Team meeting' task to include agenda items in the description"
- "Postpone my 'Review documents' task by two days"

**Managing Subtasks**
- "Create a task to plan the company retreat and add subtasks for venue, catering, and activities"
- "Break down my 'Launch website' task into logical steps"
- "Add a subtask 'Buy milk' to my shopping list task"
- "Show me all the subtasks for my project planning task"

### Task List Management

**Managing Lists**
- "Show me all my todo lists"
- "Create a new list called 'Home Renovation'"
- "Rename my 'Work' list to 'Current Projects'"
- "Delete my 'Temporary' task list"

Claude will interpret these natural language requests and translate them into the appropriate Microsoft Todo MCP commands, handling the technical details for you.

## Authentication

The service uses Microsoft's OAuth 2.0 for authentication. The token is stored locally in a `tokens.json` file and will be refreshed automatically when needed.

To re-authenticate, run:
```
npm run auth
```

## License

This project is released under the MIT License. Feel free to modify and distribute it as needed. 