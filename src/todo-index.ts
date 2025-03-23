import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { readFileSync, writeFileSync, existsSync } from "fs";
import { join } from "path";
import dotenv from "dotenv";

// Load environment variables
dotenv.config();

// Log the current working directory
console.error('Current working directory:', process.cwd());

// Microsoft Graph API endpoints
const MS_GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const USER_AGENT = "ms-todo-mcp/1.0";
// Use absolute path for token file
const TOKEN_FILE_PATH = '/Users/jhirono/Dev/todoMCP/tokens.json';
console.error('Using token file path:', TOKEN_FILE_PATH);

// Create server instance
const server = new McpServer({
  name: "mstodo",
  version: "1.0.0",
});

// Token types
interface TokenData {
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
}

// Helper to read tokens from file
function readTokens(): TokenData | null {
  try {
    console.error(`Attempting to read tokens from: ${TOKEN_FILE_PATH}`);
    if (!existsSync(TOKEN_FILE_PATH)) {
      console.error('Token file does not exist');
      return null;
    }
    const data = readFileSync(TOKEN_FILE_PATH, 'utf8');
    console.error('Token file content length:', data.length);
    
    const tokenData = JSON.parse(data) as TokenData;
    console.error('Token parsed successfully, expires at:', new Date(tokenData.expiresAt).toLocaleString());
    return tokenData;
  } catch (error) {
    console.error('Failed to read tokens from file:', error);
    return null;
  }
}

// Helper to write tokens to file
function writeTokens(tokenData: TokenData): void {
  try {
    writeFileSync(TOKEN_FILE_PATH, JSON.stringify(tokenData, null, 2), 'utf8');
  } catch (error) {
    console.error('Failed to write tokens to file:', error);
  }
}

// Helper function for making Microsoft Graph API requests
async function makeGraphRequest<T>(url: string, token: string, method = "GET", body?: any): Promise<T | null> {
  const headers = {
    "User-Agent": USER_AGENT,
    "Accept": "application/json",
    "Authorization": `Bearer ${token}`,
    "Content-Type": "application/json"
  };

  try {
    const options: RequestInit = { 
      method, 
      headers 
    };

    if (body && (method === "POST" || method === "PATCH")) {
      options.body = JSON.stringify(body);
    }

    const response = await fetch(url, options);
    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }
    return (await response.json()) as T;
  } catch (error) {
    console.error("Error making Graph API request:", error);
    return null;
  }
}

// Refresh token function
async function refreshAccessToken(refreshToken: string): Promise<TokenData | null> {
  const tokenEndpoint = `https://login.microsoftonline.com/consumers/oauth2/v2.0/token`;
  
  const formData = new URLSearchParams({
    client_id: process.env.CLIENT_ID || "",
    client_secret: process.env.CLIENT_SECRET || "",
    refresh_token: refreshToken,
    grant_type: "refresh_token",
    scope: "Tasks.Read Tasks.ReadWrite Tasks.Read.Shared Tasks.ReadWrite.Shared"
  });

  try {
    const response = await fetch(tokenEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded"
      },
      body: formData
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Token refresh failed: ${response.status} ${errorText}`);
    }

    const data = await response.json();
    
    // Calculate expiration time (subtract 5 minutes for safety margin)
    const expiresAt = Date.now() + (data.expires_in * 1000) - (5 * 60 * 1000);
    
    const tokenData: TokenData = {
      accessToken: data.access_token,
      refreshToken: data.refresh_token || refreshToken, // Use new refresh token if provided
      expiresAt
    };
    
    // Save the new tokens
    writeTokens(tokenData);
    
    return tokenData;
  } catch (error) {
    console.error("Error refreshing token:", error);
    return null;
  }
}

// Authentication helper using delegated flow with refresh token
async function getAccessToken(): Promise<string | null> {
  try {
    console.error('getAccessToken called');
    
    try {
      // Read token directly from the absolute path
      const tokenFilePath = '/Users/jhirono/Dev/todoMCP/tokens.json';
      console.error(`Directly reading token from: ${tokenFilePath}`);
      
      // Read file synchronously to avoid any async issues
      const data = readFileSync(tokenFilePath, 'utf8');
      console.error(`Read ${data.length} bytes from token file`);
      
      // Parse the token data
      const tokenData = JSON.parse(data) as TokenData;
      console.error(`Token parsed, expires at: ${new Date(tokenData.expiresAt).toLocaleString()}`);
      
      // Check if token is expired
      const now = Date.now();
      if (now > tokenData.expiresAt) {
        console.error(`Token is expired. Current time: ${now}, expires at: ${tokenData.expiresAt}`);
        return null;
      }
      
      // Success - return the token
      console.error(`Successfully retrieved valid token (${tokenData.accessToken.substring(0, 10)}...)`);
      return tokenData.accessToken;
    } catch (readError) {
      console.error(`Direct token read error: ${readError}`);
      return null;
    }
  } catch (error) {
    console.error("Error getting access token:", error);
    return null;
  }
}

// Server tool to check authentication status
server.tool(
  "auth-status",
  "Check the status of Microsoft Graph API authentication",
  {},
  async () => {
    const tokens = readTokens();
    if (!tokens) {
      return {
        content: [
          {
            type: "text",
            text: "Not authenticated. Please run auth-server.js to authenticate with Microsoft.",
          },
        ],
      };
    }
    
    const isExpired = Date.now() > tokens.expiresAt;
    const expiryTime = new Date(tokens.expiresAt).toLocaleString();
    
    if (isExpired) {
      return {
        content: [
          {
            type: "text",
            text: `Authentication expired at ${expiryTime}. Will attempt to refresh when you call any API.`,
          },
        ],
      };
    } else {
      return {
        content: [
          {
            type: "text",
            text: `Authenticated. Token expires at ${expiryTime}.`,
          },
        ],
      };
    }
  }
);

interface TaskList {
  id: string;
  displayName: string;
}

interface Task {
  id: string;
  title: string;
  status: string;
  importance: string;
  dueDateTime?: {
    dateTime: string;
    timeZone: string;
  };
}

interface ChecklistItem {
  id: string;
  displayName: string;
  isChecked: boolean;
}

// Register tools
server.tool(
  "get-task-lists",
  "Get all Microsoft Todo task lists",
  {},
  async () => {
    try {
      const token = await getAccessToken();
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        };
      }

      const response = await makeGraphRequest<{ value: TaskList[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists`,
        token
      );

      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to retrieve task lists",
            },
          ],
        };
      }

      const lists = response.value || [];
      if (lists.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No task lists found.",
            },
          ],
        };
      }

      const formattedLists = lists.map((list) => 
        `ID: ${list.id}\nName: ${list.displayName}\n---`
      );

      return {
        content: [
          {
            type: "text",
            text: `Your task lists:\n\n${formattedLists.join("\n")}`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching task lists: ${error}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "get-tasks",
  "Get tasks from a specific list",
  {
    listId: z.string().describe("ID of the task list"),
  },
  async ({ listId }) => {
    try {
      const token = await getAccessToken();
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        };
      }

      const response = await makeGraphRequest<{ value: Task[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks`,
        token
      );
      
      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve tasks for list: ${listId}`,
            },
          ],
        };
      }

      const tasks = response.value || [];
      if (tasks.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No tasks found in list with ID: ${listId}`,
            },
          ],
        };
      }

      const formattedTasks = tasks.map((task) => {
        const status = task.status === "completed" ? "✓" : "○";
        const dueDate = task.dueDateTime 
          ? `Due: ${new Date(task.dueDateTime.dateTime).toLocaleDateString()}` 
          : "No due date";
        const importance = task.importance ? `Importance: ${task.importance}` : "";
        
        return `${status} ID: ${task.id}\nTitle: ${task.title}\n${dueDate}\n${importance}\n---`;
      });

      return {
        content: [
          {
            type: "text",
            text: `Tasks in list ${listId}:\n\n${formattedTasks.join("\n")}`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching tasks: ${error}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "create-task",
  "Create a new task in a specific list",
  {
    listId: z.string().describe("ID of the task list"),
    title: z.string().describe("Title of the task"),
    dueDateTime: z.string().optional().describe("Due date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    importance: z.enum(["low", "normal", "high"]).optional().describe("Task importance"),
  },
  async ({ listId, title, dueDateTime, importance }) => {
    try {
      const token = await getAccessToken();
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        };
      }

      const taskBody: any = { title };

      if (dueDateTime) {
        taskBody.dueDateTime = {
          dateTime: dueDateTime,
          timeZone: "UTC",
        };
      }

      if (importance) {
        taskBody.importance = importance;
      }

      const response = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks`,
        token,
        "POST",
        taskBody
      );
      
      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to create task in list: ${listId}`,
            },
          ],
        };
      }

      return {
        content: [
          {
            type: "text",
            text: `Task created successfully!\nID: ${response.id}\nTitle: ${response.title}`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error creating task: ${error}`,
          },
        ],
      };
    }
  }
);

server.tool(
  "get-checklist-items",
  "Get checklist items for a specific task",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken();
      if (!token) {
        return {
          content: [
            {
              type: "text",
              text: "Failed to authenticate with Microsoft API",
            },
          ],
        };
      }

      const response = await makeGraphRequest<{ value: ChecklistItem[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
        token
      );
      
      if (!response) {
        return {
          content: [
            {
              type: "text",
              text: `Failed to retrieve checklist items for task: ${taskId}`,
            },
          ],
        };
      }

      const items = response.value || [];
      if (items.length === 0) {
        return {
          content: [
            {
              type: "text",
              text: `No checklist items found for task with ID: ${taskId}`,
            },
          ],
        };
      }

      const formattedItems = items.map((item) => {
        const status = item.isChecked ? "✓" : "○";
        return `${status} ${item.displayName} (ID: ${item.id})`;
      });

      return {
        content: [
          {
            type: "text",
            text: `Checklist items for task ${taskId}:\n\n${formattedItems.join("\n")}`,
          },
        ],
      };
    } catch (error) {
      return {
        content: [
          {
            type: "text",
            text: `Error fetching checklist items: ${error}`,
          },
        ],
      };
    }
  }
);

// Main function to start the server
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Microsoft Todo MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
}); 