// Microsoft Graph API Test for To Do Tasks (Delegated Permissions)
require('dotenv').config();
const msal = require('@azure/msal-node');
const { Client } = require('@microsoft/microsoft-graph-client');
require('isomorphic-fetch');
const express = require('express');

const app = express();
const port = 3000;

// MSAL configuration for delegated permissions
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: 'https://login.microsoftonline.com/consumers',
    clientSecret: process.env.CLIENT_SECRET,
  }
};

// Create MSAL application
const cca = new msal.ConfidentialClientApplication(msalConfig);

// Task-related permission scopes
const scopes = [
  'Tasks.Read',
  'Tasks.Read.Shared',
  'Tasks.ReadWrite',
  'Tasks.ReadWrite.Shared'
];

// Setup the auth flow
app.get('/', (req, res) => {
  const authCodeUrlParameters = {
    scopes: scopes,
    redirectUri: process.env.REDIRECT_URI,
  };

  cca.getAuthCodeUrl(authCodeUrlParameters)
    .then((response) => {
      res.redirect(response);
    })
    .catch((error) => {
      console.error(JSON.stringify(error));
      res.status(500).send(error);
    });
});

// Handle the callback from Microsoft login
app.get('/callback', (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes: scopes,
    redirectUri: process.env.REDIRECT_URI,
  };

  cca.acquireTokenByCode(tokenRequest)
    .then((response) => {
      // Use the token to call Microsoft Graph API
      accessMicrosoftTodoApi(response.accessToken)
        .then(() => {
          res.send('Successfully accessed Microsoft To Do API. Check console for results.');
        })
        .catch(error => {
          console.error('API access error:', error);
          res.status(500).send('Error accessing Microsoft To Do API');
        });
    })
    .catch((error) => {
      console.error(JSON.stringify(error));
      res.status(500).send(error);
    });
});

// Function to access Microsoft To Do API using the access token
async function accessMicrosoftTodoApi(accessToken) {
  // Initialize Graph client
  const client = Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });

  try {
    // Get all task lists
    console.log("Fetching task lists...");
    const taskLists = await client.api('/me/todo/lists').get();
    console.log("Task lists:", JSON.stringify(taskLists, null, 2));
    
    // If task lists exist, get tasks from the first list
    if (taskLists.value && taskLists.value.length > 0) {
      const firstListId = taskLists.value[0].id;
      console.log(`Fetching tasks from list: ${firstListId}`);
      const tasks = await client.api(`/me/todo/lists/${firstListId}/tasks`).get();
      console.log("Tasks:", JSON.stringify(tasks, null, 2));
      
      // If tasks exist, get checklist items for the first task
      if (tasks.value && tasks.value.length > 0) {
        const firstTaskId = tasks.value[0].id;
        console.log(`Fetching checklist items for task: ${firstTaskId}`);
        const checklistItems = await client.api(`/me/todo/lists/${firstListId}/tasks/${firstTaskId}/checklistItems`).get();
        console.log("Checklist items:", JSON.stringify(checklistItems, null, 2));
      }
    }
  } catch (error) {
    console.error("Error accessing Microsoft To Do API:", error);
    throw error;
  }
}

// Start the server
app.listen(port, () => {
  console.log(`Server running at http://localhost:${port}`);
  console.log('Please open your browser and navigate to: http://localhost:${port}');
}); 