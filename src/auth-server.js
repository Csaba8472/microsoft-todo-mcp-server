// Authentication server for Microsoft Todo MCP service
import dotenv from 'dotenv';
import express from 'express';
import { writeFileSync } from 'fs';
import { join } from 'path';
import { ConfidentialClientApplication } from '@azure/msal-node';
import { fileURLToPath } from 'url';
import { dirname } from 'path';

// Initialize environment variables
dotenv.config();
console.log('Environment loaded');
console.log('CLIENT_ID:', process.env.CLIENT_ID ? 'Present (hidden)' : 'Missing');
console.log('CLIENT_SECRET:', process.env.CLIENT_SECRET ? 'Present (hidden)' : 'Missing');
console.log('TENANT_ID:', process.env.TENANT_ID ? 'Present (hidden)' : 'Missing');
console.log('REDIRECT_URI:', process.env.REDIRECT_URI || `http://localhost:3000/callback`);

// Get current file directory in ESM
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

const app = express();
const port = 3000;
const TOKEN_FILE_PATH = join(process.cwd(), 'tokens.json');

// MSAL configuration for delegated permissions
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
  system: {
    loggerOptions: {
      loggerCallback(loglevel, message, containsPii) {
        console.log(`MSAL Log: ${message}`);
      },
      piiLoggingEnabled: true
    }
  },
  cache: {
    cachePlugin: {
      beforeCacheAccess: async (cacheContext) => {
        console.log('Cache access requested:', cacheContext);
        return null;
      },
      afterCacheAccess: async (cacheContext) => {
        console.log('Cache access completed:', cacheContext);
        return null;
      }
    }
  }
};

console.log('MSAL config created');

// Task-related permission scopes
const scopes = [
  'offline_access',  // Put offline_access first to ensure it's not dropped
  'openid',         // Add openid scope
  'profile',        // Add profile scope
  'Tasks.Read',
  'Tasks.Read.Shared',
  'Tasks.ReadWrite',
  'Tasks.ReadWrite.Shared',
  'User.Read'
];

// Create MSAL application
const cca = new ConfidentialClientApplication(msalConfig);
console.log('MSAL application created');

// Setup a test route to check if server is working
app.get('/test', (req, res) => {
  res.send('Auth server is running correctly');
});

// Setup the auth flow
app.get('/', (req, res) => {
  console.log('Root route accessed, generating auth URL...');
  const authCodeUrlParameters = {
    scopes: scopes,
    redirectUri: process.env.REDIRECT_URI || `http://localhost:${port}/callback`,
    prompt: 'consent',  // Use only consent to force refresh token
    responseMode: 'query',
  };

  console.log('Auth parameters:', {
    scopes: scopes,
    redirectUri: process.env.REDIRECT_URI || `http://localhost:${port}/callback`,
    prompt: 'consent',
    responseMode: 'query',
  });

  cca.getAuthCodeUrl(authCodeUrlParameters)
    .then((response) => {
      console.log('Auth URL generated, redirecting to:', response.substring(0, 80) + '...');
      res.redirect(response);
    })
    .catch((error) => {
      console.error('Error getting auth code URL:', error);
      res.status(500).send(`Error generating authentication URL: ${JSON.stringify(error)}`);
    });
});

// Handle the callback from Microsoft login
app.get('/callback', (req, res) => {
  console.log('Callback route accessed');
  console.log('Query parameters:', {
    code: req.query.code ? 'Present (hidden)' : 'Missing',
    state: req.query.state ? 'Present' : 'Missing',
    error: req.query.error || 'None',
    error_description: req.query.error_description || 'None'
  });

  const tokenRequest = {
    code: req.query.code,
    scopes: scopes,
    redirectUri: process.env.REDIRECT_URI || `http://localhost:${port}/callback`,
  };

  console.log('Token request parameters:', {
    scopes: scopes,
    redirectUri: process.env.REDIRECT_URI || `http://localhost:${port}/callback`,
  });

  cca.acquireTokenByCode(tokenRequest)
    .then((response) => {
      try {
        // Log full response structure (without sensitive values)
        console.log('Token response structure:', {
          keys: Object.keys(response),
          hasAccessToken: !!response.accessToken,
          hasRefreshToken: !!response.refreshToken,
          hasIdToken: !!response.idToken,
          tokenType: response.tokenType,
          expiresIn: response.expiresIn,
          expiresOn: response.expiresOn,
          scopes: response.scopes,
          account: response.account ? {
            username: response.account.username,
            tenantId: response.account.tenantId,
            localAccountId: response.account.localAccountId
          } : null
        });

        // Calculate token expiration (make sure it's never null)
        const expiresInSeconds = response.expiresIn || 3600;
        const expiresAt = Date.now() + (expiresInSeconds * 1000) - (5 * 60 * 1000);
        
        console.log('Token expiration details:', {
          expiresInSeconds,
          expiresAt: new Date(expiresAt).toLocaleString(),
          currentTime: new Date().toLocaleString()
        });
        
        // Store tokens
        const tokenData = {
          accessToken: response.accessToken,
          refreshToken: response.refreshToken || '',
          expiresAt: expiresAt,
          tokenType: response.tokenType,
          scopes: response.scopes
        };
        
        writeFileSync(TOKEN_FILE_PATH, JSON.stringify(tokenData, null, 2), 'utf8');
        
        console.log('Authentication successful! Token saved to:', TOKEN_FILE_PATH);
        
        // Format token display with safety checks
        const accessTokenDisplay = response.accessToken ? 
          `${response.accessToken.substring(0, 15)}...${response.accessToken.substring(response.accessToken.length - 5)}` : 
          'Not provided';
          
        const refreshTokenDisplay = response.refreshToken ? 
          `${response.refreshToken.substring(0, 10)}...${response.refreshToken.substring(response.refreshToken.length - 5)}` : 
          'Not provided';
        
        res.send(`
          <h1>Authentication Successful!</h1>
          <p>You can now close this window and use the Microsoft Todo MCP service.</p>
          <p>Token details:</p>
          <ul>
            <li>Access Token: ${accessTokenDisplay}</li>
            <li>Refresh Token: ${refreshTokenDisplay}</li>
            <li>Token Type: ${response.tokenType || 'Not provided'}</li>
            <li>Scopes: ${response.scopes ? response.scopes.join(', ') : 'Not provided'}</li>
            <li>Expires: ${new Date(expiresAt).toLocaleString()}</li>
          </ul>
          <p>Debug Information:</p>
          <pre>${JSON.stringify({
            hasRefreshToken: !!response.refreshToken,
            tokenType: response.tokenType,
            scopes: response.scopes
          }, null, 2)}</pre>
        `);
      } catch (error) {
        console.error('Error saving token:', error);
        res.status(500).send(`Error saving token: ${error.message}`);
      }
    })
    .catch((error) => {
      console.error('Token acquisition error:', {
        errorCode: error.errorCode,
        errorMessage: error.errorMessage,
        subError: error.subError,
        correlationId: error.correlationId,
        stack: error.stack
      });
      res.status(500).send(`Error acquiring token: ${JSON.stringify(error)}`);
    });
});

// Start the server
app.listen(port, () => {
  console.log(`Auth server running at http://localhost:${port}`);
  console.log('Open your browser and navigate to the URL above to authenticate.');
  console.log('Or try http://localhost:3000/test to verify the server is running.');
}); 