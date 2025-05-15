const readline = require('readline');
const { PublicClientApplication } = require('@azure/msal-node');
const axios = require('axios');
const fs = require('fs').promises;
const path = require('path');
// require('dotenv').config(); // Removed - env vars come from Claude Desktop config

// Configuration
const msalConfig = {
    auth: {
        clientId: process.env.AZURE_CLIENT_ID || 'your AZURE_CLIENT_ID',
        authority: 'https://login.microsoftonline.com/common',
    },
};

const SCOPES = ['Notes.Read', 'Notes.Create', 'Notes.ReadWrite', 'User.Read'];
const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';
const TOKEN_CACHE_FILE = path.join(__dirname, '.mcp-onenote-cache.json');

// Initialize MSAL
const pca = new PublicClientApplication(msalConfig);

// Logging
function log(message, level = 'info') {
    const timestamp = new Date().toISOString();
    console.error(`[${timestamp}] [${level}] ${message}`);
}

// Token management
let cachedToken = null;

async function loadCachedToken() {
    try {
        const data = await fs.readFile(TOKEN_CACHE_FILE, 'utf8');
        cachedToken = JSON.parse(data);
        log('Token loaded from cache');
    } catch (error) {
        log('No cached token found', 'debug');
    }
}

async function saveCachedToken(token) {
    cachedToken = token;
    await fs.writeFile(TOKEN_CACHE_FILE, JSON.stringify(token, null, 2));
    log('Token cached');
}

async function getAccessToken() {
    // Check if we have a valid cached token
    if (cachedToken && new Date(cachedToken.expiresOn) > new Date()) {
        log('Using cached token');
        return cachedToken.accessToken;
    }

    log('Acquiring new token via Device Code flow');
    
    const deviceCodeRequest = {
        deviceCodeCallback: (response) => {
            log(`Please navigate to: ${response.verificationUri}`);
            log(`Enter code: ${response.userCode}`);
        },
        scopes: SCOPES,
        timeout: 300000, // 5 minutes
    };

    try {
        const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
        await saveCachedToken(response);
        log('Authentication successful');
        return response.accessToken;
    } catch (error) {
        log(`Authentication failed: ${error.message}`, 'error');
        throw error;
    }
}

// Microsoft Graph API requests
async function makeGraphRequest(method, endpoint, data = null) {
    const token = await getAccessToken();
    
    try {
        const response = await axios({
            method,
            url: `${GRAPH_BASE_URL}${endpoint}`,
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
            },
            data,
        });
        return response.data;
    } catch (error) {
        log(`Graph API error: ${error.response?.data?.error?.message || error.message}`, 'error');
        throw error;
    }
}

// OneNote operations
async function readNotebooks() {
    log('Reading notebooks');
    const response = await makeGraphRequest('GET', '/me/onenote/notebooks');
    return response.value;
}

async function readSections(notebookId) {
    log(`Reading sections for notebook: ${notebookId}`);
    const response = await makeGraphRequest('GET', `/me/onenote/notebooks/${notebookId}/sections`);
    return response.value;
}

async function readPages(sectionId) {
    log(`Reading pages for section: ${sectionId}`);
    const response = await makeGraphRequest('GET', `/me/onenote/sections/${sectionId}/pages`);
    return response.value;
}

async function readPageContent(pageId) {
    log(`Reading content for page: ${pageId}`);
    const content = await makeGraphRequest('GET', `/me/onenote/pages/${pageId}/content`);
    return content;
}

async function createNotebook(displayName) {
    log(`Creating notebook: ${displayName}`);
    const data = { displayName };
    return await makeGraphRequest('POST', '/me/onenote/notebooks', data);
}

async function createSection(notebookId, displayName) {
    log(`Creating section: ${displayName} in notebook: ${notebookId}`);
    const data = { displayName };
    return await makeGraphRequest('POST', `/me/onenote/notebooks/${notebookId}/sections`, data);
}

async function createPage(sectionId, title, content) {
    log(`Creating page: ${title} in section: ${sectionId}`);
    
    const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
        <title>${title}</title>
    </head>
    <body>
        ${content}
    </body>
    </html>`;
    
    const token = await getAccessToken();
    
    try {
        const response = await axios({
            method: 'POST',
            url: `${GRAPH_BASE_URL}/me/onenote/sections/${sectionId}/pages`,
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'text/html',
            },
            data: htmlContent,
        });
        return response.data;
    } catch (error) {
        log(`Page creation error: ${error.response?.data?.error?.message || error.message}`, 'error');
        throw error;
    }
}

// MCP Protocol Implementation
async function init() {
    log('Starting OneNote MCP Server');
    
    if (!process.env.AZURE_CLIENT_ID) {
        log('Warning: AZURE_CLIENT_ID not set. Please set it before using the server.', 'warn');
    }
    
    // Load cached token if available
    await loadCachedToken();
    
    // Create readline interface
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout,
        terminal: false
    });
    
    // Handle incoming messages
    rl.on('line', async (line) => {
        try {
            log(`Received: ${line}`, 'debug');
            const message = JSON.parse(line);
            const response = await handleMessage(message);
            sendResponse(response);
        } catch (error) {
            log(`Error processing message: ${error.message}`, 'error');
            sendError(null, -32700, 'Parse error');
        }
    });
}

async function handleMessage(message) {
    const { method, params, id } = message;
    
    log(`Handling method: ${method}`, 'debug');

    switch (method) {
        case 'initialize':
            return handleInitialize(params, id);
        case 'initialized':
            return null; // This is just a notification, no response needed
        case 'notifications/initialized':
            // This is a notification, no response needed
            log('Client initialized', 'info');
            return null;
        case 'tools/list':
            return handleToolsList(id);
        case 'tools/call':
            return handleToolCall(params, id);
        case 'notifications/cancelled':
            // This is a notification, no response needed
            log(`Request ${params?.requestId} was cancelled: ${params?.reason}`, 'info');
            return null;
        default:
            return createError(id, -32601, `Method not found: ${method}`);
    }
}

function handleInitialize(params, id) {
    log('Handling initialize request', 'debug');
    const response = {
        jsonrpc: '2.0',
        id,
        result: {
            protocolVersion: '2024-11-05',
            capabilities: {
                tools: {},
                logging: {}
            },
            serverInfo: {
                name: 'OneNote MCP Server',
                version: '1.0.0'
            }
        }
    };
    log('Initialize response prepared', 'debug');
    return response;
}

function handleInitialized(id) {
    log('Server initialized', 'info');
    if (id !== null) {
        return { jsonrpc: '2.0', id, result: {} };
    }
    return null;
}

function handleToolsList(id) {
    return {
        jsonrpc: '2.0',
        id,
        result: {
            tools: [
                {
                    name: 'onenote-read',
                    description: 'Read OneNote content (notebooks, sections, pages)',
                    inputSchema: {
                        type: 'object',
                        properties: {
                            type: {
                                type: 'string',
                                enum: ['list_notebooks', 'list_sections', 'list_pages', 'read_content'],
                                description: 'Type of read operation'
                            },
                            notebookId: {
                                type: 'string',
                                description: 'Notebook ID (required for list_sections)'
                            },
                            sectionId: {
                                type: 'string',
                                description: 'Section ID (required for list_pages)'
                            },
                            pageId: {
                                type: 'string',
                                description: 'Page ID (required for read_content)'
                            }
                        },
                        required: ['type']
                    }
                },
                {
                    name: 'onenote-create',
                    description: 'Create OneNote content (notebooks, sections, pages)',
                    inputSchema: {
                        type: 'object',
                        properties: {
                            type: {
                                type: 'string',
                                enum: ['create_notebook', 'create_section', 'create_page'],
                                description: 'Type of create operation'
                            },
                            displayName: {
                                type: 'string',
                                description: 'Display name for notebook or section'
                            },
                            notebookId: {
                                type: 'string',
                                description: 'Notebook ID (required for create_section)'
                            },
                            sectionId: {
                                type: 'string',
                                description: 'Section ID (required for create_page)'
                            },
                            title: {
                                type: 'string',
                                description: 'Page title (required for create_page)'
                            },
                            content: {
                                type: 'string',
                                description: 'HTML content for the page'
                            }
                        },
                        required: ['type']
                    }
                }
            ]
        }
    };
}

async function handleToolCall(params, id) {
    const { name, arguments: args } = params;

    try {
        let result;
        
        switch (name) {
            case 'onenote-read':
                result = await handleReadCommand(args);
                break;
            case 'onenote-create':
                result = await handleCreateCommand(args);
                break;
            default:
                return createError(id, -32602, `Unknown tool: ${name}`);
        }

        return {
            jsonrpc: '2.0',
            id,
            result: {
                content: [
                    {
                        type: 'text',
                        text: JSON.stringify(result, null, 2)
                    }
                ]
            }
        };
    } catch (error) {
        return createError(id, -32603, error.message);
    }
}

async function handleReadCommand(input) {
    switch (input.type) {
        case 'list_notebooks':
            return await readNotebooks();
            
        case 'list_sections':
            if (!input.notebookId) throw new Error('notebookId is required');
            return await readSections(input.notebookId);
            
        case 'list_pages':
            if (!input.sectionId) throw new Error('sectionId is required');
            return await readPages(input.sectionId);
            
        case 'read_content':
            if (!input.pageId) throw new Error('pageId is required');
            const content = await readPageContent(input.pageId);
            return {
                pageId: input.pageId,
                content: content
            };
            
        default:
            throw new Error(`Unknown read type: ${input.type}`);
    }
}

async function handleCreateCommand(input) {
    switch (input.type) {
        case 'create_notebook':
            if (!input.displayName) throw new Error('displayName is required');
            return await createNotebook(input.displayName);
            
        case 'create_section':
            if (!input.notebookId || !input.displayName) {
                throw new Error('notebookId and displayName are required');
            }
            return await createSection(input.notebookId, input.displayName);
            
        case 'create_page':
            if (!input.sectionId || !input.title) {
                throw new Error('sectionId and title are required');
            }
            const content = input.content || '<p>New page</p>';
            return await createPage(input.sectionId, input.title, content);
            
        default:
            throw new Error(`Unknown create type: ${input.type}`);
    }
}

function createError(id, code, message) {
    return {
        jsonrpc: '2.0',
        id,
        error: {
            code,
            message
        }
    };
}

function sendResponse(response) {
    if (response) {
        const responseStr = JSON.stringify(response);
        console.log(responseStr);
        log(`Sent response: ${responseStr}`, 'debug');
    }
}

function sendError(id, code, message) {
    sendResponse(createError(id, code, message));
}

// Handle process termination
process.on('SIGINT', () => {
    log('Server shutting down');
    process.exit(0);
});

// Start the server
init().catch(error => {
    log(`Fatal error: ${error.message}`, 'error');
    process.exit(1);
});