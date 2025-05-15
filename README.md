OneNote MCP Server
A Model Context Protocol (MCP) server for Microsoft OneNote that enables reading and creating notebooks, sections, and pages via Microsoft Graph API.

Features
Read OneNote notebooks, sections, and page content
Create new notebooks, sections, and pages
Authentication via Microsoft Device Code flow
Token caching for seamless re-authentication
Comprehensive error handling and logging
Prerequisites
Node.js 14.0.0 or higher
Microsoft Azure AD app registration with Graph API permissions
Setup
1. Create Azure AD App Registration
Go to Azure Portal
Navigate to Azure Active Directory → App registrations → New registration
Name your app (e.g., "OneNote MCP Server")
Select "Accounts in any organizational directory and personal Microsoft accounts"
Set redirect URI to "Public client/native" with value https://login.microsoftonline.com/common/oauth2/nativeclient
After creation, note the Application (client) ID
2. Configure API Permissions
In your app registration, go to API permissions
Add the following Microsoft Graph permissions:
Notes.Read
Notes.Create
Notes.ReadWrite
User.Read
These are all delegated permissions (not application)
3. Install Dependencies
bash
npm install
4. Set Environment Variables
bash
export AZURE_CLIENT_ID="your-client-id-here"
Or create a .env file:

AZURE_CLIENT_ID=your-client-id-here
Running the Server
bash
npm start
The server will read JSON commands from stdin and output responses to stdout.

Usage Examples
List All Notebooks
json
{
  "tool": "onenote-read",
  "input": {
    "type": "list_notebooks"
  }
}
List Sections in a Notebook
json
{
  "tool": "onenote-read",
  "input": {
    "type": "list_sections",
    "notebookId": "notebook-id-here"
  }
}
List Pages in a Section
json
{
  "tool": "onenote-read",
  "input": {
    "type": "list_pages",
    "sectionId": "section-id-here"
  }
}
Read Page Content
json
{
  "tool": "onenote-read",
  "input": {
    "type": "read_content",
    "pageId": "page-id-here"
  }
}
Create a Notebook
json
{
  "tool": "onenote-create",
  "input": {
    "type": "create_notebook",
    "displayName": "My New Notebook"
  }
}
Create a Section
json
{
  "tool": "onenote-create",
  "input": {
    "type": "create_section",
    "notebookId": "notebook-id-here",
    "displayName": "My New Section"
  }
}
Create a Page
json
{
  "tool": "onenote-create",
  "input": {
    "type": "create_page",
    "sectionId": "section-id-here",
    "title": "My New Page",
    "content": "<h1>Welcome</h1><p>This is my new page content.</p>"
  }
}
Testing with Echo
You can test the server using echo:

bash
echo '{"tool":"onenote-read","input":{"type":"list_notebooks"}}' | node server.js
Authentication
On first run, the server will prompt you to:

Navigate to a Microsoft login URL
Enter a device code
Authenticate with your Microsoft account
The token is cached in .mcp-onenote-cache.json for subsequent runs.

Error Handling
All errors are logged to stderr with timestamps. The server returns JSON responses with:

success: boolean indicating if the operation succeeded
result: the data if successful
error: error message if failed
Logging
Logs are written to stderr in the format:

[timestamp] [level] message
Security Notes
The .mcp-onenote-cache.json file contains sensitive tokens. Add it to .gitignore
Never commit your Azure Client ID to version control
The server uses delegated permissions, requiring user consent
License
MIT

