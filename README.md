OneNote MCP Server
A Model Context Protocol (MCP) server for Microsoft OneNote that enables reading and creating notebooks, sections, and pages through Claude Desktop or any MCP-compatible client.

Features
🔐 Secure authentication via Microsoft Device Code flow
📚 List all OneNote notebooks
📑 Browse sections within notebooks
📄 View pages in sections
✍️ Create new notebooks, sections, and pages
💾 Token caching for seamless re-authentication
🔧 Full MCP protocol implementation
Prerequisites
Node.js 14.0.0 or higher
Microsoft Azure AD app registration with Graph API permissions
Claude Desktop (or any MCP-compatible client)
Quick Start
Clone the repository:
bash
git clone https://github.com/YOUR_USERNAME/onenote-mcp-server.git
cd onenote-mcp-server
Install dependencies:
bash
npm install
Set up Azure AD app registration (see Setup section below)
Configure Claude Desktop (see Configuration section below)
Setup
1. Create Azure AD App Registration
Go to Azure Portal
Navigate to Azure Active Directory → App registrations → New registration
Name your app (e.g., "OneNote MCP Server")
Select "Accounts in any organizational directory and personal Microsoft accounts"
Set redirect URI to "Public client/native" with value:
https://login.microsoftonline.com/common/oauth2/nativeclient
After creation, note the Application (client) ID
2. Configure API Permissions
In your app registration, go to API permissions
Add the following Microsoft Graph delegated permissions:
Notes.Read
Notes.Create
Notes.ReadWrite
User.Read
Grant admin consent if required
3. Configure Claude Desktop
Add the following to your Claude Desktop configuration file:

Windows: %APPDATA%\Claude\claude_desktop_config.json
macOS: ~/Library/Application Support/Claude/claude_desktop_config.json
Linux: ~/.config/Claude/claude_desktop_config.json

json
{
  "mcpServers": {
    "onenote": {
      "command": "node",
      "args": ["/path/to/onenote-mcp-server/server.js"],
      "env": {
        "AZURE_CLIENT_ID": "your-azure-client-id-here"
      }
    }
  }
}
Replace /path/to/onenote-mcp-server/ with the actual path to your cloned repository.

Usage
Once configured, you can interact with OneNote through Claude Desktop:

"List my OneNote notebooks"
"Create a new notebook called 'My Ideas'"
"Show me the sections in notebook 'Work Notes'"
"Create a page titled 'Meeting Notes' in section 'January 2024'"
"Read the content of page 'Project Plan'"
API Documentation
Tools
onenote-read
Read OneNote content (notebooks, sections, pages).

Parameters:

type: Operation type
list_notebooks: List all notebooks
list_sections: List sections in a notebook
list_pages: List pages in a section
read_content: Read page content
notebookId: Required for list_sections
sectionId: Required for list_pages
pageId: Required for read_content
onenote-create
Create OneNote content (notebooks, sections, pages).

Parameters:

type: Operation type
create_notebook: Create a new notebook
create_section: Create a new section
create_page: Create a new page
displayName: Required for notebooks and sections
notebookId: Required for create_section
sectionId: Required for create_page
title: Required for pages
content: HTML content for pages (optional)
Troubleshooting
Authentication Issues
Ensure your Azure Client ID is correct
Check that all required permissions are granted
Delete .mcp-onenote-cache.json to force re-authentication
Connection Issues
Check Claude Desktop logs for error messages
Verify the server path in your configuration
Ensure Node.js is installed and accessible
Common Errors
MODULE_NOT_FOUND: Run npm install to install dependencies
Invalid grant: Delete the token cache file and re-authenticate
Permission denied: Check Azure AD permissions
Security
Never commit your Azure Client ID to version control
The .mcp-onenote-cache.json file contains sensitive tokens - it's automatically excluded via .gitignore
Use environment variables or Claude Desktop's configuration for credentials
Contributing
Fork the repository
Create a feature branch (git checkout -b feature/amazing-feature)
Commit your changes (git commit -m 'Add amazing feature')
Push to the branch (git push origin feature/amazing-feature)
Open a Pull Request
License
This project is licensed under the MIT License - see the LICENSE file for details.

Acknowledgments
Microsoft Graph API for OneNote integration
Claude Desktop for MCP support
The Model Context Protocol community
Support
For issues and questions:

Open an issue on GitHub
Check the MCP documentation
Review Microsoft Graph API documentation
