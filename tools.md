## SharePoint MCP Tools

This server is designed to return user-friendly SharePoint results first, with structured data available alongside them for downstream use.

### Sites

- `list_sites` - Find SharePoint sites the current user can access.
- `get_site` - Get details for a known SharePoint site ID.
- `list_site_drives` - List the document libraries inside a SharePoint site.

### Document Libraries

- `list_drive_root_items` - List the top-level folders and files in a document library.
- `list_drive_children` - List the folders and files inside a specific folder.
- `get_drive_item` - Get details for a specific file or folder.
- `create_folder` - Create a new folder in a document library.
- `copy_drive_item` - Copy a file or folder into another folder.

### Lists

- `list_lists` - List the SharePoint lists available in a site.
- `list_list_items` - List the items stored in a SharePoint list.
- `create_list_item` - Create a new SharePoint list item.
- `update_list_item` - Update fields on an existing SharePoint list item.

### Permissions and Sharing

- `list_item_permissions` - Show who has access to a file or folder.
- `create_sharing_link` - Create a view or edit link for a file or folder.

### Search

- `search_drive` - Search inside one SharePoint document library.
- `search_sharepoint_global` - Search across SharePoint sites and libraries.

### User and File Reading

- `get_current_user` - Show the connected Microsoft 365 user profile.
- `read_file_content` - Extract readable text from supported SharePoint files such as DOCX, PDF, XLSX, PPTX, TXT, CSV, and JSON.
