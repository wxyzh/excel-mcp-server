# Excel MCP Server

![NPM Version](https://img.shields.io/npm/v/excel-mcp-server)

A Model Context Protocol (MCP) server that reads and writes spreadsheet data to MS Excel file.

## Features

- Read text values from MS Excel file
- Write text values to MS Excel file

## Requirements

- Node.js 20.x or later

## Supported file formats

- xlsx

## Installation

exxel-mcp-server is automatically installed by adding the following configuration to the MCP servers configuration.

For Windows:
```json
{
  "mcpServers": {
    "excel": {
        "command": "cmd",
        "args": ["/c", "npx", "--yes", "excel-mcp-server"],
    }
}
```

For other platforms:
```json
{
  "mcpServers": {
    "excel": {
        "command": "npx",
        "args": ["--yes", "excel-mcp-server"],
    }
}
```

## License

Copyright (c) 2025 Kazuki Negoro

excel-mcp-server is released under the [MIT License](LICENSE)
