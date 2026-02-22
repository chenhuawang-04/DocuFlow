# -*- coding: utf-8 -*-
"""
DocuFlow MCP - MCP Server Main Entry

Word Document Processing MCP Server
"""

import asyncio
import json
import sys
import os
import logging

# Add package path
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent

from docuflow_mcp.tools import get_all_tools
from docuflow_mcp.core.registry import dispatch_tool
from docuflow_mcp.core.middleware import (
    LoggingMiddleware,
    PerformanceMiddleware,
    ErrorHandlingMiddleware,
    add_middleware
)

# Import all operation classes to trigger decorator registration
from docuflow_mcp.document import (
    DocumentOperations,
    ParagraphOperations,
    HeadingOperations,
    TableOperations,
    ImageOperations,
    ListOperations,
    PageOperations,
    HeaderFooterOperations,
    SearchOperations,
    SpecialOperations,
    ExportOperations,
    CommentOperations
)
from docuflow_mcp.extensions.templates import TemplateManager
from docuflow_mcp.extensions.styles import StyleManager
from docuflow_mcp.extensions.converter import ConverterOperations
from docuflow_mcp.extensions.ocr import OCROperations
from docuflow_mcp.extensions.excel import ExcelOperations
from docuflow_mcp.extensions.pdf import PDFOperations
from docuflow_mcp.extensions.ppt import PPTOperations
from docuflow_mcp.extensions.image_gen import ImageGenOperations


# Create server instance
server = Server("docuflow")


# ============================================================
# Initialize middleware
# ============================================================

# 1. Error handling middleware (bottom layer, executed last)
add_middleware(ErrorHandlingMiddleware())

# 2. Performance monitoring middleware (slow query threshold: 1 second)
add_middleware(PerformanceMiddleware(slow_threshold=1.0))

# 3. Logging middleware (top layer, executed first)
# Log file saved to logs/docuflow.log
log_file = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "logs", "docuflow.log")
add_middleware(LoggingMiddleware(log_file=log_file, log_level=logging.INFO))


# ============================================================
# MCP Server Endpoints
# ============================================================


@server.list_tools()
async def list_tools():
    """Return all available tools"""
    return get_all_tools()


@server.call_tool()
async def call_tool(name: str, arguments: dict):
    """Handle tool calls (using new registry system)"""
    try:
        arguments = arguments or {}
        result = dispatch_tool(name, arguments)
        return [TextContent(type="text", text=json.dumps(result, ensure_ascii=False, indent=2))]
    except FileNotFoundError as e:
        return [TextContent(type="text", text=json.dumps({
            "success": False,
            "error": f"File not found: {str(e)}",
            "error_code": "FILE_NOT_FOUND"
        }, ensure_ascii=False))]
    except PermissionError as e:
        return [TextContent(type="text", text=json.dumps({
            "success": False,
            "error": f"Permission denied: {str(e)}",
            "error_code": "PERMISSION_DENIED"
        }, ensure_ascii=False))]
    except Exception as e:
        return [TextContent(type="text", text=json.dumps({
            "success": False,
            "error": f"Operation failed: {str(e)}",
            "error_type": type(e).__name__
        }, ensure_ascii=False))]


async def main():
    """Start MCP server"""
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options()
        )


def run():
    """Entry function"""
    asyncio.run(main())


if __name__ == "__main__":
    run()
