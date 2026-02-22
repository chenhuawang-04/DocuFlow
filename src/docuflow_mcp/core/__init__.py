"""
DocuFlow MCP - Core module

核心模块：工具注册、中间件、配置管理
"""

from .registry import register_tool, dispatch_tool, get_all_registered_tools, get_tool_info
from .middleware import (
    Middleware,
    LoggingMiddleware,
    PerformanceMiddleware,
    ErrorHandlingMiddleware,
    ValidationMiddleware,
    MiddlewareManager,
    get_middleware_manager,
    add_middleware,
    clear_middlewares
)
from .config import Config, get_config, get, set, get_section

__all__ = [
    # 工具注册
    'register_tool',
    'dispatch_tool',
    'get_all_registered_tools',
    'get_tool_info',
    # 中间件
    'Middleware',
    'LoggingMiddleware',
    'PerformanceMiddleware',
    'ErrorHandlingMiddleware',
    'ValidationMiddleware',
    'MiddlewareManager',
    'get_middleware_manager',
    'add_middleware',
    'clear_middlewares',
    # 配置管理
    'Config',
    'get_config',
    'get',
    'set',
    'get_section'
]
