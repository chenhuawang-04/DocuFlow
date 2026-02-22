"""
DocuFlow MCP - 工具注册与分发系统

实现工具的自动注册和O(1)时间复杂度的分发机制
"""
from typing import Callable, Dict, Any, List, Optional
from functools import wraps


# 全局工具注册表
_TOOL_REGISTRY: Dict[str, Dict[str, Any]] = {}

# 中间件管理器（延迟导入避免循环依赖）
_middleware_manager = None


def _get_middleware_manager():
    """获取中间件管理器（延迟导入）"""
    global _middleware_manager
    if _middleware_manager is None:
        from .middleware import get_middleware_manager
        _middleware_manager = get_middleware_manager()
    return _middleware_manager


def register_tool(
    name: str,
    required_params: Optional[List[str]] = None,
    optional_params: Optional[List[str]] = None
):
    """
    工具注册装饰器

    用法:
        @register_tool("doc_create", required_params=["path"], optional_params=["title", "template"])
        @staticmethod
        def create(path, title=None, template=None, preset_template=None):
            ...

    Args:
        name: 工具名称
        required_params: 必需参数列表
        optional_params: 可选参数列表

    Returns:
        装饰器函数
    """
    def decorator(func: Callable) -> Callable:
        _TOOL_REGISTRY[name] = {
            "handler": func,
            "required_params": required_params or [],
            "optional_params": optional_params or [],
            "function_name": func.__name__
        }

        @wraps(func)
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)

        return wrapper
    return decorator


def dispatch_tool(name: str, args: dict) -> dict:
    """
    工具分发函数（O(1)时间复杂度）

    集成中间件链，支持日志、性能监控、异常处理等

    Args:
        name: 工具名称
        args: 工具参数字典

    Returns:
        工具执行结果
    """
    # 检查工具是否存在
    if name not in _TOOL_REGISTRY:
        return {
            "success": False,
            "error": f"未知工具: {name}",
            "error_code": "TOOL_NOT_FOUND"
        }

    tool_info = _TOOL_REGISTRY[name]
    handler = tool_info["handler"]
    required = tool_info["required_params"]

    # 验证必需参数
    missing = [p for p in required if p not in args]
    if missing:
        return {
            "success": False,
            "error": f"缺少必需参数: {', '.join(missing)}",
            "error_code": "MISSING_PARAMS",
            "missing_params": missing
        }

    # 获取中间件管理器并执行
    middleware_manager = _get_middleware_manager()

    # 如果没有中间件，直接执行
    if not middleware_manager.middlewares:
        try:
            return handler(**args)
        except TypeError as e:
            return {
                "success": False,
                "error": f"参数错误: {str(e)}",
                "error_code": "INVALID_PARAMS"
            }
        except Exception as e:
            return {
                "success": False,
                "error": f"执行错误: {str(e)}",
                "error_type": type(e).__name__
            }

    # 通过中间件链执行
    return middleware_manager.execute(name, args, handler)


def get_all_registered_tools() -> List[str]:
    """
    获取所有已注册的工具名称

    Returns:
        工具名称列表
    """
    return list(_TOOL_REGISTRY.keys())


def get_tool_info(name: str) -> Dict[str, Any]:
    """
    获取工具的详细信息

    Args:
        name: 工具名称

    Returns:
        工具信息字典，如果工具不存在则返回空字典
    """
    return _TOOL_REGISTRY.get(name, {})


def clear_registry():
    """
    清空注册表（主要用于测试）
    """
    _TOOL_REGISTRY.clear()
