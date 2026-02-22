"""
DocuFlow MCP - Middleware System
中间件系统：日志、性能监控、异常处理
"""

import time
import logging
import json
from typing import Callable, Dict, Any, List, Optional
from functools import wraps
from datetime import datetime
from pathlib import Path


# ============================================================
# 中间件基类
# ============================================================

class Middleware:
    """中间件基类"""

    def __init__(self):
        self.enabled = True

    def before(self, tool_name: str, args: dict) -> Optional[dict]:
        """
        工具调用前执行

        Args:
            tool_name: 工具名称
            args: 工具参数

        Returns:
            如果返回dict，则直接返回该结果，不继续执行工具
            如果返回None，则继续执行
        """
        return None

    def after(self, tool_name: str, args: dict, result: dict, elapsed_time: float) -> dict:
        """
        工具调用后执行

        Args:
            tool_name: 工具名称
            args: 工具参数
            result: 工具执行结果
            elapsed_time: 执行时间（秒）

        Returns:
            修改后的结果（或原结果）
        """
        return result

    def on_error(self, tool_name: str, args: dict, error: Exception) -> dict:
        """
        工具执行出错时调用

        Args:
            tool_name: 工具名称
            args: 工具参数
            error: 异常对象

        Returns:
            错误响应字典
        """
        return {
            "success": False,
            "error": str(error),
            "error_type": type(error).__name__
        }


# ============================================================
# 日志中间件
# ============================================================

class LoggingMiddleware(Middleware):
    """日志中间件 - 记录所有工具调用"""

    # 需要脱敏的参数名（包含即匹配）
    SENSITIVE_KEYS = {'api_key', 'password', 'token', 'secret', 'credential', 'authorization'}

    def __init__(self, log_file: Optional[str] = None, log_level: int = logging.INFO):
        super().__init__()

        # 配置日志
        self.logger = logging.getLogger("DocuFlow.MCP")
        self.logger.setLevel(log_level)

        # 避免重复添加handler
        if not self.logger.handlers:
            # 控制台输出
            console_handler = logging.StreamHandler()
            console_handler.setLevel(log_level)
            console_formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            console_handler.setFormatter(console_formatter)
            self.logger.addHandler(console_handler)

            # 文件输出（如果指定）
            if log_file:
                # 确保日志目录存在
                log_path = Path(log_file)
                log_path.parent.mkdir(parents=True, exist_ok=True)

                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_handler.setLevel(log_level)
                file_formatter = logging.Formatter(
                    '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
                )
                file_handler.setFormatter(file_formatter)
                self.logger.addHandler(file_handler)

        self.log_params = True  # 是否记录参数
        self.log_results = True  # 是否记录结果
        self.max_param_length = 200  # 参数最大长度

    @staticmethod
    def _sanitize_args(args: dict) -> dict:
        """对敏感参数进行脱敏处理"""
        sanitized = {}
        for key, value in args.items():
            if any(s in key.lower() for s in LoggingMiddleware.SENSITIVE_KEYS):
                if isinstance(value, str) and len(value) > 4:
                    sanitized[key] = value[:2] + "***" + value[-2:]
                else:
                    sanitized[key] = "***"
            elif isinstance(value, dict):
                sanitized[key] = LoggingMiddleware._sanitize_args(value)
            else:
                sanitized[key] = value
        return sanitized

    def before(self, tool_name: str, args: dict) -> Optional[dict]:
        """记录工具调用"""
        if self.log_params:
            # 脱敏后截断过长的参数
            safe_args = self._sanitize_args(args)
            args_str = json.dumps(safe_args, ensure_ascii=False)
            if len(args_str) > self.max_param_length:
                args_str = args_str[:self.max_param_length] + "..."
            self.logger.info(f"Tool called: {tool_name} | Args: {args_str}")
        else:
            self.logger.info(f"Tool called: {tool_name}")

        return None

    def after(self, tool_name: str, args: dict, result: dict, elapsed_time: float) -> dict:
        """记录工具执行结果"""
        success = result.get("success", True)

        if success:
            if self.log_results:
                result_str = json.dumps(result, ensure_ascii=False)
                if len(result_str) > self.max_param_length:
                    result_str = result_str[:self.max_param_length] + "..."
                self.logger.info(
                    f"Tool completed: {tool_name} | Time: {elapsed_time*1000:.2f}ms | Result: {result_str}"
                )
            else:
                self.logger.info(
                    f"Tool completed: {tool_name} | Time: {elapsed_time*1000:.2f}ms"
                )
        else:
            error_msg = result.get("error", "Unknown error")
            self.logger.error(
                f"Tool failed: {tool_name} | Time: {elapsed_time*1000:.2f}ms | Error: {error_msg}"
            )

        return result

    def on_error(self, tool_name: str, args: dict, error: Exception) -> dict:
        """记录异常"""
        self.logger.error(
            f"Tool error: {tool_name} | Exception: {type(error).__name__}: {str(error)}",
            exc_info=True
        )
        return super().on_error(tool_name, args, error)


# ============================================================
# 性能监控中间件
# ============================================================

class PerformanceMiddleware(Middleware):
    """性能监控中间件 - 记录执行时间和慢查询"""

    def __init__(self, slow_threshold: float = 1.0):
        """
        Args:
            slow_threshold: 慢查询阈值（秒），超过此时间将被记录
        """
        super().__init__()
        self.slow_threshold = slow_threshold
        self.stats: Dict[str, Dict[str, Any]] = {}  # 工具统计信息
        self.logger = logging.getLogger("DocuFlow.Performance")
        self.logger.setLevel(logging.WARNING)

        if not self.logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)

    def after(self, tool_name: str, args: dict, result: dict, elapsed_time: float) -> dict:
        """记录性能数据"""
        # 更新统计信息
        if tool_name not in self.stats:
            self.stats[tool_name] = {
                "count": 0,
                "total_time": 0.0,
                "min_time": float('inf'),
                "max_time": 0.0,
                "slow_count": 0
            }

        stats = self.stats[tool_name]
        stats["count"] += 1
        stats["total_time"] += elapsed_time
        stats["min_time"] = min(stats["min_time"], elapsed_time)
        stats["max_time"] = max(stats["max_time"], elapsed_time)

        # 检查是否为慢查询
        if elapsed_time > self.slow_threshold:
            stats["slow_count"] += 1
            safe_args = LoggingMiddleware._sanitize_args(args)
            args_str = json.dumps(safe_args, ensure_ascii=False)
            if len(args_str) > 200:
                args_str = args_str[:200] + "..."

            self.logger.warning(
                f"SLOW TOOL: {tool_name} took {elapsed_time*1000:.2f}ms (threshold: {self.slow_threshold*1000:.0f}ms) | Args: {args_str}"
            )

        # 将性能数据添加到结果中（可选）
        if not result.get("_performance"):
            result["_performance"] = {
                "elapsed_ms": round(elapsed_time * 1000, 2),
                "tool_name": tool_name
            }

        return result

    def get_stats(self) -> Dict[str, Dict[str, Any]]:
        """获取性能统计信息"""
        # 计算平均时间
        stats_with_avg = {}
        for tool_name, stats in self.stats.items():
            stats_copy = stats.copy()
            if stats_copy["count"] > 0:
                stats_copy["avg_time"] = stats_copy["total_time"] / stats_copy["count"]
            else:
                stats_copy["avg_time"] = 0.0
            stats_with_avg[tool_name] = stats_copy

        return stats_with_avg

    def reset_stats(self):
        """重置统计信息"""
        self.stats.clear()


# ============================================================
# 异常处理中间件
# ============================================================

class ErrorHandlingMiddleware(Middleware):
    """异常处理中间件 - 统一异常捕获和格式化"""

    def __init__(self):
        super().__init__()
        self.logger = logging.getLogger("DocuFlow.ErrorHandler")
        self.logger.setLevel(logging.ERROR)

        if not self.logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)

    def on_error(self, tool_name: str, args: dict, error: Exception) -> dict:
        """统一格式化错误响应"""
        error_type = type(error).__name__
        error_msg = str(error)

        # 根据异常类型返回不同的错误码
        if isinstance(error, FileNotFoundError):
            return {
                "success": False,
                "error": f"文件不存在: {error_msg}",
                "error_code": "FILE_NOT_FOUND",
                "error_type": error_type,
                "tool_name": tool_name
            }
        elif isinstance(error, PermissionError):
            return {
                "success": False,
                "error": f"权限不足: {error_msg}",
                "error_code": "PERMISSION_DENIED",
                "error_type": error_type,
                "tool_name": tool_name
            }
        elif isinstance(error, ValueError):
            return {
                "success": False,
                "error": f"参数值错误: {error_msg}",
                "error_code": "INVALID_VALUE",
                "error_type": error_type,
                "tool_name": tool_name
            }
        elif isinstance(error, KeyError):
            return {
                "success": False,
                "error": f"缺少必需的键: {error_msg}",
                "error_code": "MISSING_KEY",
                "error_type": error_type,
                "tool_name": tool_name
            }
        elif isinstance(error, TypeError):
            return {
                "success": False,
                "error": f"参数类型错误: {error_msg}",
                "error_code": "TYPE_ERROR",
                "error_type": error_type,
                "tool_name": tool_name
            }
        else:
            # 通用错误
            self.logger.error(
                f"Unhandled error in {tool_name}: {error_type}: {error_msg}",
                exc_info=True
            )
            return {
                "success": False,
                "error": f"操作失败: {error_msg}",
                "error_code": "INTERNAL_ERROR",
                "error_type": error_type,
                "tool_name": tool_name
            }


# ============================================================
# 参数验证中间件
# ============================================================

class ValidationMiddleware(Middleware):
    """参数验证中间件"""

    def __init__(self):
        super().__init__()
        self.validators: Dict[str, Callable] = {}

    def register_validator(self, tool_name: str, validator: Callable[[dict], Optional[str]]):
        """
        为工具注册验证器

        Args:
            tool_name: 工具名称
            validator: 验证函数，接收args字典，返回错误信息（如果验证失败）或None（验证成功）
        """
        self.validators[tool_name] = validator

    def before(self, tool_name: str, args: dict) -> Optional[dict]:
        """执行参数验证"""
        if tool_name in self.validators:
            error_msg = self.validators[tool_name](args)
            if error_msg:
                return {
                    "success": False,
                    "error": error_msg,
                    "error_code": "VALIDATION_ERROR",
                    "tool_name": tool_name
                }

        return None


# ============================================================
# 中间件管理器
# ============================================================

class MiddlewareManager:
    """中间件管理器 - 管理中间件链"""

    def __init__(self):
        self.middlewares: List[Middleware] = []

    def add(self, middleware: Middleware):
        """添加中间件"""
        self.middlewares.append(middleware)

    def remove(self, middleware: Middleware):
        """移除中间件"""
        if middleware in self.middlewares:
            self.middlewares.remove(middleware)

    def clear(self):
        """清空所有中间件"""
        self.middlewares.clear()

    def execute(self, tool_name: str, args: dict, handler: Callable) -> dict:
        """
        执行中间件链

        Args:
            tool_name: 工具名称
            args: 工具参数
            handler: 实际的工具处理函数

        Returns:
            工具执行结果
        """
        # 执行 before 钩子
        for middleware in self.middlewares:
            if not middleware.enabled:
                continue

            result = middleware.before(tool_name, args)
            if result is not None:
                # 中间件返回了结果，直接返回，不执行工具
                return result

        # 执行工具并计时
        start_time = time.time()
        error = None
        result = None

        try:
            result = handler(**args)
        except Exception as e:
            error = e

        elapsed_time = time.time() - start_time

        # 如果有错误，执行 on_error 钩子
        if error:
            for middleware in reversed(self.middlewares):
                if not middleware.enabled:
                    continue
                result = middleware.on_error(tool_name, args, error)
                # 第一个处理错误的中间件决定返回值
                break

            # 如果没有中间件处理，使用默认错误响应
            if result is None:
                result = {
                    "success": False,
                    "error": str(error),
                    "error_type": type(error).__name__
                }

        # 执行 after 钩子
        for middleware in reversed(self.middlewares):
            if not middleware.enabled:
                continue
            result = middleware.after(tool_name, args, result, elapsed_time)

        return result


# ============================================================
# 全局中间件管理器实例
# ============================================================

_global_middleware_manager = MiddlewareManager()


def get_middleware_manager() -> MiddlewareManager:
    """获取全局中间件管理器"""
    return _global_middleware_manager


def add_middleware(middleware: Middleware):
    """添加中间件到全局管理器"""
    _global_middleware_manager.add(middleware)


def clear_middlewares():
    """清空全局中间件"""
    _global_middleware_manager.clear()
