"""
DocuFlow MCP - Configuration Management

Global configuration management with environment variable support
"""

import os
import json
import logging
from typing import Dict, Any, Optional
from pathlib import Path


class Config:
    """Global configuration manager"""

    # Singleton instance
    _instance: Optional['Config'] = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        """Initialize configuration (only once due to singleton pattern)"""
        if self._initialized:
            return

        self._config: Dict[str, Any] = {}
        self._load_defaults()
        self._load_from_env()
        self._initialized = True

    def _load_defaults(self):
        """Load default configuration"""
        self._config = {
            # Logging settings
            'logging': {
                'enabled': True,
                'level': logging.INFO,
                'file': None,  # Will be set to logs/docuflow.log if enabled
                'console': True,
                'format': '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                'max_param_length': 200  # Maximum length for logged parameters
            },

            # Performance settings
            'performance': {
                'monitoring_enabled': True,
                'slow_threshold': 1.0,  # seconds
                'stats_enabled': True
            },

            # Error handling settings
            'error_handling': {
                'detailed_errors': True,  # Include error type and stack trace
                'error_codes': True  # Return standardized error codes
            },

            # Document settings
            'document': {
                'default_font': 'Arial',
                'default_font_size': '12pt',
                'default_line_spacing': 1.15,
                'default_page_width': '21cm',  # A4 width
                'default_page_height': '29.7cm',  # A4 height
                'default_margins': {
                    'top': '2.54cm',  # 1 inch
                    'bottom': '2.54cm',
                    'left': '2.54cm',
                    'right': '2.54cm'
                }
            },

            # Middleware settings
            'middleware': {
                'logging_middleware': True,
                'performance_middleware': True,
                'error_handling_middleware': True,
                'validation_middleware': False  # Disabled by default
            },

            # Template settings
            'templates': {
                'presets_file': 'templates/presets.json',
                'custom_templates_dir': 'templates/custom'
            },

            # Paths
            'paths': {
                'log_dir': 'logs',
                'template_dir': 'templates',
                'temp_dir': 'temp'
            }
        }

    def _load_from_env(self):
        """Load configuration from environment variables"""
        # Logging level from environment
        log_level_str = os.getenv('DOCUFLOW_LOG_LEVEL', '').upper()
        if log_level_str in ('DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'):
            self._config['logging']['level'] = getattr(logging, log_level_str)

        # Log file path
        log_file = os.getenv('DOCUFLOW_LOG_FILE')
        if log_file:
            self._config['logging']['file'] = log_file

        # Performance monitoring
        if os.getenv('DOCUFLOW_DISABLE_PERFORMANCE_MONITORING'):
            self._config['performance']['monitoring_enabled'] = False

        slow_threshold = os.getenv('DOCUFLOW_SLOW_THRESHOLD')
        if slow_threshold:
            try:
                self._config['performance']['slow_threshold'] = float(slow_threshold)
            except ValueError:
                pass

        # Document defaults
        default_font = os.getenv('DOCUFLOW_DEFAULT_FONT')
        if default_font:
            self._config['document']['default_font'] = default_font

        default_font_size = os.getenv('DOCUFLOW_DEFAULT_FONT_SIZE')
        if default_font_size:
            self._config['document']['default_font_size'] = default_font_size

    def get(self, key: str, default: Any = None) -> Any:
        """
        Get configuration value using dot notation

        Args:
            key: Configuration key in dot notation (e.g., 'logging.level')
            default: Default value if key not found

        Returns:
            Configuration value or default
        """
        keys = key.split('.')
        value = self._config

        for k in keys:
            if isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default

        return value

    def set(self, key: str, value: Any):
        """
        Set configuration value using dot notation

        Args:
            key: Configuration key in dot notation
            value: Value to set
        """
        keys = key.split('.')
        config = self._config

        # Navigate to the parent dict
        for k in keys[:-1]:
            if k not in config:
                config[k] = {}
            config = config[k]

        # Set the value
        config[keys[-1]] = value

    def get_section(self, section: str) -> Dict[str, Any]:
        """
        Get entire configuration section

        Args:
            section: Section name

        Returns:
            Section dictionary or empty dict if not found
        """
        return self._config.get(section, {})

    def update_section(self, section: str, values: Dict[str, Any]):
        """
        Update entire configuration section

        Args:
            section: Section name
            values: New values
        """
        if section in self._config:
            self._config[section].update(values)
        else:
            self._config[section] = values

    def load_from_file(self, file_path: str):
        """
        Load configuration from JSON file

        Args:
            file_path: Path to JSON configuration file
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            file_config = json.load(f)

        # Merge file config with existing config
        self._merge_config(self._config, file_config)

    def save_to_file(self, file_path: str):
        """
        Save configuration to JSON file

        Args:
            file_path: Path to output JSON file
        """
        # Ensure directory exists
        os.makedirs(os.path.dirname(file_path) or '.', exist_ok=True)

        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(self._config, f, indent=2, ensure_ascii=False)

    def _merge_config(self, base: dict, updates: dict):
        """Recursively merge configuration dictionaries"""
        for key, value in updates.items():
            if key in base and isinstance(base[key], dict) and isinstance(value, dict):
                self._merge_config(base[key], value)
            else:
                base[key] = value

    def reset(self):
        """Reset configuration to defaults"""
        self._load_defaults()
        self._load_from_env()

    def to_dict(self) -> Dict[str, Any]:
        """
        Get complete configuration as dictionary

        Returns:
            Complete configuration dictionary
        """
        return self._config.copy()

    def print_config(self):
        """Print current configuration (for debugging)"""
        import pprint
        print("=" * 60)
        print("DocuFlow Configuration")
        print("=" * 60)
        pprint.pprint(self._config)
        print("=" * 60)


# Global configuration instance
_config = Config()


def get_config() -> Config:
    """
    Get global configuration instance

    Returns:
        Global Config instance
    """
    return _config


def get(key: str, default: Any = None) -> Any:
    """
    Convenience function to get configuration value

    Args:
        key: Configuration key in dot notation
        default: Default value if not found

    Returns:
        Configuration value
    """
    return _config.get(key, default)


def set(key: str, value: Any):
    """
    Convenience function to set configuration value

    Args:
        key: Configuration key in dot notation
        value: Value to set
    """
    _config.set(key, value)


def get_section(section: str) -> Dict[str, Any]:
    """
    Convenience function to get configuration section

    Args:
        section: Section name

    Returns:
        Section dictionary
    """
    return _config.get_section(section)


# Example usage
if __name__ == "__main__":
    config = get_config()

    # Print current configuration
    config.print_config()

    # Get specific values
    print("\nGetting specific values:")
    print(f"Log level: {config.get('logging.level')}")
    print(f"Slow threshold: {config.get('performance.slow_threshold')}")
    print(f"Default font: {config.get('document.default_font')}")

    # Set values
    print("\nSetting new values:")
    config.set('logging.level', logging.DEBUG)
    config.set('performance.slow_threshold', 0.5)
    print(f"New log level: {config.get('logging.level')}")
    print(f"New slow threshold: {config.get('performance.slow_threshold')}")

    # Get entire section
    print("\nGetting entire section:")
    logging_config = config.get_section('logging')
    print(f"Logging config: {logging_config}")

    # Save to file
    print("\nSaving configuration:")
    config.save_to_file('config.json')
    print("Configuration saved to config.json")
