"""DocuFlow MCP - Extensions module"""

# Import all extension modules to trigger tool registration
from . import templates
from . import styles
from . import batch
from . import validator
from . import advanced
from . import image_gen
from . import html_to_pptx

__all__ = ["templates", "styles", "batch", "validator", "advanced", "image_gen", "html_to_pptx"]

