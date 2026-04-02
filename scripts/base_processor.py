"""
Office Pro - Base Processor Module

Abstract base class for document processors
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from functools import wraps
from pathlib import Path
from typing import Any, Dict, Generic, Optional, TypeVar

from .exceptions import DocumentNotLoadedError, TemplateNotFoundError

DocumentType = TypeVar('DocumentType')
TemplateType = TypeVar('TemplateType')


def require_document(func):
    """
    Decorator to ensure document is loaded before method execution
    
    Usage:
        @require_document
        def some_method(self, ...):
            # self._document is guaranteed to be not None
    """
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        if not self._document:
            raise DocumentNotLoadedError(self._document_type_name)
        return func(self, *args, **kwargs)
    return wrapper


def require_template(func):
    """
    Decorator to ensure template is loaded before method execution
    """
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        if not self._template:
            raise DocumentNotLoadedError("template")
        return func(self, *args, **kwargs)
    return wrapper


class DocumentProcessor(ABC, Generic[DocumentType, TemplateType]):
    """
    Abstract base class for document processors
    
    Provides common interface for Word and Excel processing
    """
    
    _document_type_name: str = "document"
    
    def __init__(self, template_dir: Optional[str] = None):
        """
        Initialize processor
        
        Args:
            template_dir: Custom template directory path
        """
        self.template_dir = template_dir or self._get_default_template_dir()
        self._document: Optional[DocumentType] = None
        self._template: Optional[TemplateType] = None
    
    @abstractmethod
    def _get_default_template_dir(self) -> str:
        """Get default template directory"""
        pass
    
    @abstractmethod
    def create_document(self) -> DocumentType:
        """Create new document"""
        pass
    
    @abstractmethod
    def load_document(self, path: str) -> DocumentType:
        """Load existing document"""
        pass
    
    @abstractmethod
    def load_template(self, template_name: str) -> TemplateType:
        """Load template file"""
        pass
    
    @abstractmethod
    def render_template(self, data: Dict[str, Any]) -> DocumentType:
        """Render template with data"""
        pass
    
    @abstractmethod
    def save(self, path: str) -> None:
        """Save document to file"""
        pass
    
    def render_and_save(self, data: Dict[str, Any], output_path: str) -> str:
        """
        Render template and save to file
        
        Args:
            data: Template data dictionary
            output_path: Output file path
            
        Returns:
            Saved file path
        """
        self.render_template(data)
        self.save(output_path)
        return output_path
    
    def _validate_template_path(self, template_name: str) -> Path:
        """
        Validate template path exists
        
        Args:
            template_name: Template filename
            
        Returns:
            Resolved template path
            
        Raises:
            TemplateNotFoundError: If template not found
        """
        template_path = Path(self.template_dir) / template_name
        if not template_path.exists():
            raise TemplateNotFoundError(str(template_path))
        return template_path
    
    def _ensure_output_dir(self, path: str) -> None:
        """Ensure output directory exists"""
        Path(path).parent.mkdir(parents=True, exist_ok=True)
    
    @property
    def document(self) -> Optional[DocumentType]:
        """Get current document object"""
        return self._document
    
    @property
    def template(self) -> Optional[TemplateType]:
        """Get current template object"""
        return self._template
    
    def __enter__(self) -> 'DocumentProcessor':
        """Context manager entry"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        """Context manager exit - cleanup resources"""
        self._document = None
        self._template = None
        return False
