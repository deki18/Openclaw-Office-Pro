"""
Office Pro - Exception System and Error Codes

OpenClaw Skill standard error codes and custom exceptions
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional


class ErrorCode:
    """OpenClaw Skill Standard Error Codes"""
    
    SUCCESS = "SKILL_000"
    
    INVALID_PARAMS = "SKILL_101"
    MISSING_REQUIRED_PARAM = "SKILL_102"
    INVALID_TEMPLATE_NAME = "SKILL_103"
    INVALID_OUTPUT_PATH = "SKILL_104"
    INVALID_DATA_FORMAT = "SKILL_105"
    
    TEMPLATE_NOT_FOUND = "SKILL_201"
    TEMPLATE_INVALID = "SKILL_202"
    TEMPLATE_RENDER_ERROR = "SKILL_203"
    
    DATA_PARSE_ERROR = "SKILL_301"
    DATA_VALIDATION_ERROR = "SKILL_302"
    DATA_FILE_NOT_FOUND = "SKILL_303"
    DATA_ENCODING_ERROR = "SKILL_304"
    
    FILE_NOT_FOUND = "SKILL_401"
    FILE_ACCESS_DENIED = "SKILL_402"
    FILE_WRITE_ERROR = "SKILL_403"
    PATH_TRAVERSAL_DETECTED = "SKILL_404"
    
    DEPENDENCY_MISSING = "SKILL_501"
    DEPENDENCY_VERSION_MISMATCH = "SKILL_502"
    
    INTERNAL_ERROR = "SKILL_999"
    NOT_IMPLEMENTED = "SKILL_998"


class OfficeProError(Exception):
    """Base exception for Office Pro"""
    
    def __init__(self, message: str, error_code: str = ErrorCode.INTERNAL_ERROR):
        self.message = message
        self.error_code = error_code
        super().__init__(self.message)
    
    def to_dict(self) -> dict:
        return {
            "success": False,
            "error": self.message,
            "error_code": self.error_code
        }


class ParameterError(OfficeProError):
    """Parameter validation error"""
    
    def __init__(self, message: str, error_code: str = ErrorCode.INVALID_PARAMS):
        super().__init__(message, error_code)


class TemplateError(OfficeProError):
    """Template related error"""
    
    def __init__(self, message: str, error_code: str = ErrorCode.TEMPLATE_INVALID):
        super().__init__(message, error_code)


class TemplateNotFoundError(TemplateError):
    """Template file not found"""
    
    def __init__(self, template_path: str):
        super().__init__(
            f"Template not found: {template_path}",
            ErrorCode.TEMPLATE_NOT_FOUND
        )
        self.template_path = template_path


class TemplateRenderError(TemplateError):
    """Template rendering error"""
    
    def __init__(self, message: str):
        super().__init__(message, ErrorCode.TEMPLATE_RENDER_ERROR)


class DataError(OfficeProError):
    """Data processing error"""
    
    def __init__(self, message: str, error_code: str = ErrorCode.DATA_PARSE_ERROR):
        super().__init__(message, error_code)


class DataFileNotFoundError(DataError):
    """Data file not found"""
    
    def __init__(self, file_path: str):
        super().__init__(
            f"Data file not found: {file_path}",
            ErrorCode.DATA_FILE_NOT_FOUND
        )
        self.file_path = file_path


class DataParseError(DataError):
    """Data parsing error (JSON, CSV, etc.)"""
    
    def __init__(self, message: str, original_error: Optional[Exception] = None):
        super().__init__(message, ErrorCode.DATA_PARSE_ERROR)
        self.original_error = original_error


class DataEncodingError(DataError):
    """Data encoding error"""
    
    def __init__(self, file_path: str, encoding: str = "utf-8"):
        super().__init__(
            f"Encoding error reading {file_path}, expected {encoding}",
            ErrorCode.DATA_ENCODING_ERROR
        )


class FileError(OfficeProError):
    """File operation error"""
    
    def __init__(self, message: str, error_code: str = ErrorCode.FILE_WRITE_ERROR):
        super().__init__(message, error_code)


class PathTraversalError(FileError):
    """Path traversal attack detected"""
    
    def __init__(self, path: str):
        super().__init__(
            f"Path traversal detected: {path}",
            ErrorCode.PATH_TRAVERSAL_DETECTED
        )
        self.path = path


class FileAccessDeniedError(FileError):
    """File access denied"""
    
    def __init__(self, file_path: str):
        super().__init__(
            f"Access denied: {file_path}",
            ErrorCode.FILE_ACCESS_DENIED
        )


class DependencyError(OfficeProError):
    """Dependency error"""
    
    def __init__(self, message: str, error_code: str = ErrorCode.DEPENDENCY_MISSING):
        super().__init__(message, error_code)


class DocumentNotLoadedError(OfficeProError):
    """Document not loaded error"""
    
    def __init__(self, document_type: str = "document"):
        super().__init__(
            f"No {document_type} loaded. Call load_template() or create_document() first.",
            ErrorCode.INTERNAL_ERROR
        )


def validate_safe_path(file_path: str, allowed_base: Optional[str] = None) -> Path:
    """
    Validate file path security
    
    Args:
        file_path: File path to validate
        allowed_base: Optional base directory that the path must be within
        
    Returns:
        Resolved Path object
        
    Raises:
        PathTraversalError: If path traversal is detected
        FileAccessDeniedError: If path is outside allowed directory
    """
    path = Path(file_path)
    
    if '..' in file_path:
        raise PathTraversalError(file_path)
    
    try:
        resolved_path = path.resolve()
    except Exception as e:
        raise FileError(f"Invalid path: {file_path} - {e}")
    
    if allowed_base:
        base = Path(allowed_base).resolve()
        try:
            resolved_path.relative_to(base)
        except ValueError:
            raise FileAccessDeniedError(file_path)
    
    return resolved_path


def validate_file_exists(file_path: str, file_type: str = "File") -> Path:
    """
    Validate that a file exists
    
    Args:
        file_path: File path to validate
        file_type: Type description for error messages
        
    Returns:
        Path object if file exists
        
    Raises:
        DataFileNotFoundError: If file does not exist
    """
    path = Path(file_path)
    if not path.exists():
        raise DataFileNotFoundError(f"{file_type} not found: {file_path}")
    if not path.is_file():
        raise DataFileNotFoundError(f"{file_type} is not a file: {file_path}")
    return path
