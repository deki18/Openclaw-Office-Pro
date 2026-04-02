"""
Office Pro - Utility Functions

Common utilities for file handling, data loading, and validation
"""

from __future__ import annotations

import json
import os
import hashlib
import time
from functools import lru_cache, wraps
from json import JSONDecodeError
from pathlib import Path
from typing import Any, Dict, Optional, Union, Callable, TypeVar, Generic

from .exceptions import (
    DataEncodingError,
    DataFileNotFoundError,
    DataParseError,
    PathTraversalError,
    validate_safe_path,
    validate_file_exists,
)

T = TypeVar('T')


class TemplateCache:
    """
    Simple template cache with TTL support
    
    Caches template file metadata and modification times to avoid
    repeated filesystem checks.
    """
    _instance: Optional['TemplateCache'] = None
    _cache: Dict[str, Dict[str, Any]] = {}
    _ttl: int = 300
    
    def __new__(cls) -> 'TemplateCache':
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def get(self, key: str) -> Optional[Any]:
        """Get cached value if not expired"""
        if key not in self._cache:
            return None
        
        entry = self._cache[key]
        if time.time() - entry['timestamp'] > self._ttl:
            del self._cache[key]
            return None
        
        return entry['value']
    
    def set(self, key: str, value: Any) -> None:
        """Set cache value with current timestamp"""
        self._cache[key] = {
            'value': value,
            'timestamp': time.time()
        }
    
    def invalidate(self, key: str) -> None:
        """Invalidate a specific cache entry"""
        self._cache.pop(key, None)
    
    def clear(self) -> None:
        """Clear all cache entries"""
        self._cache.clear()
    
    @staticmethod
    def get_file_hash(file_path: Union[str, Path]) -> str:
        """Get MD5 hash of file for cache key"""
        path = Path(file_path)
        if not path.exists():
            return ""
        
        hasher = hashlib.md5()
        hasher.update(str(path.resolve()).encode())
        hasher.update(str(path.stat().st_mtime).encode())
        return hasher.hexdigest()


def cached_file_read(
    ttl: int = 300,
    key_func: Optional[Callable[..., str]] = None
) -> Callable[[Callable[..., T]], Callable[..., T]]:
    """
    Decorator for caching file read operations
    
    Args:
        ttl: Time-to-live in seconds
        key_func: Optional function to generate cache key
        
    Returns:
        Decorated function
    """
    cache: Dict[str, Dict[str, Any]] = {}
    
    def decorator(func: Callable[..., T]) -> Callable[..., T]:
        @wraps(func)
        def wrapper(*args, **kwargs) -> T:
            if key_func:
                cache_key = key_func(*args, **kwargs)
            else:
                cache_key = str(args[0]) if args else str(kwargs.get('file_path', ''))
            
            now = time.time()
            
            if cache_key in cache:
                entry = cache[cache_key]
                if now - entry['timestamp'] < ttl:
                    return entry['value']
            
            result = func(*args, **kwargs)
            cache[cache_key] = {
                'value': result,
                'timestamp': now
            }
            
            return result
        
        wrapper.cache_clear = lambda: cache.clear()
        return wrapper
    
    return decorator


@lru_cache(maxsize=128)
def get_cached_template_list(template_dir: str, template_type: str) -> tuple:
    """
    Get cached list of templates in a directory
    
    Args:
        template_dir: Directory to search
        template_type: Type of template ('word' or 'excel')
        
    Returns:
        Tuple of template filenames
    """
    dir_path = Path(template_dir)
    if not dir_path.exists():
        return ()
    
    if template_type == 'word':
        pattern = '*.docx'
    elif template_type == 'excel':
        pattern = '*.xlsx'
    else:
        pattern = '*'
    
    return tuple(sorted(p.name for p in dir_path.glob(pattern)))


def invalidate_template_cache() -> None:
    """Invalidate the template list cache"""
    get_cached_template_list.cache_clear()


def load_json_file(
    file_path: Union[str, Path],
    encoding: str = "utf-8",
    allowed_base: Optional[str] = None,
    validate_path: bool = True
) -> Dict[str, Any]:
    """
    Safely load a JSON file with comprehensive error handling
    
    Args:
        file_path: Path to JSON file
        encoding: File encoding (default: utf-8)
        allowed_base: Optional base directory restriction
        validate_path: Whether to validate path security
        
    Returns:
        Parsed JSON data as dictionary
        
    Raises:
        DataFileNotFoundError: If file does not exist
        DataParseError: If JSON parsing fails
        DataEncodingError: If file encoding is incorrect
        PathTraversalError: If path traversal is detected
    """
    if isinstance(file_path, Path):
        path = file_path
    else:
        path = Path(file_path)
    
    if validate_path:
        path = validate_safe_path(str(path), allowed_base)
    
    if not path.exists():
        raise DataFileNotFoundError(str(path))
    
    if not path.is_file():
        raise DataFileNotFoundError(f"Not a file: {path}")
    
    try:
        with open(path, 'r', encoding=encoding) as f:
            return json.load(f)
    except JSONDecodeError as e:
        raise DataParseError(
            f"Invalid JSON format in {path}: {e.msg} at line {e.lineno}, column {e.colno}",
            original_error=e
        )
    except UnicodeDecodeError as e:
        raise DataEncodingError(str(path), encoding)
    except PermissionError:
        raise DataFileNotFoundError(f"Permission denied: {path}")


def safe_resolve_path(
    file_path: str,
    base_dir: Optional[str] = None,
    must_exist: bool = False
) -> Path:
    """
    Safely resolve a file path
    
    Args:
        file_path: File path to resolve
        base_dir: Optional base directory
        must_exist: Whether the file must exist
        
    Returns:
        Resolved Path object
    """
    path = Path(file_path)
    
    if not path.is_absolute() and base_dir:
        path = Path(base_dir) / path
    
    try:
        resolved = path.resolve()
    except Exception:
        resolved = path
    
    if must_exist and not resolved.exists():
        raise DataFileNotFoundError(str(resolved))
    
    return resolved


def ensure_directory(dir_path: Union[str, Path]) -> Path:
    """
    Ensure a directory exists, create if necessary
    
    Args:
        dir_path: Directory path
        
    Returns:
        Path object for the directory
    """
    path = Path(dir_path)
    path.mkdir(parents=True, exist_ok=True)
    return path


def get_output_path(
    output_path: str,
    default_dir: Optional[str] = None,
    default_name: Optional[str] = None
) -> Path:
    """
    Get output file path with defaults
    
    Args:
        output_path: Specified output path
        default_dir: Default directory if not specified
        default_name: Default filename if not specified
        
    Returns:
        Resolved output Path
    """
    path = Path(output_path)
    
    if not path.is_absolute() and default_dir:
        path = Path(default_dir) / path
    
    if not path.suffix and default_name:
        path = path / default_name
    
    return path


def format_file_size(size_bytes: int) -> str:
    """Format file size in human readable format"""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def get_template_dir(template_type: str = "word") -> Path:
    """
    Get default template directory
    
    Args:
        template_type: Type of template ('word' or 'excel')
        
    Returns:
        Path to template directory
    """
    skill_root = Path(__file__).parent.parent
    return skill_root / "assets" / "templates" / template_type


def validate_template_path(
    template_name: str,
    template_type: str = "word",
    template_dir: Optional[str] = None
) -> Path:
    """
    Validate and resolve template path
    
    Args:
        template_name: Template filename
        template_type: Type of template ('word' or 'excel')
        template_dir: Custom template directory
        
    Returns:
        Resolved template Path
        
    Raises:
        DataFileNotFoundError: If template not found
    """
    if template_dir:
        base_dir = Path(template_dir)
    else:
        base_dir = get_template_dir(template_type)
    
    template_path = base_dir / template_name
    
    if not template_path.exists():
        raise DataFileNotFoundError(f"Template not found: {template_path}")
    
    return template_path
