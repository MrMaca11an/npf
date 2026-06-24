"""Слой источников данных: абстракция над файловой системой, API и т.д."""
from .base import RawDocument, Source
from .filesystem import FileSystemSource

__all__ = ["RawDocument", "Source", "FileSystemSource"]
