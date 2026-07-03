"""
Реализация источника данных: чтение файлов с файловой системы.

При переходе на API этот модуль заменяется на ApiSource, реализующий
тот же интерфейс Source, — бизнес-логика не меняется.
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Literal

from .base import RawDocument, Source

logger = logging.getLogger(__name__)

# Расширения, которые игнорируются при сканировании
_SKIP_EXTENSIONS = {".tmp", ".bak"}
# Префиксы временных файлов Excel (открытые в Office)
_SKIP_PREFIXES = ("~$",)
# Имена системных файлов, которые нужно пропустить
_SKIP_NAMES = {".ds_store", ".gitkeep", "thumbs.db"}

# Соответствие расширений формату
_EXT_TO_FMT: dict[str, Literal["excel", "json"]] = {
    ".xlsx": "excel",
    ".xls": "excel",
    ".json": "json",
}


class FileSystemSource(Source):
    """
    Источник данных: обходит папки ПУ и БУ на файловой системе.

    :param pu_dir: путь к директории персонифицированного учёта.
    :param bu_dir: путь к директории бухгалтерского учёта.
    """

    def __init__(self, pu_dir: Path, bu_dir: Path) -> None:
        self.pu_dir = pu_dir
        self.bu_dir = bu_dir

    def scan(self) -> list[RawDocument]:
        """
        Обходит директории ПУ и БУ, возвращает список RawDocument.

        Пропускаются временные файлы Excel (~$), системные файлы (.gitkeep и т.п.),
        а также файлы с неподдерживаемыми расширениями.

        :return: отсортированный список RawDocument.
        """
        docs: list[RawDocument] = []
        for side, directory in [("ПУ", self.pu_dir), ("БУ", self.bu_dir)]:
            if not directory.is_dir():
                logger.warning("Директория не найдена: %s", directory)
                continue
            for file_path in sorted(directory.rglob("*")):
                if not file_path.is_file():
                    continue
                name_lower = file_path.name.lower()
                # Пропускаем системные и временные файлы
                if name_lower in _SKIP_NAMES:
                    continue
                if any(name_lower.startswith(p) for p in _SKIP_PREFIXES):
                    logger.debug("Пропуск временного файла: %s", file_path.name)
                    continue
                ext = file_path.suffix.lower()
                if ext in _SKIP_EXTENSIONS or ext not in _EXT_TO_FMT:
                    logger.debug("Пропуск файла с неподдерживаемым расширением: %s", file_path.name)
                    continue
                fmt = _EXT_TO_FMT[ext]
                base_name = file_path.stem.lower()  # имя без расширения, строчные
                docs.append(
                    RawDocument(
                        path=file_path,
                        base_name=base_name,
                        ext=ext.lstrip("."),
                        fmt=fmt,
                        side=side,  # type: ignore[arg-type]
                    )
                )
                logger.debug("Обнаружен файл [%s]: %s", side, file_path.name)

        logger.info("Сканирование завершено: найдено %d файл(ов)", len(docs))
        return docs
