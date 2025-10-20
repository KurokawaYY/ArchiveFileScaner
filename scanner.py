import os
import zipfile
import py7zr
import rarfile
from pathlib import Path
import logging
from typing import List, Dict, Any, Callable
import threading
import binascii
import tempfile
import subprocess
import sys
import re


class ArchiveScanner:
    def __init__(self):
        self.supported_formats = {'.zip', '.7z', '.rar'}
        self.logger = self._setup_logger()
        self._stop_event = threading.Event()
        self._progress_callback = None
        self._total_files = 0
        self._processed_files = 0
        self.max_archive_depth = 5
        self._7zip_path = self._find_7zip()
        self.root_path = None

    def _setup_logger(self):
        logging.basicConfig(level=logging.INFO)
        return logging.getLogger(__name__)

    def _find_7zip(self):
        """Находит путь к 7z.exe"""
        possible_paths = [
            "7z", "7z.exe",
            r"C:\Program Files\7-Zip\7z.exe",
            r"C:\Program Files (x86)\7-Zip\7z.exe",
        ]
        for path in possible_paths:
            if os.path.exists(path):
                return path
        return None

    def stop_scanning(self):
        self._stop_event.set()

    def set_progress_callback(self, callback: Callable):
        self._progress_callback = callback

    def _update_progress(self):
        if self._progress_callback and self._total_files > 0:
            progress = (self._processed_files / self._total_files) * 100
            self._progress_callback(progress)

    def _calculate_file_crc32(self, file_path: Path, chunk_size=65536):
        try:
            if not file_path.exists():
                return None
            file_size = file_path.stat().st_size
            if file_size == 0:
                return 0
            crc = 0
            with open(file_path, "rb") as f:
                while chunk := f.read(chunk_size):
                    if self._stop_event.is_set():
                        return None
                    crc = binascii.crc32(chunk, crc)
            return crc & 0xFFFFFFFF
        except Exception as e:
            self.logger.warning(f"Не удалось вычислить CRC32 для {file_path}: {e}")
            return None

    def _decode_filename(self, filename):
        """Декодирует имя файла с обработкой разных кодировок"""
        if isinstance(filename, str):
            return filename

        encodings = ['utf-8', 'cp1251', 'cp866', 'iso-8859-1', 'latin1']

        for encoding in encodings:
            try:
                return filename.decode(encoding)
            except (UnicodeDecodeError, AttributeError):
                continue

        try:
            return filename.decode('utf-8', errors='replace')
        except:
            return str(filename)

    def _decode_zip_filename(self, filename):
        """Специальное декодирование для ZIP архивов с улучшенной обработкой"""
        if isinstance(filename, str):
            if self._has_garbled_chars(filename):
                fixed = self._fix_garbled_text(filename)
                if not self._has_garbled_chars(fixed):
                    return fixed
            return filename

        # Пробуем разные стратегии декодирования
        decoded = self._try_multiple_decodings(filename)
        if decoded and not self._has_garbled_chars(decoded):
            return decoded

        # Если все еще есть проблемы, пробуем агрессивное исправление
        return self._aggressive_fix_decoding(filename)

    def _try_multiple_decodings(self, filename_bytes):
        """Пробует различные методы декодирования"""
        strategies = [
            # Стандартные кодировки
            lambda: self._try_encoding_list(filename_bytes, ['cp866', 'cp1251', 'koi8-r', 'iso-8859-5']),
            # CP437 цепочки
            lambda: self._try_cp437_chain(filename_bytes),
            # UTF-8 с обработкой ошибок
            lambda: filename_bytes.decode('utf-8', errors='replace'),
            # Windows-1252 как запасной вариант
            lambda: self._try_encoding_list(filename_bytes, ['windows-1252', 'latin1']),
        ]

        for strategy in strategies:
            try:
                result = strategy()
                if result and not self._has_garbled_chars(result):
                    return result
            except:
                continue

        return None

    def _try_encoding_list(self, filename_bytes, encodings):
        """Пробует список кодировок"""
        for encoding in encodings:
            try:
                decoded = filename_bytes.decode(encoding)
                if not self._has_garbled_chars(decoded):
                    return decoded
            except:
                continue
        return None

    def _try_cp437_chain(self, filename_bytes):
        """Пробует цепочки преобразований CP437"""
        chains = [
            # CP437 -> CP1251
            lambda: filename_bytes.decode('cp437').encode('latin1').decode('cp1251'),
            # CP437 -> CP866
            lambda: filename_bytes.decode('cp437').encode('latin1').decode('cp866'),
            # CP850 -> CP1251 (альтернативная DOS кодировка)
            lambda: filename_bytes.decode('cp850').encode('latin1').decode('cp1251'),
        ]

        for chain in chains:
            try:
                result = chain()
                if not self._has_garbled_chars(result):
                    return result
            except:
                continue
        return None

    def _aggressive_fix_decoding(self, filename_bytes):
        """Агрессивное исправление проблемных случаев"""
        # Сначала пробуем стандартное декодирование
        try:
            decoded = filename_bytes.decode('utf-8', errors='replace')
        except:
            decoded = str(filename_bytes)

        # КОМПЛЕКСНЫЙ СЛОВАРЬ ЗАМЕН ДЛЯ ВСЕХ ВОЗМОЖНЫХ КРАКОЗЯБР
        replacements = {
            # === РУССКИЕ СТРОЧНЫЕ БУКВЫ ===
            'à': 'а', 'á': 'а', 'â': 'а', 'ã': 'а', 'ä': 'а', 'å': 'а', 'æ': 'ж',
            'ç': 'з', 'è': 'е', 'é': 'е', 'ê': 'е', 'ë': 'е', 'ì': 'и', 'í': 'и',
            'î': 'и', 'ï': 'и', 'ð': 'о', 'ñ': 'н', 'ò': 'о', 'ó': 'о', 'ô': 'о',
            'õ': 'о', 'ö': 'о', '÷': 'ч', 'ø': 'ш', 'ù': 'у', 'ú': 'у', 'û': 'у',
            'ü': 'у', 'ý': 'ы', 'þ': 'ю', 'ÿ': 'я',

            # === РУССКИЕ ЗАГЛАВНЫЕ БУКВЫ ===
            'À': 'А', 'Á': 'А', 'Â': 'А', 'Ã': 'А', 'Ä': 'А', 'Å': 'А', 'Æ': 'Ж',
            'Ç': 'З', 'È': 'Е', 'É': 'Е', 'Ê': 'Е', 'Ë': 'Е', 'Ì': 'И', 'Í': 'И',
            'Î': 'И', 'Ï': 'И', 'Ð': 'О', 'Ñ': 'Н', 'Ò': 'О', 'Ó': 'О', 'Ô': 'О',
            'Õ': 'О', 'Ö': 'О', '×': 'Ч', 'Ø': 'Ш', 'Ù': 'У', 'Ú': 'У', 'Û': 'У',
            'Ü': 'У', 'Ý': 'Ы', 'Þ': 'Ю', 'ß': 'Я',

            # === ГРЕЧЕСКИЕ И МАТЕМАТИЧЕСКИЕ СИМВОЛЫ ===
            'Σ': 'Е', 'Γ': 'Г', 'τ': 'т', 'α': 'а', 'µ': 'м', 'π': 'п', 'δ': 'д',
            '∞': '', '∩': '', '⌐': '',

            # === ПУНКТУАЦИЯ И СПЕЦИАЛЬНЫЕ СИМВОЛЫ ===
            '«': ' ', '»': ' ', '¿': ' ', '¡': ' ', '½': ' ', '¼': ' ', 'º': ' ',
            '~': ' ', '´': ' ', '¨': ' ', '^': ' ', '`': ' ', '°': ' ', '¬': ' ',

            # === ДОПОЛНИТЕЛЬНЫЕ СИМВОЛЫ ===
            '‚': ',', '„': '"', '…': '...', '†': '*', '‡': '**', '•': '*',
            '‹': '<', '›': '>', '™': '', '¢': '', '£': '', '¤': '',
            '¥': '', '¦': '|', '§': '', '©': '', 'ª': '', '¬': '-',
        }

        # Применяем замены
        fixed = decoded
        for wrong, right in replacements.items():
            fixed = fixed.replace(wrong, right)

        # Убираем множественные пробелы и лишние символы
        fixed = re.sub(r'\s+', ' ', fixed).strip()

        return fixed

    def _has_garbled_chars(self, text):
        """Проверяет наличие кракозябр в тексте"""
        # Расширенный список проблемных символов
        garbled_patterns = [
            # Латинские буквы с диакритиками (не в русском тексте)
            'æ', 'â', 'à', 'á', 'ã', 'ä', 'å', 'ç', 'è', 'é', 'ê', 'ë',
            'ì', 'í', 'î', 'ï', 'ð', 'ñ', 'ò', 'ó', 'ô', 'õ', 'ö', '÷',
            'ø', 'ù', 'ú', 'û', 'ü', 'ý', 'þ', 'ÿ',
            'Æ', 'Â', 'À', 'Á', 'Ã', 'Ä', 'Å', 'Ç', 'È', 'É', 'Ê', 'Ë',
            'Ì', 'Í', 'Î', 'Ï', 'Ð', 'Ñ', 'Ò', 'Ó', 'Ô', 'Õ', 'Ö', '×',
            'Ø', 'Ù', 'Ú', 'Û', 'Ü', 'Ý', 'Þ', 'ß',

            # Греческие и математические символы
            'Σ', 'Γ', 'τ', 'α', 'µ', 'π', 'δ', '∞', '∩', '⌐',

            # Специальные символы
            '«', '»', '¿', '¡', '½', '¼', 'º', '~', '´', '¨', '^', '`', '°', '¬',

            # Редкие символы
            '‚', '„', '…', '†', '‡', '•', '‹', '›', '™', '¢', '£', '¤',
            '¥', '¦', '§', '©', 'ª', '¬'
        ]

        # Если текст содержит много таких символов - это кракозябры
        garbled_count = sum(1 for char in text if char in garbled_patterns)
        return garbled_count > len(text) * 0.1  # Более 10% текста - кракозябры

    def _fix_garbled_text(self, text):
        """Быстрое исправление текста (старый метод для обратной совместимости)"""
        return self._aggressive_fix_decoding(text.encode('utf-8') if isinstance(text, str) else text)

    def _format_crc32_display(self, crc_value, item_type):
        """Форматирует CRC32 для отображения"""
        if item_type == 'directory':
            return 'Папка'
        elif crc_value is None:
            return 'Не указан'
        elif crc_value == 0:
            return 'Пустой файл'
        else:
            return f"CRC32: {crc_value:08X}"

    def _format_relative_path(self, full_path: str) -> str:
        """Форматирует путь относительно корневой папки"""
        if not self.root_path:
            return full_path

        try:
            relative_path = str(Path(full_path).relative_to(self.root_path))
            return f"\\{relative_path}"
        except ValueError:
            return full_path

    def _is_valid_archive(self, file_path: Path) -> bool:
        """Проверяет, является ли файл действительным архивом"""
        try:
            suffix = file_path.suffix.lower()

            if suffix == '.zip':
                with open(file_path, 'rb') as f:
                    header = f.read(4)
                    return header == b'PK\x03\x04' or header == b'PK\x05\x06' or header == b'PK\x07\x08'

            elif suffix == '.7z':
                with open(file_path, 'rb') as f:
                    header = f.read(6)
                    return header == b'7z\xBC\xAF\x27\x1C'

            elif suffix == '.rar':
                with open(file_path, 'rb') as f:
                    header = f.read(8)
                    return header[:4] == b'Rar!' or header == b'\x52\x61\x72\x21\x1A\x07\x00'

        except Exception as e:
            self.logger.debug(f"Ошибка проверки архива {file_path}: {e}")

        return False

    def scan_directory(self, directory_path: str) -> List[Dict[str, Any]]:
        self._stop_event.clear()
        self._processed_files = 0
        self.root_path = Path(directory_path)
        results = []

        try:
            self._total_files = self._count_files(self.root_path)
            self.logger.info(f"Всего файлов для сканирования: {self._total_files}")

            results.append({
                'type': 'directory',
                'path': str(self.root_path),
                'name': self.root_path.name,
                'parent': '',
                'depth': 0,
                'archive_depth': 0,
                'size': 0,
                'crc32': 'Папка'
            })

            for item in self.root_path.iterdir():
                if self._stop_event.is_set():
                    break
                self._scan_recursive(item, results, 1, 0)

        except Exception as e:
            self.logger.error(f"Ошибка при сканировании: {e}")

        return results

    def _count_files(self, path: Path) -> int:
        if self._stop_event.is_set():
            return 0
        count = 0
        try:
            if path.is_file():
                return 1
            else:
                for item in path.iterdir():
                    if self._stop_event.is_set():
                        break
                    try:
                        if item.is_dir():
                            count += self._count_files(item)
                        else:
                            count += 1
                    except (PermissionError, OSError):
                        continue
        except (PermissionError, OSError):
            pass
        return count

    def _scan_recursive(self, path: Path, results: List, depth: int, archive_depth: int):
        if self._stop_event.is_set():
            return
        try:
            if path.is_file():
                if path.suffix.lower() in self.supported_formats and archive_depth < self.max_archive_depth:
                    if self._is_valid_archive(path):
                        self._scan_archive(path, results, depth, archive_depth)
                    else:
                        self.logger.warning(f"Файл с расширением архива, но не архив: {path}")
                        crc32 = self._calculate_file_crc32(path)
                        crc_display = self._format_crc32_display(crc32, 'file')

                        results.append({
                            'type': 'file',
                            'path': self._format_relative_path(str(path)),
                            'name': f"[НЕ АРХИВ] {path.name}",
                            'parent': '',
                            'depth': depth,
                            'archive_depth': archive_depth,
                            'size': path.stat().st_size,
                            'crc32': crc_display
                        })
                        self._processed_files += 1
                        self._update_progress()
                else:
                    crc32 = self._calculate_file_crc32(path)
                    crc_display = self._format_crc32_display(crc32, 'file')

                    results.append({
                        'type': 'file',
                        'path': self._format_relative_path(str(path)),
                        'name': path.name,
                        'parent': '',
                        'depth': depth,
                        'archive_depth': archive_depth,
                        'size': path.stat().st_size,
                        'crc32': crc_display
                    })

                    self._processed_files += 1
                    self._update_progress()
            else:
                results.append({
                    'type': 'directory',
                    'path': self._format_relative_path(str(path)),
                    'name': path.name,
                    'parent': '',
                    'depth': depth,
                    'archive_depth': archive_depth,
                    'size': 0,
                    'crc32': 'Папка'
                })

                for item in path.iterdir():
                    if self._stop_event.is_set():
                        break
                    self._scan_recursive(item, results, depth + 1, archive_depth)

        except PermissionError as e:
            self.logger.warning(f"Нет доступа к: {path} - {e}")
            self._processed_files += 1
            self._update_progress()
        except Exception as e:
            self.logger.error(f"Ошибка при сканировании {path}: {e}")
            self._processed_files += 1
            self._update_progress()

    def _scan_archive(self, archive_path: Path, results: List, depth: int, parent_archive_depth: int):
        if self._stop_event.is_set():
            return
        current_archive_depth = parent_archive_depth + 1

        try:
            suffix = archive_path.suffix.lower()
            archive_crc = self._calculate_file_crc32(archive_path)
            crc_display = self._format_crc32_display(archive_crc, 'archive')

            results.append({
                'type': 'archive',
                'path': self._format_relative_path(str(archive_path)),
                'name': f"[АРХИВ] {archive_path.name}",
                'parent': '',
                'depth': depth,
                'archive_depth': parent_archive_depth,
                'size': archive_path.stat().st_size,
                'crc32': crc_display
            })

            if suffix == '.zip':
                self._scan_zip_contents(archive_path, results, depth + 1, current_archive_depth)
            elif suffix == '.7z':
                self._scan_7z_contents(archive_path, results, depth + 1, current_archive_depth)
            elif suffix == '.rar':
                self._scan_rar_contents_safe(archive_path, results, depth + 1, current_archive_depth)

        except Exception as e:
            self.logger.error(f"Ошибка при сканировании архива {archive_path}: {e}")

            results.append({
                'type': 'archive_error',
                'path': self._format_relative_path(str(archive_path)),
                'name': f"[ОШИБКА АРХИВА] {archive_path.name}",
                'parent': '',
                'depth': depth,
                'archive_depth': parent_archive_depth,
                'size': archive_path.stat().st_size,
                'crc32': f'Ошибка: {str(e)}'
            })

        finally:
            self._processed_files += 1
            self._update_progress()

    def _scan_zip_contents(self, archive_path: Path, results: List, depth: int, archive_depth: int):
        try:
            if not self._is_valid_archive(archive_path):
                raise zipfile.BadZipFile(f"Файл не является ZIP архивом: {archive_path}")

            with zipfile.ZipFile(archive_path, 'r') as zip_ref:
                for file_info in zip_ref.infolist():
                    if self._stop_event.is_set():
                        break

                    filename = self._decode_zip_filename(file_info.filename)

                    is_directory = filename.endswith('/') or file_info.file_size == 0

                    inner_suffix = Path(filename).suffix.lower()
                    is_nested_archive = inner_suffix in self.supported_formats

                    crc_display = self._format_crc32_display(file_info.CRC, 'archive_file')

                    if is_directory:
                        item_type = 'archive_directory'
                        name_prefix = ""
                        if filename.endswith('/'):
                            filename = filename[:-1]
                    elif is_nested_archive:
                        item_type = 'nested_archive'
                        name_prefix = "[ВЛ.АРХ] "
                    else:
                        item_type = 'archive_file'
                        name_prefix = ""

                    results.append({
                        'type': item_type,
                        'path': f"{self._format_relative_path(str(archive_path))}/{filename}",
                        'name': f"{name_prefix}{filename}",
                        'parent': '',
                        'depth': depth,
                        'archive_depth': archive_depth,
                        'size': file_info.file_size,
                        'crc32': crc_display
                    })

        except zipfile.BadZipFile as e:
            self.logger.error(f"Поврежденный ZIP архив {archive_path}: {e}")
            raise
        except Exception as e:
            self.logger.error(f"Ошибка при чтении ZIP архива {archive_path}: {e}")
            raise

    def _scan_7z_contents(self, archive_path: Path, results: List, depth: int, archive_depth: int):
        """Сканирует содержимое 7z архивов с правильным API"""
        try:
            if not self._is_valid_archive(archive_path):
                raise py7zr.Bad7zFile(f"Файл не является 7z архивом: {archive_path}")

            with py7zr.SevenZipFile(archive_path, 'r') as archive:
                archive_info = archive.list()

                for file_info in archive_info:
                    if self._stop_event.is_set():
                        break

                    filename = self._decode_filename(file_info.filename)
                    inner_suffix = Path(filename).suffix.lower()
                    is_nested_archive = inner_suffix in self.supported_formats

                    is_directory = file_info.is_directory if hasattr(file_info, 'is_directory') else False

                    crc_value = None
                    if hasattr(file_info, 'crc32') and file_info.crc32:
                        crc_value = file_info.crc32
                    elif hasattr(file_info, 'crc') and file_info.crc:
                        crc_value = file_info.crc

                    crc_display = self._format_crc32_display(crc_value, 'archive_file')

                    if is_directory:
                        item_type = 'archive_directory'
                        name_prefix = ""
                    elif is_nested_archive:
                        item_type = 'nested_archive'
                        name_prefix = "[ВЛ.АРХ] "
                    else:
                        item_type = 'archive_file'
                        name_prefix = ""

                    file_size = 0
                    if hasattr(file_info, 'uncompressed'):
                        file_size = file_info.uncompressed
                    elif hasattr(file_info, 'size'):
                        file_size = file_info.size

                    results.append({
                        'type': item_type,
                        'path': f"{self._format_relative_path(str(archive_path))}/{filename}",
                        'name': f"{name_prefix}{filename}",
                        'parent': '',
                        'depth': depth,
                        'archive_depth': archive_depth,
                        'size': file_size,
                        'crc32': crc_display
                    })

        except py7zr.Bad7zFile as e:
            self.logger.error(f"Поврежденный 7z архив {archive_path}: {e}")
            raise
        except Exception as e:
            self.logger.error(f"Ошибка при чтении 7z архива {archive_path}: {e}")
            raise

    def _scan_rar_contents_safe(self, archive_path: Path, results: List, depth: int, archive_depth: int):
        """Безопасное сканирование RAR архивов"""
        try:
            if not self._is_valid_archive(archive_path):
                raise rarfile.BadRarFile(f"Файл не является RAR архивом: {archive_path}")

            with rarfile.RarFile(archive_path, 'r') as archive:
                for file_info in archive.infolist():
                    if self._stop_event.is_set():
                        break

                    filename = self._decode_filename(file_info.filename)

                    inner_suffix = Path(filename).suffix.lower()
                    is_nested_archive = inner_suffix in self.supported_formats

                    crc_display = self._format_crc32_display(file_info.CRC, 'archive_file')

                    archive_type = 'nested_archive' if is_nested_archive else 'archive_file'
                    name_prefix = "[ВЛ.АРХ] " if is_nested_archive else ""

                    if is_nested_archive and not self._7zip_path:
                        name_prefix = "[ВЛ.АРХ - НЕТ 7Z] "

                    results.append({
                        'type': archive_type,
                        'path': f"{self._format_relative_path(str(archive_path))}/{filename}",
                        'name': f"{name_prefix}{filename}",
                        'parent': '',
                        'depth': depth,
                        'archive_depth': archive_depth,
                        'size': file_info.file_size,
                        'crc32': crc_display
                    })

        except rarfile.NeedFirstVolume:
            self.logger.error(f"Многотомный RAR архив: {archive_path}")
            raise
        except rarfile.BadRarFile as e:
            self.logger.error(f"Поврежденный RAR архив {archive_path}: {e}")
            raise
        except Exception as e:
            self.logger.error(f"Ошибка при чтении RAR архива {archive_path}: {e}")
            raise