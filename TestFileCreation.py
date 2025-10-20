# test_crc.py
import os
from pathlib import Path
import scanner


def test_crc_functionality():
    """Тестирование функциональности CRC32"""

    # Создаем тестовый файл
    test_content = b"Hello, World! This is a test file for CRC32 calculation."
    test_file = Path("test_crc_file.txt")

    with open(test_file, "wb") as f:
        f.write(test_content)

    print("🧪 Тестирование CRC32 функциональности")
    print("=" * 50)

    # Тестируем сканер
    test_scanner = scanner.ArchiveScanner()

    # Тестируем вычисление CRC32 для обычного файла
    crc_result = test_scanner._calculate_file_crc32(test_file)
    print(f"📄 Обычный файл: {test_file}")
    print(f"🔢 CRC32: {crc_result:08X}" if crc_result else "❌ Не удалось вычислить")

    # Тестируем сканирование директории
    print(f"\n📁 Сканирование текущей директории...")
    results = test_scanner.scan_directory(".")

    # Показываем первые 5 результатов с CRC32
    print(f"\n📋 Первые 5 результатов:")
    for i, item in enumerate(results[:5]):
        print(f"{i + 1}. {item['type']}: {item['name']}")
        print(f"   📏 Размер: {item.get('size', 0)} байт")
        print(f"   🔢 CRC32: {item.get('crc32', 'N/A')}")
        print()

    # Удаляем тестовый файл
    test_file.unlink()
    print("✅ Тестовый файл удален")


if __name__ == "__main__":
    test_crc_functionality()