import subprocess
import shutil
from pathlib import Path


def check_7zip_availability():
    """Проверяет доступность 7-Zip"""

    print("🔍 Проверка доступности 7-Zip...")

    # Варианты путей к 7z.exe
    possible_paths = [
        "7z",
        "7z.exe",
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
        r"D:\Program Files\7-Zip\7z.exe",
    ]

    for path in possible_paths:
        try:
            result = subprocess.run([path, "--help"], capture_output=True, text=True, timeout=5)
            if result.returncode == 0 or "7-Zip" in result.stdout:
                print(f"✅ 7-Zip найден: {path}")
                return path
        except:
            continue

    print("❌ 7-Zip не найден. Установите 7-Zip для полной поддержки RAR архивов")
    print("📥 Скачать: https://www.7-zip.org/download.html")
    return None


def check_rar_support():
    """Проверяет поддержку RAR архивов"""
    print("\n🔍 Проверка поддержки RAR...")

    try:
        import rarfile
        rarfile.tool_setup()  # Проверяем настройки RAR
        print("✅ RAR поддержка: доступна (только чтение)")
        return True
    except Exception as e:
        print(f"❌ RAR поддержка: ограничена - {e}")
        return False


if __name__ == "__main__":
    zip_path = check_7zip_availability()
    rar_support = check_rar_support()

    if not zip_path and not rar_support:
        print("\n💡 Рекомендация: Установите 7-Zip для полной поддержки архивов")