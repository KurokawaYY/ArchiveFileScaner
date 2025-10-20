import PyInstaller.__main__
import os
import shutil
import sys
from pathlib import Path
from datetime import datetime


def build_exe():
    """Сборка проекта в EXE файл с версией и иконкой"""

    # Конфигурация сборки
    APP_VERSION = "0.6"  # ⬅️ ЗДЕСЬ МЕНЯЕМ ВЕРСИЮ
    APP_NAME = f"ArchScanner_v{APP_VERSION}"
    ICON_PATH = "Res/app_icon.ico"  # ⬅️ ПУТЬ К ИКОНКЕ

    print("🚀 Начинаем сборку EXE файла...")
    print("=" * 50)
    print(f"📋 Версия: {APP_VERSION}")
    print(f"📁 Имя приложения: {APP_NAME}")

    # Создаем папку Releases если её нет
    releases_dir = Path("Releases")
    releases_dir.mkdir(exist_ok=True)
    print(f"✅ Папка Releases: {releases_dir.absolute()}")

    # Очистка предыдущих сборок
    if Path("build").exists():
        shutil.rmtree("build")
        print("✅ Очищена папка build")

    if Path("dist").exists():
        shutil.rmtree("dist")
        print("✅ Очищена папка dist")

    # Параметры сборки
    build_params = [
        'main.py',
        '--onefile',
        '--windowed',
        f'--name={APP_NAME}',
        '--clean',
        '--noconfirm',
        '--add-data=scanner.py;.',
        '--add-data=excel_exporter.py;.',
        '--hidden-import=py7zr',
        '--hidden-import=rarfile',
        '--hidden-import=openpyxl',
        '--hidden-import=openpyxl.styles',
        '--hidden-import=openpyxl.workbook',
    ]

    # Добавляем иконку если она существует
    if Path(ICON_PATH).exists():
        build_params.append(f'--icon={ICON_PATH}')
        print(f"✅ Используется иконка: {ICON_PATH}")
    else:
        print("⚠️ Иконка не найдена, будет использована стандартная")

    print("\n📦 Параметры сборки:")
    for param in build_params:
        print(f"   {param}")

    print("\n⏳ Сборка может занять несколько минут...")
    print("=" * 50)

    try:
        # Запускаем сборку
        PyInstaller.__main__.run(build_params)

        # Перемещаем и переименовываем EXE файл
        source_exe = Path("dist") / f"{APP_NAME}.exe"
        if source_exe.exists():
            # Формируем имя файла с датой
            current_date = datetime.now().strftime("%Y%m%d")
            final_exe_name = f"{APP_NAME}_{current_date}.exe"
            final_exe_path = releases_dir / final_exe_name

            # Копируем в Releases
            shutil.copy2(source_exe, final_exe_path)

            print("🎉 СБОРКА УСПЕШНО ЗАВЕРШЕНА!")
            print("=" * 50)
            print(f"📁 Исходный EXE: {source_exe.absolute()}")
            print(f"📁 Финальный EXE: {final_exe_path.absolute()}")
            print(f"📏 Размер файла: {final_exe_path.stat().st_size / (1024 * 1024):.2f} MB")

            # Создаем файл с информацией о сборке
            info_file = releases_dir / f"build_info_{APP_VERSION}_{current_date}.txt"
            with open(info_file, 'w', encoding='utf-8') as f:
                f.write(f"Archive Structure Scanner\n")
                f.write(f"Version: {APP_VERSION}\n")
                f.write(f"Build date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"File: {final_exe_name}\n")
                f.write(f"Size: {final_exe_path.stat().st_size} bytes\n")

            print(f"📄 Создан файл информации: {info_file.name}")

            # Предлагаем открыть папку
            open_folder = input("\n📂 Открыть папку Releases? (y/n): ")
            if open_folder.lower() == 'y':
                os.startfile(str(releases_dir))

        else:
            print("❌ Ошибка: EXE файл не найден после сборки!")

    except Exception as e:
        print(f"❌ Ошибка при сборке: {e}")
        input("Нажмите Enter для выхода...")


def create_bat_file():
    """Создает BAT файл для простой сборки"""
    bat_content = """@echo off
echo Запуск сборки Archive Structure Scanner...
python build.py
pause
"""

    with open("build.bat", "w", encoding="utf-8") as f:
        f.write(bat_content)
    print("✅ Создан build.bat для простой сборки")


if __name__ == "__main__":
    create_bat_file()
    build_exe()