# Notification Generator (исходники)

Это простое приложение на Python + PyQt5 для генерации уведомлений по шаблону.
Особенности:
- Вводите шаблон (текст) с плейсхолдерами {{flat}}, {{date}}, {{time}}.
- Вставляете список номеров квартир (через пробел, запятую или построчно).
- Выбираете диапазон дат и времени.
- Программа генерирует `.docx` для каждой квартиры и упаковывает в `.zip`.
- Сохраняет настройки и историю в `~\NotificationGenerator\`.

## Быстрый запуск (на Windows)
1. Установите Python 3.8+.
2. Создайте виртуальное окружение:
```powershell
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
python main.py
```

## Сборка в .exe (локально)
Установите pyinstaller:
```
pip install pyinstaller
pyinstaller --onefile --windowed main.py
```
Готовый `main.exe` будет в папке `dist`.

## Сборка через GitHub Actions
В репозитории есть workflow `.github/workflows/build-windows.yml`, который собирает `.exe` на Windows runner.
