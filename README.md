# AlbumBuilder

## Назначение
Автоматизация сборки фотоальбомов в CorelDRAW 2025.

## Структура
- `source/` — эталонные исходники (`.bas`)
- `forms/`  — формы (`.frm`)
- `build/`  — временная сборка для импорта в Corel
- `AlbumBuilder.gsm` — переносимый контейнер

## Рабочий процесс
 
1. Редактирование кода: **VS Code**
2. Сборка:
   ```powershell
   .\tools\sync.ps1 AlbumBuilder

    Импорт файлов из build/ в Corel VBA

    Отладка — в Corel

    Экспорт изменений обратно в source/ и forms/

    Commit в Git

Зависимости

    CorelDRAW Graphics Suite 2025

    PowerShell 5+

Важные правила

    .gsm не является источником кода

    В Git хранятся только .bas, .frm, .cls

    Папка build/ не коммитится