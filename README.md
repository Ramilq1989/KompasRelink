# KompasRelink

Массовая замена путей к внешним файлам (ссылкам) в документах КОМПАС-3D:
сборках, деталях, чертежах, спецификациях и фрагментах — во всех подпапках.

Пример: `\metall\Проекты` → `G:\Проекты`.

## Скачать

Готовые `KompasRelink.exe` и `KompasRelinkLib.dll` — в разделе [Releases](../../releases).

## Состав

| Файл | Назначение |
|---|---|
| `src/KompasRelink.cs`, `src/KompasLibrary.cs` | Исходники программы с интерфейсом (C#, .NET Framework 4) |
| `build.bat` | Сборка `KompasRelink.exe` и `KompasRelinkLib.dll` |
| `relink.py` | Консольный вариант на Python (`pywin32`) |
| `relink.bat` | Интерактивный запуск `relink.py` |
| `Инструкция.html` | Подробная инструкция |

## Сборка

```bat
build.bat
```

Нужен только установленный .NET Framework 4 (есть в Windows).

## Python-вариант

```bat
pip install pywin32
python relink.py G:\Проекты --map "//metall/Проекты=G:/Проекты"           & rem пробный прогон
python relink.py G:\Проекты --map "//metall/Проекты=G:/Проекты" --apply   & rem записать
```

Перед записью делаются резервные копии. Работает с КОМПАС-3D v20 и новее (проверено на v24).
