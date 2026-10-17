"""relink — массовая замена путей к внешним файлам в документах КОМПАС-3D.

Находит все сборки/детали/чертежи/спецификации/фрагменты в папке (рекурсивно),
открывает каждый невидимо, читает его ссылки на внешние файлы и меняет начало
пути по таблице замен (например  \\\\metall\\Проекты  ->  G:\\Проекты).

    python relink.py G:\\Проекты --map "//metall/Проекты=G:/Проекты"            # пробный прогон
    python relink.py G:\\Проекты --map "//metall/Проекты=G:/Проекты" --apply    # записать

Слэши / и \\ равнозначны, регистр не важен. Можно несколько --map.
Перед записью оригинал копируется в папку резервных копий (--backup, по умолчанию
<папка>\\_relink_backup_<дата>). Отчёт пишется в relink_report_<дата>.csv.

Работает с КОМПАС-3D v20 и новее (проверено на v24); установленный КОМПАС находится сам.
Версия файла сохраняется: файл v23, обработанный в v24, останется v23
(--save-version current — сохранять в текущей версии, --save-version 23 — всегда в v23).
"""
import argparse
import csv
import datetime
import glob
import os
import re
import shutil
import sys
import winreg

import pythoncom
from win32com.client import Dispatch, gencache

EXTS = (".a3d", ".m3d", ".cdw", ".spw", ".frw", ".kdw")
# ksSaveDocumentVersionEnum: основная версия КОМПАСа -> код для SaveAsEx
SDV = {17: 21, 18: 23, 19: 25, 20: 26, 21: 27, 22: 28, 23: 29}


def find_bin():
    """Папка Bin того КОМПАСа, который зарегистрирован как Kompas.Application.7."""
    try:
        clsid = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, r"Kompas.Application.7\CLSID")
        exe = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, rf"CLSID\{clsid}\LocalServer32")
        d = os.path.dirname(exe.strip().strip('"').split('"')[0])
        if os.path.exists(os.path.join(d, "kAPI7.tlb")):
            return d
    except OSError:
        pass
    found = sorted(glob.glob(r"C:\Program Files\ASCON\KOMPAS-3D v*\Bin\kAPI7.tlb"))
    if not found:
        sys.exit("Не найден установленный КОМПАС-3D (kAPI7.tlb).")
    return os.path.dirname(found[-1])


BIN = find_bin()


def _tlb(name):
    a = pythoncom.LoadTypeLib(os.path.join(BIN, name)).GetLibAttr()
    return gencache.EnsureModule(str(a[0]), a[1], a[3], a[4])


def file_major(d1):
    """Основная версия КОМПАСа, в которой сохранён открытый файл (None — неизвестно)."""
    v = d1.OpenVersion >> 24
    return v if v >= 17 else None


def norm(p):
    """Привести путь к виду с обратными слэшами: '//metall/x' -> '\\\\metall\\x', 'G/x' -> 'G:\\x'."""
    p = p.strip().strip('"').replace("/", "\\")
    if len(p) >= 2 and p[0].isalpha() and p[1] == "\\":      # 'G\...' -> 'G:\...'
        p = p[0] + ":" + p[1:]
    if len(p) == 1 and p.isalpha():
        p += ":"
    return p.rstrip("\\")


def parse_map(items):
    pairs = []
    for it in items:
        if "=" not in it:
            sys.exit(f"--map должен быть вида СТАРЫЙ=НОВЫЙ, получено: {it}")
        old, new = it.split("=", 1)
        pairs.append((norm(old), norm(new)))
    # длинные префиксы первыми, чтобы более точное правило срабатывало раньше
    return sorted(pairs, key=lambda x: -len(x[0]))


def remap(path, pairs):
    p = path.replace("/", "\\")
    low = p.lower()
    for old, new in pairs:
        o = old.lower()
        if low == o or low.startswith(o + "\\"):
            return new + p[len(old):]
    return None


def find_files(root, skip_dirs):
    for d, dirs, files in os.walk(root):
        dirs[:] = [x for x in dirs if os.path.join(d, x).lower() not in skip_dirs
                   and not x.lower().startswith("_relink_backup")]
        for f in files:
            if f.lower().endswith(EXTS):
                yield os.path.join(d, f)


def main():
    ap = argparse.ArgumentParser(description="Массовая замена путей ссылок в файлах КОМПАС-3D")
    ap.add_argument("roots", nargs="+", help="папки (или отдельные файлы) для обработки")
    ap.add_argument("--map", action="append", required=True, metavar="СТАРЫЙ=НОВЫЙ",
                    help="замена начала пути; можно указать несколько раз")
    ap.add_argument("--apply", action="store_true", help="реально сохранить изменения")
    ap.add_argument("--backup", help="папка резервных копий")
    ap.add_argument("--no-backup", action="store_true", help="не делать резервных копий")
    ap.add_argument("--save-version", default="keep", metavar="keep|current|23|22…",
                    help="в какой версии сохранять: keep — как был файл (по умолчанию), "
                         "current — в версии запущенного КОМПАСа, число — в указанной версии")
    ap.add_argument("--ext", help="расширения через запятую (по умолчанию все: a3d,m3d,cdw,spw,frw,kdw)")
    a = ap.parse_args()

    global EXTS
    if a.ext:
        EXTS = tuple("." + e.strip(". ").lower() for e in a.ext.split(","))
    pairs = parse_map(a.map)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    files = []
    for r in a.roots:
        r = os.path.abspath(r)
        files += [r] if os.path.isfile(r) else list(find_files(r, set()))
    if not files:
        sys.exit("Файлы КОМПАС не найдены.")
    first_root = os.path.abspath(a.roots[0])
    base_dir = first_root if os.path.isdir(first_root) else os.path.dirname(first_root)
    backup = a.backup or os.path.join(base_dir, f"_relink_backup_{stamp}")

    print("Замены:")
    for o, n in pairs:
        print(f"  {o}  ->  {n}")
    print(f"Файлов найдено: {len(files)}   режим: {'ЗАПИСЬ' if a.apply else 'ПРОБНЫЙ (ничего не сохраняется)'}\n")

    K7 = _tlb("kAPI7.tlb")
    C = _tlb("ksConstants.tlb").constants
    app = K7.IApplication(Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(
        K7.IApplication.CLSID, pythoncom.IID_IDispatch))
    m = re.search(r"v(\d+)", app.ApplicationName(True) or "") or re.search(r"v(\d+)", BIN)
    app_major = int(m.group(1)) if m else None
    sv = a.save_version.lower()
    if sv not in ("keep", "current") and not (sv.isdigit() and int(sv) in SDV):
        sys.exit(f"--save-version: keep, current или одна из версий {sorted(SDV)}")
    print(f"КОМПАС: {app.ApplicationName(True)}  ({BIN})   сохранение: {sv}\n")
    old_hide = app.HideMessage
    app.HideMessage = C.ksHideMessageNo      # на вопросы «найти файл?» и т.п. отвечать «Нет»

    report = os.path.join(base_dir, f"relink_report_{stamp}.csv")
    rows = []
    n_changed = n_err = 0
    try:
        for i, path in enumerate(files, 1):
            tag = f"[{i}/{len(files)}] {path}"
            doc = None
            try:
                doc = app.Documents.Open(path, False, not a.apply)   # невидимо; в пробном — только чтение
                if doc is None:
                    raise RuntimeError("КОМПАС не открыл файл")
                d1 = K7.IKompasDocument1(doc._oleobj_.QueryInterface(
                    K7.IKompasDocument1.CLSID, pythoncom.IID_IDispatch))
                seen = set()     # в чертеже одна модель встречается по разу на каждый вид
                links = [l for l in (d1.ExternalFilesNames(False) or ())
                         if not (l.lower() in seen or seen.add(l.lower()))]
                todo = [(l, remap(l, pairs)) for l in links]
                todo = [(o, n) for o, n in todo if n and n != o]
                if not todo:
                    doc.Close(C.kdDoNotSaveChanges)
                    continue
                print(tag)
                ok_all = True
                for o, n in todo:
                    exists = os.path.exists(n)
                    status = "будет заменено"
                    if a.apply:
                        ok = d1.ReplaceExternalFilesNames(False, o, n)
                        ok_all &= bool(ok)
                        status = "заменено" if ok else "ОШИБКА замены"
                    print(f"    {o}\n      -> {n}   [{status}{'' if exists else ', НОВОГО ФАЙЛА НЕТ'}]")
                    rows.append([path, o, n, "да" if exists else "НЕТ", status])
                if a.apply:
                    if not a.no_backup:
                        dst = os.path.join(backup, os.path.relpath(path, base_dir)
                                           if path.lower().startswith(base_dir.lower())
                                           else os.path.basename(path))
                        os.makedirs(os.path.dirname(dst), exist_ok=True)
                        shutil.copy2(path, dst)
                    fv = file_major(d1)
                    target = (fv if sv == "keep" else None if sv == "current" else int(sv))
                    if target is not None and app_major is not None and target >= app_major:
                        target = None                # своя версия — обычное сохранение
                    if target is not None and target not in SDV:
                        raise RuntimeError(f"файл версии v{target}: сохранение в неё не поддерживается, "
                                           "укажите --save-version current или номер версии")
                    mtime = os.path.getmtime(path)
                    if target is None:
                        doc.Save()                   # в v24 возвращает void — проверяем по дате файла
                    elif not d1.SaveAsEx(path, SDV[target]):
                        raise RuntimeError(f"не удалось сохранить в формате v{target}")
                    if os.path.getmtime(path) == mtime:
                        raise RuntimeError("не удалось сохранить (файл занят или только для чтения?)")
                    ver = f"v{target}" if target else f"v{app_major or '?'}"
                    print(f"    сохранено ({ver})" if ok_all else f"    сохранено ({ver}), но не все ссылки заменились")
                doc.Close(C.kdDoNotSaveChanges)
                doc = None
                n_changed += 1
            except Exception as e:
                n_err += 1
                print(f"{tag}\n    ОШИБКА: {e}")
                rows.append([path, "", "", "", f"ОШИБКА: {e}"])
                if doc is not None:
                    try:
                        doc.Close(C.kdDoNotSaveChanges)
                    except Exception:
                        pass
    finally:
        app.HideMessage = old_hide

    with open(report, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Файл", "Старая ссылка", "Новая ссылка", "Новый файл существует", "Статус"])
        w.writerows(rows)

    print(f"\nФайлов со ссылками на замену: {n_changed}, ошибок: {n_err}")
    print(f"Отчёт: {report}")
    if a.apply and not a.no_backup and n_changed:
        print(f"Резервные копии: {backup}")
    if not a.apply:
        print("Это был пробный прогон. Для записи добавьте --apply")


if __name__ == "__main__":
    main()
