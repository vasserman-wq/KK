r"""Решения разметки «включить в карту / не включать».

После первого запуска режима 1 хранятся в рабочей книге карты (скрытый лист KK_Решения),
то есть в одном файле рядом с чертежом. До появления карты — временно в %LOCALAPPDATA%\KK\pending.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

from .paths import user_data_dir
from .xlshim import ShimError

SHEET = 'KK_Решения'
HEADER = ['Лист', 'Вид', 'Имя', 'Тип', 'Текст', 'X', 'Y', 'Включить']
SCHEMA = 'KK_DECISIONS_1'


def _pending_path(drawing):
    digest = hashlib.sha1(str(drawing).strip().lower().encode('utf-8')).hexdigest()[:16]
    folder = user_data_dir() / 'pending'
    folder.mkdir(parents=True, exist_ok=True)
    return folder / f'{digest}.json'


def row_key(row):
    return (row['sheet'], row['name'], round(float(row['x']), 5), round(float(row['y']), 5))


def load(card, drawing):
    """Список словарей sheet/view/name/kind/text/x/y/include."""
    rows = []
    if card is not None and card.book.has_sheet(SHEET):
        ws = card.book.sheet(SHEET)
        if str(ws.range('A1').get_value2() or '') == SCHEMA:
            for r in range(3, ws.max_row() + 1):
                v = ws.range(f'A{r}:H{r}').get_value2()[0]
                if v[0] in (None, ''):
                    continue
                rows.append({'sheet': str(v[0]), 'view': str(v[1] or ''), 'name': str(v[2]), 'kind': str(v[3] or ''),
                             'text': str(v[4] or ''), 'x': float(v[5]), 'y': float(v[6]),
                             'include': str(v[7]).strip() in ('1', '1.0', 'True')})
    pending = _pending_path(drawing)
    if pending.exists():
        known = {row_key(r) for r in rows}
        for r in json.loads(pending.read_text(encoding='utf-8')).get('rows', []):
            if row_key(r) not in known:
                rows.append(r)
    return rows


def save(card, drawing, rows):
    """В карту, если она есть (сохраняет вызывающий), иначе во временный файл."""
    rows = list(rows)
    if card is None:
        path = _pending_path(drawing)
        tmp = path.with_suffix('.tmp')
        tmp.write_text(json.dumps({'drawing': str(drawing), 'rows': rows}, ensure_ascii=False), encoding='utf-8')
        tmp.replace(path)
        return 'pending'
    book = card.book
    ws = book.sheet(SHEET) if book.has_sheet(SHEET) else book.add_sheet(after=book.sheets()[-1])
    if ws.name != SHEET:
        ws.set_name(SHEET)
    ws.set_visible(2)
    ws.used_range().clear()
    ws.range('A1').set_value2(SCHEMA)
    ws.range('B1').set_value2(str(drawing))
    ws.range('A2:H2').set_value2([HEADER])
    if rows:
        ws.range(f'A3:H{len(rows) + 2}').set_value2(
            [[r['sheet'], r.get('view', ''), r['name'], r.get('kind', ''), r.get('text', ''), float(r['x']), float(r['y']),
              1 if r['include'] else 0] for r in rows])
    pending = _pending_path(drawing)
    if pending.exists():
        pending.unlink()
    return 'card'


def to_factory_rows(rows):
    return [(r['sheet'], r['name'], r['x'], r['y'], r['include']) for r in rows]
