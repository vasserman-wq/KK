"""Модель рабочей карты контроля (лист «Карта контроля» книги v19).

Правила перенесены из модуля OK_IntroEditor шаблона (добавление и удаление вводных,
нумерация ключей, сводка I24:I27, состояние PENDING). Нумерацию параметров чертежа
выполняет макрос 1 (режим 1); программа номера баллонов не меняет.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

from .xlshim import Factory, ShimError

CARD_SHEET = 'Карта контроля'
SERVICE_SHEET = 'ОК_Служебное'
RAW_SHEET = 'Данные из чертежа'
TOOLS_SHEET = 'СписокТО'
FIRST_ROW = 7
BLOCK_ROWS = 4
END_KEY = 'ОК'
STAMP_TOOL = 'Клеймо ОТК'
_NUM_PREFIX = re.compile(r'^\s*[0-9]+[.)]\s*(?=[^0-9]|$)')

# поле -> (имя в книге, адрес по умолчанию)
HEADER_FIELDS = {
    'kd_number': ('OK_KDNumber', 'L4', 'Номер КД'),
    'product_name': ('OK_ProductName', 'L5', 'Наименование изделия'),
    'primary_application': ('OK_PrimaryApplication', 'L6', 'Первичное применение'),
    'mass': ('OK_Mass', 'I17', 'Масса детали, кг'),
    'equipment': ('OK_Equipment', 'I18', 'Оборудование ОК1'),
    'variant_enabled': ('OK_VariantEnabled', 'I11', 'Таблица исполнений'),
    'variant_text': ('OK_VariantText', 'I12', 'Текст исполнений'),
}
ROLES_ADDR, NAMES_ADDR, FIXED_ADDR = ('OK_HeaderRoles', 'H4:H8'), ('OK_HeaderNames', 'I4:I8'), ('OK_FixedRows', 'I20:I22')


def default_card_path(drawing):
    """Как KC14_DefaultPath в макросе 1: «<имя чертежа> — карта контроля.xlsm» в папке чертежа."""
    p = Path(drawing)
    return p.with_name(f'{p.stem} — карта контроля.xlsm')


def numbered_text(text, number):
    """Как OK_NumberedText: убрать старый номер «N.» в начале и поставить новый."""
    text = str(text or '').strip()
    text = _NUM_PREFIX.sub('', text, count=1).strip()
    return f'{number}. {text}' if text else ''


def strip_number(text):
    return _NUM_PREFIX.sub('', str(text or '').strip(), count=1).strip()


def to_minutes(value):
    if value in (None, ''):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(',', '.').strip())
    except ValueError:
        return None


def percent_text(value):
    """Как KC15_PercentText: 1 -> 100%, '50' -> 50%, текст без числа остаётся как есть."""
    text = str(value if value is not None else '').replace('\xa0', ' ').strip()
    if not text:
        return ''
    explicit = '%' in text
    raw = text.replace('%', '').replace(' ', '').replace(',', '.')
    try:
        n = float(raw)
    except ValueError:
        return text
    if not explicit and 0 <= n <= 1:
        n *= 100
    return (f'{n:g}') + '%'


@dataclass
class Block:
    row: int                 # первая строка на листе
    rows: int                # число строк блока
    key: str
    text: str
    tools: list = field(default_factory=list)
    volume: str = ''
    minutes: float | None = None
    kind: str = 'param'      # intro | param | stamp
    extra_text: list = field(default_factory=list)

    @property
    def number(self):
        try:
            return int(str(self.key).strip())
        except ValueError:
            return None

    # права на правку (docs/ARCHITECTURE.md, раздел 3)
    def can_edit(self, column):
        if column == 'key':
            return False
        return column in ('text', 'tools', 'volume', 'minutes')

    @property
    def can_delete(self):
        return self.kind == 'intro'


class Card:
    """Рабочая книга карты. Открывается тем же эмулятором, что и у макросов."""

    def __init__(self, path, factory=None):
        self.path = Path(path)
        self.factory = factory or Factory()
        self.app = self.factory.new_app()
        self.book = self.app.open(str(self.path))
        if not self.book.has_sheet(CARD_SHEET) or not self.book.has_sheet(SERVICE_SHEET):
            raise ShimError('Файл не похож на рабочую карту контроля (нет листов «Карта контроля» и «ОК_Служебное»).')
        self.ws = self.book.sheet(CARD_SHEET)

    def close(self):
        self.app.quit()

    @property
    def dirty(self):
        return not self.book.saved

    def save(self):
        self.book.save()

    # --- служебные данные
    def service(self, key):
        ws = self.book.sheet(SERVICE_SHEET)
        for r in range(1, 41):
            if str(ws.cells(r, 1).get_value2() or '').strip().lower() == key.lower():
                return ws.cells(r, 2).get_value2()
        return None

    @property
    def state(self):
        return str(self.book.sheet(SERVICE_SHEET).range('B10').get_value2() or '')

    @property
    def drawing_path(self):
        return str(self.book.sheet(SERVICE_SHEET).range('B2').get_value2() or '')

    def set_pending(self):
        self.book.sheet(SERVICE_SHEET).range('B10').set_value2('PENDING')

    # --- поля по именам
    def _field(self, name, address):
        try:
            rng = self.book.names.item(name).refers_to_range()
            if rng.sheet is self.ws:
                return rng
        except ShimError:
            pass
        return self.ws.range(address)

    def header(self):
        data = {key: self._field(name, addr).cells(1, 1).get_value2() for key, (name, addr, _) in HEADER_FIELDS.items()}
        data = {k: '' if v is None else (f'{v:g}' if isinstance(v, float) else str(v)) for k, v in data.items()}
        roles = self._field(*ROLES_ADDR).get_value2()
        names = self._field(*NAMES_ADDR).get_value2()
        data['signatures'] = [(str(r[0] or ''), str(n[0] or '')) for r, n in zip(roles, names)]
        data['fixed_rows'] = [str(x[0] or '') for x in self._field(*FIXED_ADDR).get_value2()]
        return data

    def set_header(self, **values):
        for key, value in values.items():
            if key in HEADER_FIELDS:
                name, addr, _ = HEADER_FIELDS[key]
                if key == 'mass' and str(value).strip():
                    num = to_minutes(value)
                    if num is None:
                        raise ShimError('Масса должна быть числом.')
                    value = num
                self._field(name, addr).cells(1, 1).set_value2(value if value != '' else None)
            elif key == 'signatures':
                roles, names = self._field(*ROLES_ADDR), self._field(*NAMES_ADDR)
                for i, (role, person) in enumerate(list(value)[:5], 1):
                    roles.cells(i, 1).set_value2(role or None)
                    names.cells(i, 1).set_value2(person or None)
            elif key == 'fixed_rows':
                area = self._field(*FIXED_ADDR)
                for i, text in enumerate(list(value)[:3], 1):
                    area.cells(i, 1).set_value2(text or None)

    def tools_list(self):
        try:
            values = self.book.names.item('ToolsList').refers_to_range().get_value2()
        except ShimError:
            if not self.book.has_sheet(TOOLS_SHEET):
                return []
            ws = self.book.sheet(TOOLS_SHEET)
            values = ws.range(f'A1:A{max(ws.max_row(), 1)}').get_value2()
        values = values if isinstance(values, list) else [[values]]
        return [str(v[0]).strip() for v in values if v and v[0] not in (None, '')]

    # --- блоки
    def intro_last(self):
        try:
            nm = self.book.names.item('OK_IntroArea')
        except ShimError:
            return FIRST_ROW - 1
        if nm.refers_to.strip() == '=0':
            return FIRST_ROW - 1
        area = nm.refers_to_range()
        if area.sheet is not self.ws or area.r1 != FIRST_ROW or area.c1 != 1 or area.ncols != 5:
            raise ShimError('Повреждён диапазон вводных OK_IntroArea. Восстановите книгу из резервной копии.')
        return area.r2

    def set_intro_last(self, last):
        if last < FIRST_ROW - 1:
            raise ShimError('Неверная граница вводных пунктов.')
        refers = '=0' if last < FIRST_ROW else f"='{self.ws.name}'!$A${FIRST_ROW}:$E${last}"
        self.book.names.add('OK_IntroArea', refers)

    def last_row(self):
        last = FIRST_ROW - 1
        for r, c in self.ws.filled_cells():
            if c <= 5 and r > last:
                last = r
        return last

    def blocks(self):
        last, intro_last = self.last_row(), self.intro_last()
        starts = []
        for r in range(FIRST_ROW, last + 1):
            key = str(self.ws.cells(r, 1).get_value2() or '').strip()
            if key.upper() == END_KEY:
                last = r - 1
                break
            if key:
                starts.append(r)
        result = []
        for i, r in enumerate(starts):
            end = (starts[i + 1] - 1) if i + 1 < len(starts) else max(last, r + BLOCK_ROWS - 1)
            vals = self.ws.range(f'A{r}:E{end}').get_value2()
            tools = [str(v[2]).strip() for v in vals if v[2] not in (None, '')]
            kind = 'intro' if r <= intro_last else ('stamp' if STAMP_TOOL in tools else 'param')
            result.append(Block(row=r, rows=end - r + 1, key=str(vals[0][0]).strip(),
                                text=str(vals[0][1] or ''), tools=tools,
                                volume=percent_text(vals[0][3]), minutes=to_minutes(vals[0][4]), kind=kind,
                                extra_text=[str(v[1]) for v in vals[1:] if v[1] not in (None, '')]))
        return result

    def update_block(self, block, text=None, tools=None, volume=None, minutes=None):
        ws = self.ws
        if text is not None:
            number = block.number
            new = numbered_text(text, number) if number is not None else str(text).strip()
            if not new:
                raise ShimError('Текст пункта не может быть пустым.')
            ws.cells(block.row, 2).set_value2(new)
        if tools is not None:
            tools = [t for t in (str(x).strip() for x in tools) if t]
            if len(tools) > block.rows:
                raise ShimError(f'В блоке {block.rows} строки: не больше {block.rows} средств ТО.')
            for i in range(block.rows):
                ws.cells(block.row + i, 3).set_value2(tools[i] if i < len(tools) else None)
        if volume is not None:
            ws.cells(block.row, 4).set_number_format('@')
            ws.cells(block.row, 4).set_value2(percent_text(volume) or None)
        if minutes is not None:
            num = to_minutes(minutes) if minutes != '' else None
            if minutes != '' and num is None:
                raise ShimError('Время должно быть числом (минуты).')
            if num is not None and num < 0:
                raise ShimError('Время не может быть отрицательным.')
            ws.cells(block.row, 5).set_value2(num)

    def add_intro(self, after=None):
        """Как OK_AddIntroPoint: новый блок из 4 строк после выбранного вводного (или в конец вводных)."""
        last = self.intro_last()
        intro = [b for b in self.blocks() if b.kind == 'intro']
        if len(intro) >= 1000:
            raise ShimError('Предел — 1000 блоков.')
        insert_at = last + 1
        if after is not None and after.kind == 'intro':
            insert_at = after.row + after.rows
        if intro:
            tail = intro[-1]
            insert_at = min(insert_at, max(last + 1, tail.row + BLOCK_ROWS)) if after is None else insert_at
        insert_at = max(insert_at, FIRST_ROW)
        self.ws.range(f'A{insert_at}:E{insert_at + BLOCK_ROWS - 1}').insert()
        self.ws.range(f'A{insert_at}:E{insert_at + BLOCK_ROWS - 1}').clear()
        self.ws.cells(insert_at, 1).set_value2('новый')
        self.ws.cells(insert_at, 4).set_number_format('@')
        self.ws.cells(insert_at, 4).set_value2('100%')
        self.set_intro_last(max(last, FIRST_ROW - 1) + BLOCK_ROWS)
        self.set_pending()
        self.renumber_intro()
        self.update_summary()
        return insert_at

    def delete_intro(self, block):
        if block.kind != 'intro':
            raise ShimError('Удалять можно только вводный пункт, а не параметр чертежа.')
        last = self.intro_last()
        self.set_pending()
        self.ws.range(f'A{block.row}:E{block.row + block.rows - 1}').delete()
        self.set_intro_last(last - block.rows)
        self.renumber_intro()
        self.update_summary()

    def renumber_intro(self):
        n = 0
        for r in range(FIRST_ROW, self.intro_last() + 1):
            if str(self.ws.cells(r, 1).get_value2() or '').strip():
                n += 1
                self.ws.cells(r, 1).set_value2(f'{n:02d}')
                text = numbered_text(self.ws.cells(r, 2).get_value2(), n)
                self.ws.cells(r, 2).set_value2(text or None)

    def update_summary(self):
        """Как OK_UpdateSummary: заполненные, первый баллон, пустые, без начального текста."""
        groups = empty = invalid = 0
        for b in (x for x in self.blocks() if x.kind == 'intro'):
            vals = self.ws.range(f'B{b.row}:E{b.row + b.rows - 1}').get_value2()
            any_value = any(str(v).strip() and not (j == 2 and percent_text(v) == '100%')
                            for row in vals for j, v in enumerate(row) if v is not None)
            if not any_value:
                empty += 1
            elif not str(vals[0][0] or '').strip():
                invalid += 1
            else:
                groups += 1
        for r, (label, value) in enumerate((('Заполненных пунктов', groups), ('Первый баллон', groups + 1),
                                            ('Пустых блоков', empty), ('Нет начального текста', invalid)), 24):
            self.ws.cells(r, 8).set_value2(label)
            self.ws.cells(r, 9).set_value2(value)
        return {'filled': groups, 'first_balloon': groups + 1, 'empty': empty, 'invalid': invalid}

    def add_stamp(self, minutes=None):
        """Клеймо ОТК без номера: номер назначается автоматически как последний свободный (максимум + 1)."""
        blocks = self.blocks()
        if any(b.kind == 'stamp' for b in blocks):
            raise ShimError('Клеймо ОТК уже есть в карте.')
        number = max([b.number for b in blocks if b.number is not None] or [0]) + 1
        end = blocks[-1].row + blocks[-1].rows if blocks else self.intro_last() + 1
        self.ws.range(f'A{end}:E{end + BLOCK_ROWS - 1}').insert()
        self.ws.range(f'A{end}:E{end + BLOCK_ROWS - 1}').clear()
        self.ws.cells(end, 1).set_value2(str(number))
        self.ws.cells(end, 2).set_value2(f'{number}. Клеймить ударным способом.')
        self.ws.cells(end, 3).set_value2(STAMP_TOOL)
        self.ws.cells(end, 4).set_value2('100%')
        self.ws.cells(end, 5).set_value2(to_minutes(minutes))
        return number

    # --- итоги и проверки
    def totals(self, blocks=None):
        blocks = self.blocks() if blocks is None else blocks
        out = {'items': len(blocks), 'intro': 0, 'param': 0, 'stamp': 0, 'minutes': 0.0, 'no_time': 0}
        for b in blocks:
            out[b.kind] += 1
            if b.minutes is None:
                out['no_time'] += 1
            else:
                out['minutes'] += b.minutes
        return out

    def problems(self, for_macro2=False):
        issues = []
        blocks = self.blocks()
        seen = {}
        for b in blocks:
            if b.kind != 'intro' and b.number is not None:
                if b.number in seen:
                    issues.append(('error', f'Номер {b.number} повторяется (строки {seen[b.number]} и {b.row}).'))
                seen[b.number] = b.row
            if not strip_number(b.text) and b.kind != 'intro':
                issues.append(('error', f'Пункт {b.key}: нет текста.'))
            if not b.tools and strip_number(b.text):
                issues.append(('warning', f'Пункт {b.key}: не выбрано средство ТО.'))
            if b.minutes is None and strip_number(b.text):
                issues.append(('warning', f'Пункт {b.key}: не указано время.'))
        if for_macro2:
            h = self.header()
            for key in ('kd_number', 'product_name', 'mass'):
                if not h[key].strip():
                    issues.append(('error', f'Заполните «{HEADER_FIELDS[key][2]}».'))
            if not h['signatures'] or not h['signatures'][0][1].strip():
                issues.append(('error', 'Заполните фамилию разработчика («Разраб.»).'))
            if self.state != 'READY':
                issues.append(('error', 'Карта не в состоянии READY: сначала выполните шаг «Баллоны и карта».'))
        return issues
