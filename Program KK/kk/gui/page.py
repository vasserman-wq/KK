"""Базовый класс страницы-шага."""
from __future__ import annotations

from . import base

TOTAL_STEPS = 6


class Page:
    key = ''
    title = ''
    hint = ''

    def __init__(self, app, number):
        self.app = app
        self.number = number
        self.frame = base.page_frame(app.content, number, TOTAL_STEPS, self.title, self.hint)
        self.build()

    def build(self):
        pass

    def on_show(self):
        pass

    def tick(self):
        """Вызывается каждые 150 мс (шкала макроса)."""

    def commit(self):
        """Перенести несохранённые поля формы в карту. False — ошибка ввода."""
        return True

    def go(self, key):
        return lambda: self.app.show_step(key)
