#!/usr/bin/env python
"""Отдельный процесс для автонумерации в AutoCAD."""

import sys
import os

# Добавляем путь к проекту
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.autocad_tasks.utils.numbering.gui import NumberingWindow

if __name__ == "__main__":
    app = NumberingWindow()
    app.run()  # Теперь работает, так как метод run добавлен