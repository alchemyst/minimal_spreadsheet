#!/usr/bin/env python

import re
import sys
from collections import ChainMap
import math

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTableWidget,
        QTableWidgetItem, QItemDelegate, QLineEdit)


cellre = re.compile(r'\b[A-Z][0-9]\b')


def cellname(i, j):
    return f'{chr(ord("A")+j)}{i+1}'


class SpreadSheetDelegate(QItemDelegate):
    def __init__(self, parent=None):
        super(SpreadSheetDelegate, self).__init__(parent)

    def createEditor(self, parent, styleOption, index):
        editor = QLineEdit(parent)
        editor.editingFinished.connect(self.commitAndCloseEditor)
        return editor

    def commitAndCloseEditor(self):
        editor = self.sender()
        self.commitData.emit(editor)
        self.closeEditor.emit(editor, QItemDelegate.NoHint)

    def setEditorData(self, editor, index):
        editor.setText(index.model().data(index, Qt.EditRole))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.text())


class SpreadSheetItem(QTableWidgetItem):
    def __init__(self, siblings):
        super(SpreadSheetItem, self).__init__()
        self.siblings = siblings
        self.value = 0
        self.deps = set()
        self.reqs = set()

    def formula(self):
        return super().data(Qt.DisplayRole)

    def data(self, role):
        if role == Qt.EditRole:
            return self.formula()
        if role == Qt.DisplayRole:
            return self.display()

        return super(SpreadSheetItem, self).data(role)

    def calculate(self):
        formula = self.formula()

        if formula is None or formula == '':
            self.value = 0
            return

        currentreqs = set(cellre.findall(formula))

        name = cellname(self.row(), self.column())

        # Add this cell to the new requirement's dependents
        for r in currentreqs - self.reqs:
            self.siblings[r].deps.add(name)
        # Add remove this cell from dependents no longer referenced
        for r in self.reqs - currentreqs:
            self.siblings[r].deps.remove(name)

        # Look up the values of our required cells
        reqvalues = {r: self.siblings[r].value for r in currentreqs}
        # Build an environment with these values and basic math functions
        environment = ChainMap(math.__dict__, reqvalues)
        # Note that eval is DANGEROUS and should not be used in production
        self.value = eval(formula, {}, environment)

        self.reqs = currentreqs

    def propagate(self):
        for d in self.deps:
            self.siblings[d].calculate()
            self.siblings[d].propagate()

    def display(self):
        self.calculate()
        self.propagate()
        return str(self.value)


class SpreadSheet(QMainWindow):
    def __init__(self, rows, cols, parent=None):
        super(SpreadSheet, self).__init__(parent)

        self.rows = rows
        self.cols = cols

        self.cells = {}

        self.create_widgets()

    def create_widgets(self):
        table = self.table = QTableWidget(self.rows, self.cols, self)

        headers = [chr(ord('A') + j) for j in range(self.cols)]
        table.setHorizontalHeaderLabels(headers)

        table.setItemDelegate(SpreadSheetDelegate(self))

        for i in range(self.rows):
            for j in range(self.cols):
                cell = SpreadSheetItem(self.cells)
                self.cells[cellname(i, j)] = cell
                self.table.setItem(i, j, cell)

        self.setCentralWidget(table)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    sheet = SpreadSheet(5, 5)
    sheet.resize(520, 200)
    sheet.show()
    sys.exit(app.exec_())
