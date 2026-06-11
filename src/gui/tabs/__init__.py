# -*- coding: utf-8 -*-
"""Mixins de tabs da interface gráfica.

Cada módulo contém um mixin com os métodos de uma tab do ConverterApp.
A classe principal (src/gui/app.py) compõe todos os mixins.
"""

from src.gui.tabs.dashboard import DashboardTabMixin
from src.gui.tabs.convert import ConvertTabMixin
from src.gui.tabs.profiles import ProfilesTabMixin
from src.gui.tabs.batch import BatchTabMixin
from src.gui.tabs.history import HistoryTabMixin
from src.gui.tabs.settings import SettingsTabMixin
from src.gui.tabs.contabilidade import ContabilidadeTabMixin
from src.gui.tabs.banking import BankingTabMixin
from src.gui.tabs.doc_sequence import DocSequenceTabMixin
from src.gui.tabs.qrcode_fonts import QrcodeFontsTabMixin
from src.gui.tabs.automation import AutomationTabMixin

__all__ = [
    'DashboardTabMixin',
    'ConvertTabMixin',
    'ProfilesTabMixin',
    'BatchTabMixin',
    'HistoryTabMixin',
    'SettingsTabMixin',
    'ContabilidadeTabMixin',
    'BankingTabMixin',
    'DocSequenceTabMixin',
    'QrcodeFontsTabMixin',
    'AutomationTabMixin',
]
