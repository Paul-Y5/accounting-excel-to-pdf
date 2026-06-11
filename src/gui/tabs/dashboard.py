# -*- coding: utf-8 -*-
"""Tab Dashboard — resumo de actividade, acções rápidas e gráfico."""

from datetime import datetime
import tkinter as tk
from tkinter import ttk


class DashboardTabMixin:
    """Métodos da tab Dashboard."""

    def _setup_dashboard_tab(self):
        """Tab Dashboard — resumo de actividade, acções rápidas e gráfico."""
        outer = ttk.Frame(self.tab_dashboard, padding=self._PAD_OUTER)
        outer.pack(fill='both', expand=True)

        # ---- Título ----
        title_row = ttk.Frame(outer)
        title_row.pack(fill='x', pady=(0, 10))
        ttk.Label(title_row, text='Dashboard', style='Header.TLabel').pack(side='left')
        self._dash_refresh_btn = ttk.Button(
            title_row, text='Actualizar', command=self._refresh_dashboard)
        self._dash_refresh_btn.pack(side='right')

        # ---- Cards de resumo (linha superior) ----
        cards_frame = ttk.Frame(outer)
        cards_frame.pack(fill='x', pady=(0, 12))
        for col in range(4):
            cards_frame.columnconfigure(col, weight=1, uniform='card')

        self._dash_cards = {}
        card_defs = [
            ('conv_mes',   'Conversões\neste mês',        '#0078D4'),
            ('taxa',       'Taxa de\nSucesso',            '#38A169'),
            ('clientes',   'Clientes\neste mês',          '#805AD5'),
            ('series',     'Séries de\nDocumento ativas', '#D69E2E'),
        ]
        for col, (key, label, color) in enumerate(card_defs):
            card = ttk.LabelFrame(cards_frame, padding=10)
            card.grid(row=0, column=col, padx=4, sticky='nsew')
            val_lbl = tk.Label(card, text='—', font=(self._FONT_FAMILY, 22, 'bold'),
                               foreground=color)
            val_lbl.pack()
            tk.Label(card, text=label, font=(self._FONT_FAMILY, 8),
                     justify='center').pack()
            self._dash_cards[key] = val_lbl

        # ---- Fila inferior: gráfico | acções rápidas ----
        bottom = ttk.Frame(outer)
        bottom.pack(fill='both', expand=True)
        bottom.columnconfigure(0, weight=3)
        bottom.columnconfigure(1, weight=2)

        # Gráfico de barras — conversões por mês
        chart_frame = ttk.LabelFrame(bottom, text='Conversões por Mês', padding=8)
        chart_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
        self._dash_canvas = tk.Canvas(chart_frame, height=160, bg='white',
                                      highlightthickness=0)
        self._dash_canvas.pack(fill='both', expand=True)
        self._dash_canvas.bind('<Configure>', lambda e: self._redraw_chart_on_resize())

        # Coluna direita: acções rápidas + actividade recente
        right_col = ttk.Frame(bottom)
        right_col.grid(row=0, column=1, sticky='nsew')

        # Acções rápidas
        actions_frame = ttk.LabelFrame(right_col, text='Acções Rápidas', padding=8)
        actions_frame.pack(fill='x', pady=(0, 8))
        btn_defs = [
            ('Converter Ficheiro',    lambda: self.notebook.select(self.tab_convert)),
            ('Multificheiros',        lambda: self.notebook.select(self.tab_batch)),
            ('Relatório Anual',       self._generate_annual_report),
            ('Nova Série Documento',  lambda: (self.notebook.select(self.tab_settings),
                                               self._add_doc_serie())),
        ]
        for label, cmd in btn_defs:
            ttk.Button(actions_frame, text=label, command=cmd).pack(
                fill='x', pady=2)

        # Actividade recente
        recent_frame = ttk.LabelFrame(right_col, text='Actividade Recente', padding=8)
        recent_frame.pack(fill='both', expand=True)
        cols = ('data', 'ficheiro', 'ok')
        self._dash_recent = ttk.Treeview(
            recent_frame, columns=cols, show='headings',
            height=5, selectmode='none')
        self._dash_recent.heading('data',     text='Data')
        self._dash_recent.heading('ficheiro', text='Ficheiro')
        self._dash_recent.heading('ok',       text='OK')
        self._dash_recent.column('data',     width=75,  anchor='center')
        self._dash_recent.column('ficheiro', width=120, anchor='w')
        self._dash_recent.column('ok',       width=30,  anchor='center')
        self._dash_recent.pack(fill='both', expand=True)

        # Carregar dados iniciais
        self._refresh_dashboard()

    def _refresh_dashboard(self):
        """Actualiza todos os widgets do dashboard com dados recentes."""
        from src.annual_report import get_annual_data
        from src.doc_sequence import list_series
        from src.database import get_history

        now = datetime.now()
        ano = now.year
        mes = now.month

        # --- Dados do ano actual ---
        anual = get_annual_data(ano)
        mes_data = anual['by_month'][mes - 1]

        conv_mes = mes_data['conversions']
        taxa = (
            f"{mes_data['success'] / conv_mes * 100:.0f}%"
            if conv_mes else '—'
        )
        clientes_mes = mes_data['clients']
        series_ativas = len(list_series())

        self._dash_cards['conv_mes'].config(text=str(conv_mes))
        self._dash_cards['taxa'].config(text=taxa)
        self._dash_cards['clientes'].config(text=str(clientes_mes))
        self._dash_cards['series'].config(text=str(series_ativas))

        # --- Gráfico de barras (12 meses) ---
        self._dash_chart_data = anual['by_month']
        self._draw_bar_chart(anual['by_month'])

        # --- Actividade recente (últimas 5) ---
        for row in self._dash_recent.get_children():
            self._dash_recent.delete(row)
        for entry in get_history(limit=5):
            try:
                ts = datetime.fromisoformat(entry['timestamp'])
                data_str = ts.strftime('%d/%m %H:%M')
            except (ValueError, TypeError):
                data_str = entry['timestamp'][:16]
            ok_str = '✓' if entry['success'] else '✗'
            self._dash_recent.insert('', 'end', values=(
                data_str,
                entry['source_file'][:20],
                ok_str,
            ))

    def _redraw_chart_on_resize(self):
        if hasattr(self, '_dash_chart_data'):
            self._draw_bar_chart(self._dash_chart_data)

    def _draw_bar_chart(self, by_month: list):
        """Desenha um gráfico de barras simples no canvas do dashboard."""
        canvas = self._dash_canvas
        canvas.delete('all')
        canvas.update_idletasks()

        W = canvas.winfo_width() or 300
        H = canvas.winfo_height() or 160
        pad_left = 30
        pad_right = 8
        pad_top = 12
        pad_bottom = 22

        max_val = max((m['conversions'] for m in by_month), default=0)
        if max_val == 0:
            canvas.create_text(W // 2, H // 2, text='Sem dados',
                               fill='gray', font=(self._FONT_FAMILY, 9))
            return

        n = 12
        bar_area_w = W - pad_left - pad_right
        bar_w = bar_area_w / n
        chart_h = H - pad_top - pad_bottom
        mes_atual = datetime.now().month
        meses_abrev = ['J', 'F', 'M', 'A', 'M', 'J',
                       'J', 'A', 'S', 'O', 'N', 'D']

        for i, m in enumerate(by_month):
            x0 = pad_left + i * bar_w + bar_w * 0.15
            x1 = pad_left + (i + 1) * bar_w - bar_w * 0.15
            bar_h = (m['conversions'] / max_val) * chart_h if max_val else 0
            y0 = H - pad_bottom - bar_h
            y1 = H - pad_bottom

            color = '#0078D4' if (i + 1) != mes_atual else '#005a9e'
            if m['conversions']:
                canvas.create_rectangle(x0, y0, x1, y1, fill=color, outline='')
                canvas.create_text(
                    (x0 + x1) / 2, y0 - 4,
                    text=str(m['conversions']),
                    font=(self._FONT_FAMILY, 7), fill='#2d3748',
                )
            # Rótulo do mês
            canvas.create_text(
                (x0 + x1) / 2, H - pad_bottom + 8,
                text=meses_abrev[i],
                font=(self._FONT_FAMILY, 7), fill='gray',
            )

        # Eixo Y — linha de base
        canvas.create_line(pad_left - 2, H - pad_bottom,
                           W - pad_right, H - pad_bottom,
                           fill='#e2e8f0')
        # Valor máximo
        canvas.create_text(pad_left - 4, pad_top,
                           text=str(max_val), anchor='e',
                           font=(self._FONT_FAMILY, 7), fill='gray')
