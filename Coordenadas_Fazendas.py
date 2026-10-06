# Decompiled with PyLingual (https://pylingual.io)
# Internal filename: 'app.py'
# Bytecode version: 3.14rc3 (3627)
# Source timestamp: 1970-01-01 00:00:00 UTC (0)

"""Aplicativo Windows para gerar coordenadas centrais de KML/KMZ/ZIP/SHP."""
from __future__ import annotations
import os
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
try:
    from excel_export import export_xlsx
    from kml_processor import SUPPORTED_SUFFIXES, calculate_feature, collect_documents, extract_features, group_features, supported_files_in_folder
    from reference_data import enrich_farm_result, load_reference_data
    DEPENDENCY_ERROR = None
except ModuleNotFoundError as exc:
    DEPENDENCY_ERROR = exc
APP_TITLE = 'Coordenadas de Fazendas'
BG = '#F4F8F2'
DARK_GREEN = '#3B7D23'
GREEN = '#4EA72E'
LIGHT_GREEN = '#B4E5A2'
HEADER_GREEN = '#8ED973'
TEXT = '#20331C'
MUTED = '#5D7156'
SIDEBAR_GREEN = '#155A25'
SIDEBAR_ACTIVE = '#477D4D'
SOFT_GREEN = '#E7F1E2'
class KmlCoordinatesApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry('1180x760')
        self.minsize(1000, 680)
        self.configure(bg=BG)
        self.selected_paths = []
        self.processing = False
        self.output_path = tk.StringVar(value=str(self._default_output()))
        self.status_text = tk.StringVar(value='Adicione os arquivos para começar.')
        self._configure_style()
        self._build_ui()
        if DEPENDENCY_ERROR:
            self.after(100, self._show_dependency_error)
    def _default_output(self) -> Path:
        desktop = Path.home() / 'Desktop'
        base = desktop if desktop.exists() else Path.home()
        stamp = datetime.now().strftime('%Y%m%d_%H%M')
        return base / f'coordenadas_fazendas_talhoes_{stamp}.xlsx'
    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure('TFrame', background=BG)
        style.configure('Card.TFrame', background='white', relief='flat', bordercolor=HEADER_GREEN)
        style.configure('TLabel', background=BG, foreground=TEXT, font=('Segoe UI', 10))
        style.configure('Card.TLabel', background='white', foreground=TEXT, font=('Segoe UI', 10))
        style.configure('Title.TLabel', background=DARK_GREEN, foreground='white', font=('Segoe UI Semibold', 20))
        style.configure('Subtitle.TLabel', background=DARK_GREEN, foreground='#EAF5E6', font=('Segoe UI', 10))
        style.configure('Primary.TButton', font=('Segoe UI Semibold', 10), padding=(16, 10), background=GREEN, foreground='white', bordercolor=DARK_GREEN)
        style.map('Primary.TButton', background=[('active', DARK_GREEN), ('disabled', '#A8C69C')])
        style.configure('Secondary.TButton', font=('Segoe UI', 10), padding=(12, 9), background=LIGHT_GREEN, foreground=TEXT, bordercolor=HEADER_GREEN)
        style.map('Secondary.TButton', background=[('active', HEADER_GREEN)])
        style.configure('Treeview', font=('Segoe UI', 9), rowheight=27, background='white', fieldbackground='white', foreground=TEXT, bordercolor=HEADER_GREEN)
        style.map('Treeview', background=[('selected', GREEN)], foreground=[('selected', 'white')])
        style.configure('Treeview.Heading', font=('Segoe UI Semibold', 9), background=HEADER_GREEN, foreground='#000000')
        style.map('Treeview.Heading', background=[('active', LIGHT_GREEN)])
        style.configure('Horizontal.TProgressbar', background=GREEN, troughcolor=LIGHT_GREEN)
        style.configure('Green.TLabelframe', background=BG, bordercolor=HEADER_GREEN, relief='solid')
        style.configure('Green.TLabelframe.Label', background=BG, foreground=DARK_GREEN, font=('Segoe UI Semibold', 10))
    def _build_ui(self) -> None:
        shell = tk.Frame(self, bg='white')
        shell.pack(fill='both', expand=True)
        self.sidebar = tk.Frame(shell, bg=SIDEBAR_GREEN, width=205)
        self.sidebar.pack(side='left', fill='y')
        self.sidebar.pack_propagate(False)
        self._build_sidebar()
        self.content = tk.Frame(shell, bg='white')
        self.content.pack(side='left', fill='both', expand=True)
        self._show_home()
    def _build_sidebar(self) -> None:
        brand = tk.Canvas(self.sidebar, width=165, height=120, bg=SIDEBAR_GREEN, highlightthickness=0)
        brand.pack(pady=(30, 2))
        brand.create_polygon(66, 34, 82, 66, 98, 34, fill='white', outline='white')
        brand.create_oval(65, 16, 99, 50, fill='white', outline='white')
        brand.create_oval(76, 27, 88, 39, fill=SIDEBAR_GREEN, outline=SIDEBAR_GREEN)
        brand.create_arc(45, 60, 120, 98, start=10, extent=160, style='arc', width=5, outline='#8ED973')
        brand.create_arc(52, 69, 113, 102, start=10, extent=160, style='arc', width=5, outline='#65B54A')
        brand.create_line(82, 64, 82, 99, fill='#B4E5A2', width=4)
        tk.Label(self.sidebar, text='Coordenadas\nde Fazendas', bg=SIDEBAR_GREEN, fg='white', font=('Segoe UI Semibold', 15), justify='center').pack(pady=(0, 30))
        create_button = tk.Button(self.sidebar, text='  ⊕   Criar', command=self._show_processing, bg=SIDEBAR_ACTIVE, activebackground='#5B9160', fg='white', activeforeground='white', font=('Segoe UI Semibold', 12), relief='flat', bd=0, cursor='hand2', anchor='w', padx=16, pady=13)
        create_button.pack(fill='x', padx=12)
        decoration = tk.Canvas(self.sidebar, bg=SIDEBAR_GREEN, highlightthickness=0)
        decoration.pack(side='bottom', fill='both', expand=True)
        decoration.create_polygon(0, 360, 205, 285, 205, 430, 0, 430, fill='#246B31', outline='')
        decoration.create_polygon(0, 400, 205, 330, 205, 470, 0, 470, fill='#1C602B', outline='')
    def _clear_content(self, background: str='white') -> None:
        for child in self.content.winfo_children():
            child.destroy()
        self.content.configure(bg=background)
    def _draw_home_icon(self, parent) -> None:
        canvas = tk.Canvas(parent, width=190, height=165, bg='white', highlightthickness=0)
        canvas.pack()
        canvas.create_oval(20, 2, 170, 152, fill=SOFT_GREEN, outline='')
        canvas.create_polygon(73, 52, 95, 98, 117, 52, fill=DARK_GREEN, outline=DARK_GREEN)
        canvas.create_oval(72, 29, 118, 74, fill=DARK_GREEN, outline=DARK_GREEN)
        canvas.create_oval(87, 43, 103, 59, fill=SOFT_GREEN, outline=SOFT_GREEN)
        canvas.create_arc(25, 82, 165, 155, start=4, extent=175, style='pieslice', fill='#4EA72E', outline='')
        canvas.create_arc(31, 100, 165, 158, start=6, extent=170, style='arc', width=6, outline='#286D2C')
        canvas.create_arc(70, 98, 173, 155, start=5, extent=170, style='arc', width=6, outline='white')
        canvas.create_oval(25, 68, 55, 84, fill='#B4E5A2', outline='')
        canvas.create_oval(42, 62, 72, 84, fill='#B4E5A2', outline='')
        canvas.create_oval(126, 59, 158, 80, fill='#B4E5A2', outline='')
    def _show_home(self) -> None:
        if self.processing:
            messagebox.showinfo('Processamento em andamento', 'Aguarde a conclusão antes de voltar à tela inicial.')
            return
        else:
            self._clear_content('white')
            home = tk.Frame(self.content, bg='white')
            home.pack(fill='both', expand=True)
            center = tk.Frame(home, bg='white')
            center.place(relx=0.5, rely=0.48, anchor='center')
            self._draw_home_icon(center)
            tk.Label(center, text='Bem-vindo', bg='white', fg='#244F25', font=('Segoe UI Semibold', 31)).pack(pady=(0, 8))
            tk.Label(center, text='Crie um novo processamento para gerar a planilha de coordenadas de fazendas\na partir dos seus arquivos geoespaciais.', bg='white', fg='#4D554B', font=('Segoe UI', 12), justify='center').pack(pady=(0, 24))
            tk.Button(center, text='⊕   Criar novo processamento', command=self._show_processing, bg=DARK_GREEN, activebackground='#2D671C', fg='white', activeforeground='white', font=('Segoe UI Semibold', 13), relief='flat', bd=0, cursor='hand2', padx=95, pady=14).pack()
            separator = tk.Frame(center, bg='white')
            separator.pack(fill='x', pady=20)
            tk.Frame(separator, bg='#D8DED5', height=1, width=225).pack(side='left', pady=9)
            tk.Label(separator, text='ou', bg='white', fg='#727A70', font=('Segoe UI', 10)).pack(side='left', padx=20)
            tk.Frame(separator, bg='#D8DED5', height=1, width=225).pack(side='left', pady=9)
            upload = tk.Canvas(center, width=700, height=135, bg='#FBFDF9', highlightthickness=0, cursor='hand2')
            upload.pack()
            upload.create_rectangle(3, 3, 697, 132, outline='#6B9C66', width=2, dash=(6, 5))
            upload.create_rectangle(112, 38, 153, 88, outline=DARK_GREEN, width=3)
            upload.create_polygon(140, 38, 153, 52, 140, 52, fill='#DCEBD7', outline=DARK_GREEN)
            upload.create_line(120, 61, 144, 61, fill='#A9CBA0', width=3)
            upload.create_line(120, 71, 140, 71, fill='#A9CBA0', width=3)
            upload.create_oval(139, 75, 171, 107, fill=GREEN, outline=GREEN)
            upload.create_line(155, 98, 155, 84, fill='white', width=3, arrow=tk.LAST)
            upload.create_text(205, 51, text='Adicione arquivos para começar', anchor='w', fill='#315334', font=('Segoe UI Semibold', 13))
            upload.create_text(205, 78, text='Clique nesta área para selecionar os arquivos.', anchor='w', fill='#5F665C', font=('Segoe UI', 10))
            upload.create_text(205, 104, text='Formatos aceitos: SHP, KML, KMZ e ZIP', anchor='w', fill='#7A8278', font=('Segoe UI', 9))
            upload.bind('<Button-1>', self._home_add_files)
            tk.Label(center, text='♢  Seus arquivos são processados localmente e não são armazenados.', bg='white', fg='#899187', font=('Segoe UI', 9)).pack(pady=(24, 0))
    def _show_processing(self) -> None:
        self._clear_content(BG)
        header = tk.Frame(self.content, bg='white', height=88)
        header.pack(fill='x')
        header.pack_propagate(False)
        title_row = tk.Frame(header, bg='white')
        title_row.pack(fill='x', padx=28, pady=(16, 0))
        tk.Label(title_row, text='Novo processamento', bg='white', fg='#244F25', font=('Segoe UI Semibold', 20)).pack(side='left')
        tk.Button(title_row, text='←  Voltar ao início', command=self._show_home, bg='white', activebackground='#F1F6EF', fg=DARK_GREEN, activeforeground=DARK_GREEN, font=('Segoe UI', 9), relief='flat', cursor='hand2').pack(side='right')
        tk.Label(header, text='Selecione os arquivos e gere a planilha em SIRGAS 2000 com fazendas, talhões e posto meteorológico mais próximo.', bg='white', fg=MUTED, font=('Segoe UI', 9)).pack(anchor='w', padx=29, pady=(2, 0))
        body = ttk.Frame(self.content, padding=(26, 18, 26, 20))
        body.pack(fill='both', expand=True)
        actions = ttk.Frame(body)
        actions.pack(fill='x', pady=(0, 12))
        ttk.Button(actions, text='Adicionar arquivos', style='Secondary.TButton', command=self.add_files).pack(side='left', padx=(0, 8))
        ttk.Button(actions, text='Adicionar pasta', style='Secondary.TButton', command=self.add_folder).pack(side='left', padx=(0, 8))
        ttk.Button(actions, text='Remover selecionados', style='Secondary.TButton', command=self.remove_selected).pack(side='left', padx=(0, 8))
        ttk.Button(actions, text='Limpar lista', style='Secondary.TButton', command=self.clear_list).pack(side='left')
        table_card = ttk.Frame(body, style='Card.TFrame', padding=1)
        table_card.pack(fill='both', expand=True)
        self.tree = ttk.Treeview(table_card, columns=('tipo', 'caminho'), show='headings', selectmode='extended')
        self.tree.heading('tipo', text='Tipo')
        self.tree.heading('caminho', text='Arquivo ou pasta')
        self.tree.column('tipo', width=90, anchor='center', stretch=False)
        self.tree.column('caminho', width=720, anchor='w')
        scrollbar = ttk.Scrollbar(table_card, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        for path in self.selected_paths:
            self.tree.insert('', 'end', values=(path.suffix.upper().lstrip('.'), str(path)))
        mode_frame = ttk.LabelFrame(body, text='Resultado da planilha', padding=(12, 8), style='Green.TLabelframe')
        mode_frame.pack(fill='x', pady=(12, 0))
        ttk.Label(mode_frame, text='O Excel terá duas abas: Fazendas (áreas agrupadas, posto mais próximo e distância até o posto) e Talhões (cada área individual). Referência: SIRGAS 2000. Base completa: 200 localizações meteorológicas.', wraplength=850).pack(anchor='w', pady=3)
        output_frame = ttk.Frame(body)
        output_frame.pack(fill='x', pady=(14, 8))
        ttk.Label(output_frame, text='Salvar planilha em:').pack(anchor='w', pady=(0, 4))
        output_row = ttk.Frame(output_frame)
        output_row.pack(fill='x')
        self.output_entry = ttk.Entry(output_row, textvariable=self.output_path)
        self.output_entry.pack(side='left', fill='x', expand=True, padx=(0, 8))
        ttk.Button(output_row, text='Escolher local', style='Secondary.TButton', command=self.choose_output).pack(side='right')
        footer = ttk.Frame(body)
        footer.pack(fill='x', pady=(8, 0))
        self.progress = ttk.Progressbar(footer, mode='determinate', length=280)
        self.progress.pack(side='left', fill='x', expand=True, padx=(0, 18))
        self.process_button = ttk.Button(footer, text='Gerar planilha Excel', style='Primary.TButton', command=self.start_processing)
        self.process_button.pack(side='right')
        ttk.Label(body, textvariable=self.status_text, foreground=MUTED).pack(anchor='w', pady=(8, 0))
        if DEPENDENCY_ERROR:
            self.process_button.configure(state='disabled')
    def _home_add_files(self, _event=None) -> None:
        files = self._select_files_dialog()
        if not files:
            return
        else:
            self._show_processing()
            self._add_paths((Path(item) for item in files))
    def _show_dependency_error(self) -> None:
        messagebox.showerror('Dependência ausente', 'O módulo openpyxl não está instalado. Abra o aplicativo pelo arquivo Executar_Aplicativo.bat.')
        if hasattr(self, 'process_button'):
            self.process_button.configure(state='disabled')
    def _select_files_dialog(self):
        return filedialog.askopenfilenames(title='Selecione KML, KMZ, ZIP ou SHP', filetypes=[('Arquivos geográficos', '*.kml *.kmz *.zip *.shp'), ('Todos os arquivos', '*.*')])
    def add_files(self) -> None:
        files = self._select_files_dialog()
        self._add_paths((Path(item) for item in files))
    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title='Selecione uma pasta')
        if not folder:
            return
        else:
            files = supported_files_in_folder(Path(folder))
            if not files:
                messagebox.showinfo('Nenhum arquivo', 'A pasta não contém arquivos KML, KMZ, ZIP ou SHP.')
                return
            else:
                self._add_paths(files)
    def _add_paths(self, paths) -> None:
        existing = {str(path.resolve()).lower() for path in self.selected_paths}
        added = 0
        for path in paths:
            if not path.is_file() or path.suffix.lower() not in SUPPORTED_SUFFIXES:
                continue
            else:
                key = str(path.resolve()).lower()
                if key in existing:
                    continue
                else:
                    existing.add(key)
                    self.selected_paths.append(path)
                    self.tree.insert('', 'end', values=(path.suffix.upper().lstrip('.'), str(path)))
                    added += 1
        self.status_text.set(f'{len(self.selected_paths)} entrada(s) selecionada(s). {added} adicionada(s) agora.')
    def remove_selected(self) -> None:
        indexes = sorted((self.tree.index(item) for item in self.tree.selection()), reverse=True)
        for index in indexes:
            self.selected_paths.pop(index)
        for item in self.tree.selection():
            self.tree.delete(item)
        self.status_text.set(f'{len(self.selected_paths)} entrada(s) selecionada(s).')
    def clear_list(self) -> None:
        self.selected_paths.clear()
        self.tree.delete(*self.tree.get_children())
        self.progress['value'] = 0
        self.status_text.set('Lista limpa. Adicione os arquivos para começar.')
    def choose_output(self) -> None:
        selected = filedialog.asksaveasfilename(title='Salvar planilha', defaultextension='.xlsx', filetypes=[('Planilha Excel', '*.xlsx')], initialfile=Path(self.output_path.get()).name, initialdir=str(Path(self.output_path.get()).parent))
        if selected:
            self.output_path.set(selected)
    def start_processing(self) -> None:
        if DEPENDENCY_ERROR:
            self._show_dependency_error()
            return
        else:
            if not self.selected_paths:
                messagebox.showwarning('Nenhum arquivo', 'Adicione pelo menos um arquivo KML, KMZ, ZIP, SHP ou uma pasta.')
                return
            else:
                output = Path(self.output_path.get().strip())
                if output.suffix.lower() != '.xlsx':
                    output = output.with_suffix('.xlsx')
                    self.output_path.set(str(output))
                self.process_button.configure(state='disabled')
                self.processing = True
                self.progress['value'] = 0
                self.status_text.set('Lendo os arquivos...')
                paths = list(self.selected_paths)
                threading.Thread(target=self._process_worker, args=(paths, output), daemon=True).start()
    def _process_worker(self, paths: list[Path], output: Path) -> None:
        try:
            reference_data = load_reference_data()
            documents = collect_documents(paths)
            if not documents:
                raise ValueError('Nenhum arquivo geográfico compatível foi encontrado nas entradas selecionadas.')
            else:
                errors = []
                features = []
                for document in documents:
                    try:
                        features.extend(extract_features(document))
                    except Exception as exc:
                        errors.append((document.origin, str(exc)))
                    else:
                        pass
                if not features:
                    raise ValueError('Nenhuma fazenda ou geometria foi encontrada nos arquivos selecionados.')
                else:
                    farm_features = group_features(features)
                    farm_results = []
                    plot_results = []
                    total = len(farm_features) + len(features)
                    progress_index = 0
                    for feature in farm_features:
                        try:
                            farm_results.append(enrich_farm_result(calculate_feature(feature), reference_data))
                        except Exception as exc:
                            errors.append((f'Fazenda: {feature.arquivo} — {feature.fazenda}', str(exc)))
                        progress_index += 1
                        self.after(0, self._update_progress, int(progress_index / total * 100), f'Fazenda {feature.fazenda}')
                    for feature in features:
                        try:
                            plot_results.append(calculate_feature(feature))
                        except Exception as exc:
                            errors.append((f'Talhão: {feature.arquivo} — {feature.fazenda} — {feature.talhao}', str(exc)))
                        progress_index += 1
                        self.after(0, self._update_progress, int(progress_index / total * 100), f'Talhão {feature.talhao or feature.fazenda}')
                    if not farm_results and (not plot_results):
                        raise ValueError('Nenhum arquivo pôde ser calculado. Consulte os arquivos selecionados.')
                    else:
                        export_xlsx(farm_results, plot_results, errors, output)
                        self.after(0, self._processing_done, output, len(farm_results), len(plot_results), len(errors))
        except Exception as exc:
            self.after(0, self._processing_failed, str(exc))
    def _update_progress(self, value: int, message: str) -> None:
        if self.progress.winfo_exists():
            self.progress['value'] = value
        self.status_text.set(message)
    def _processing_done(self, output: Path, farm_count: int, plot_count: int, error_count: int) -> None:
        self.processing = False
        self.progress['value'] = 100
        self.process_button.configure(state='normal')
        summary = f'{farm_count} fazenda(s) e {plot_count} talhão(ões)'
        self.status_text.set(f'Concluído: {summary}; {error_count} erro(s).')
        answer = messagebox.askyesno('Planilha criada', f'A planilha foi criada com {summary}.\nErros: {error_count}.\n\nDeseja abrir o arquivo agora?')
        if answer:
            try:
                os.startfile(output)
            except Exception:
                os.startfile(output.parent)
    def _processing_failed(self, message: str) -> None:
        self.processing = False
        self.process_button.configure(state='normal')
        self.status_text.set('Não foi possível concluir o processamento.')
        messagebox.showerror('Erro', message)
if __name__ == '__main__':
    KmlCoordinatesApp().mainloop()
