"""Interface local do botBMG: seleção de arquivo e controle do processo."""
import os
import queue
import re
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

RAIZ = Path(__file__).resolve().parent


class Aplicativo:
    def __init__(self, janela):
        self.janela = janela
        self.eventos = queue.Queue()
        self.processo = None
        self.ocupado = False
        self.cancelar = threading.Event()
        janela.title('botBMG • Consulta de cartões')
        janela.geometry('820x680')
        janela.minsize(740, 620)
        janela.configure(bg='#f4f5f7')
        estilo = ttk.Style()
        estilo.theme_use('clam')
        estilo.configure('TFrame', background='#f4f5f7')
        estilo.configure('TLabel', background='#f4f5f7', font=('Segoe UI', 10))
        estilo.configure('TButton', font=('Segoe UI', 10), padding=8)
        painel = ttk.Frame(janela, padding=24)
        painel.pack(fill='both', expand=True)
        ttk.Label(painel, text='Consulta de cartões', font=('Segoe UI', 22, 'bold')).pack(anchor='w')
        ttk.Label(painel, text='Escolha a lista, faça login e acompanhe as consultas.').pack(anchor='w', pady=(4, 20))
        linha = ttk.Frame(painel)
        linha.pack(fill='x')
        self.arquivo = tk.StringVar(value='')
        self.caminho = ttk.Entry(linha, textvariable=self.arquivo, state='readonly')
        self.caminho.pack(side='left', fill='x', expand=True, padx=(0, 8))
        self.escolher = ttk.Button(linha, text='Escolher planilha', command=self.selecionar)
        self.escolher.pack(side='right')
        self.status = tk.StringVar(value='Pronto • CSV ou Excel com uma coluna CPF')
        ttk.Label(painel, textvariable=self.status, wraplength=650).pack(anchor='w', pady=12)
        botoes = ttk.Frame(painel)
        botoes.pack(fill='x')
        self.iniciar = ttk.Button(botoes, text='1. Abrir Chrome para login', command=self.executar)
        self.iniciar.pack(side='left')
        self.continuar = ttk.Button(botoes, text='2. Login feito • Iniciar bot', command=self.responder, state='disabled')
        self.continuar.pack(side='left', padx=8)
        self.parar = ttk.Button(botoes, text='Parar', command=self.interromper, state='disabled')
        self.parar.pack(side='left')
        self.barra = ttk.Progressbar(painel, mode='determinate')
        self.barra.pack(fill='x', pady=16)
        self.log = tk.Text(painel, height=12, state='disabled', font=('Segoe UI', 10), wrap='word', relief='flat', padx=12, pady=12)
        rodape = ttk.Frame(painel)
        rodape.pack(side='bottom', fill='x', pady=(12, 0))
        ttk.Button(rodape, text='Abrir resultados', command=self.abrir_resultados).pack(side='left')
        ttk.Button(rodape, text='Baixar planilha de consultas', command=self.baixar_planilha).pack(side='left', padx=8)
        self.log.pack(fill='both', expand=True)
        self.registrar('A sequência automática ainda precisa ser validada no portal.\nResultados: CPF, cartão, status e disponível para saque.')
        janela.protocol('WM_DELETE_WINDOW', self.fechar)
        janela.after(100, self.atualizar)

    def registrar(self, texto):
        self.log.configure(state='normal')
        self.log.insert('end', texto + '\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def selecionar(self):
        nome = filedialog.askopenfilename(title='Escolha sua lista de CPFs', filetypes=[('Planilhas', '*.xlsx *.csv')])
        if nome:
            self.arquivo.set(nome)
            self.status.set('Planilha selecionada. Clique em Iniciar consultas.')

    def executar(self):
        if self.ocupado:
            return
        if not Path(self.arquivo.get()).is_file():
            messagebox.showerror('Planilha', 'Selecione um arquivo CSV ou XLSX existente.')
            return
        self.ocupado = True
        self.cancelar.clear()
        self.iniciar.configure(state='disabled')
        self.escolher.configure(state='disabled')
        self.parar.configure(state='normal')
        self.barra['value'] = 0
        arquivo = self.arquivo.get()
        threading.Thread(target=self.trabalhar, args=(arquivo,), daemon=True).start()

    def comando(self, argumentos, interativo=False):
        if self.cancelar.is_set():
            raise InterruptedError('Execução interrompida.')
        ambiente = dict(os.environ, PYTHONIOENCODING='utf-8', PYTHONUNBUFFERED='1')
        self.processo = subprocess.Popen(argumentos, cwd=RAIZ, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                                        stderr=subprocess.STDOUT, text=True, encoding='utf-8', errors='replace',
                                        env=ambiente, creationflags=subprocess.CREATE_NO_WINDOW)
        processo = self.processo
        for linha in processo.stdout:
            texto = linha.strip()
            if interativo and texto.startswith('BOT_PROMPT|'):
                texto = texto.removeprefix('BOT_PROMPT|')
                self.eventos.put(('texto', texto))
                self.eventos.put(('pergunta', 'fechar' if 'para fechar' in texto else 'login'))
            elif texto:
                self.eventos.put(('texto', texto))
        codigo = processo.wait()
        self.processo = None
        if self.cancelar.is_set():
            raise InterruptedError('Execução interrompida. O progresso salvo será retomado.')
        if codigo:
            raise RuntimeError('Não foi possível concluir. Veja os detalhes acima.')

    def trabalhar(self, arquivo):
        try:
            python_base = str(Path(sys.executable).with_name('python.exe'))
            python = RAIZ / '.venv' / 'Scripts' / 'python.exe'
            if not python.exists():
                self.eventos.put(('texto', 'Preparando o ambiente na primeira utilização…'))
                self.comando([python_base, '-m', 'venv', str(RAIZ / '.venv')])
            # A checagem não instala novamente em cada execução.
            checagem = subprocess.run([str(python), '-c', 'import importlib.util; raise SystemExit(not all(importlib.util.find_spec(m) for m in ("playwright", "openpyxl")))'], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
            if checagem.returncode:
                self.eventos.put(('texto', 'Instalando componentes. Aguarde…'))
                self.comando([str(python), '-m', 'pip', 'install', '-r', str(RAIZ / 'requirements.txt')])
            self.comando([str(python), '-u', str(RAIZ / 'bot.py'), arquivo], interativo=True)
            self.eventos.put(('fim', 'Execução encerrada. Confira os resultados e as mensagens acima.'))
        except Exception as erro:
            self.eventos.put(('fim', str(erro)))

    def responder(self):
        if self.processo and self.processo.poll() is None:
            try:
                self.processo.stdin.write('\n')
                self.processo.stdin.flush()
                self.continuar.configure(state='disabled')
                self.status.set('Processando…')
            except (OSError, ValueError):
                self.status.set('O processo já foi encerrado.')

    def interromper(self):
        self.cancelar.set()
        if self.processo and self.processo.poll() is None:
            # Encerra também o navegador filho; resultados já gravados são preservados.
            subprocess.run(['taskkill', '/PID', str(self.processo.pid), '/T', '/F'], capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
        self.continuar.configure(state='disabled')
        self.status.set('Interrompendo…')

    def abrir_resultados(self):
        pasta = RAIZ / 'saida'
        pasta.mkdir(exist_ok=True)
        relatorio = pasta / 'relatorio_consultas_bmg.xlsx'
        os.startfile(relatorio if relatorio.exists() else pasta)

    def baixar_planilha(self):
        if self.ocupado:
            messagebox.showinfo('Consultas em andamento', 'Aguarde terminar ou clique em Parar antes de gerar o relatório.')
            return
        origem = RAIZ / 'saida' / 'progresso.json'
        if not origem.exists():
            messagebox.showinfo('Planilha de consultas', 'Ainda não há relatório. Inicie uma consulta primeiro.')
            return
        destino = filedialog.asksaveasfilename(
            parent=self.janela,
            title='Salvar planilha de consultas',
            initialfile='relatorio_consultas_bmg.xlsx',
            defaultextension='.xlsx',
            filetypes=[('Planilha Excel', '*.xlsx')],
        )
        if not destino:
            return
        try:
            python = RAIZ / '.venv' / 'Scripts' / 'python.exe'
            resultado = subprocess.run([str(python), str(RAIZ / 'relatorio.py'), self.arquivo.get(), destino],
                                       capture_output=True, text=True, encoding='utf-8', errors='replace',
                                       env=dict(os.environ, PYTHONIOENCODING='utf-8'), creationflags=subprocess.CREATE_NO_WINDOW)
            if resultado.returncode:
                raise OSError(resultado.stderr.strip() or 'Erro ao gerar o relatório Excel.')
            self.status.set(resultado.stdout.strip())
            messagebox.showinfo('Planilha salva', f'Arquivo salvo em:\n{destino}\n\nAbas: Resumo, Consultas, Cartoes, Pendencias e Historico.\nInclui os CPFs pendentes e os motivos dos erros.')
        except (OSError, ValueError) as erro:
            messagebox.showerror('Não foi possível salvar', f'Escolha uma pasta disponível e feche o arquivo no Excel.\n\n{erro}')

    def atualizar(self):
        try:
            while True:
                tipo, texto = self.eventos.get_nowait()
                if tipo == 'texto':
                    self.registrar(texto)
                    self.status.set(texto[:160])
                    progresso = re.search(r'Consulta (\d+)/(\d+) salva', texto)
                    if progresso:
                        self.barra['maximum'] = int(progresso[2])
                        self.barra['value'] = int(progresso[1])
                elif tipo == 'pergunta':
                    self.continuar.configure(state='normal', text='Fechar navegador' if texto == 'fechar' else '2. Login feito • Iniciar bot')
                elif tipo == 'fim':
                    self.ocupado = False
                    self.status.set(texto)
                    self.registrar(texto)
                    self.iniciar.configure(state='normal')
                    self.escolher.configure(state='normal')
                    self.parar.configure(state='disabled')
                    self.continuar.configure(state='disabled')
        except queue.Empty:
            pass
        self.janela.after(100, self.atualizar)

    def fechar(self):
        if self.ocupado:
            if not messagebox.askyesno('Fechar', 'Interromper as consultas e fechar? O progresso salvo será mantido.'):
                return
            self.interromper()
        self.janela.destroy()


if __name__ == '__main__':
    janela = tk.Tk()
    Aplicativo(janela)
    janela.mainloop()
