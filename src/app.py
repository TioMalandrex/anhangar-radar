import os
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

sys.path.insert(0, str(Path(__file__).parent))
from ia_processor import extrair_dados
from excel_manager import adicionar_linha, garantir_excel, contar_registros_hoje
from importador import importar_planilha
import config

ROOT       = Path(__file__).parent.parent
EXCEL_PATH = ROOT / "data" / "contatos.xlsx"

# ── Paleta ───────────────────────────────────────────────────────────────────
BG      = "#0f0f1a"
CARD    = "#1a1a2e"
SURFACE = "#252540"
ACCENT  = "#7c3aed"
ACCENT2 = "#6d28d9"
TEXT    = "#e2e8f0"
TEXT2   = "#94a3b8"
INPUT   = "#2a2a3e"
BORDER  = "#3d3d5c"

FONTES   = ["Ligação", "WhatsApp", "E-mail", "Visita", "Outro"]
STATUSES = ["Sem resposta", "Não atendida", "Número inexistente",
            "Falha no contato", "Sem interesse", "Contato realizado", "Agendado"]


def _aba_para_status(status: str) -> str:
    s = status.lower()
    if any(k in s for k in ["interesse", "recusou"]):
        return "Sem Interesse"
    return "Falhas e Sem Contato"


# ─────────────────────────────────────────────────────────────────────────────

class AnhangaRadar(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Anhangá Radar")
        self.geometry("720x900")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._setup_style()
        self._build()
        self.bind_all("<Control-Return>", lambda _e: self._processar())
        chave_salva = config.carregar_api_key(ROOT)
        if chave_salva:
            self.var_apikey.set(chave_salva)
        self._atualizar_contador()

    # ── Estilo ───────────────────────────────────────────────────────────────

    def _setup_style(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure("Dark.TCombobox",
                    fieldbackground=INPUT, background=SURFACE,
                    foreground=TEXT, bordercolor=BORDER,
                    arrowcolor=TEXT2, selectbackground=ACCENT,
                    padding=6)
        s.map("Dark.TCombobox",
              fieldbackground=[("readonly", INPUT)],
              foreground=[("readonly", TEXT)],
              selectbackground=[("readonly", ACCENT)])

    # ── Layout ───────────────────────────────────────────────────────────────

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg=CARD, height=60)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="ANHANGÁ RADAR", bg=CARD, fg=TEXT,
                 font=("Segoe UI", 16, "bold")).pack(side="left", padx=20, pady=16)
        tk.Label(hdr, text="github.com/Jitterkkk  •  @jitterkkk", bg=CARD, fg=TEXT2,
                 font=("Segoe UI", 9)).pack(side="left", pady=22)

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=24, pady=14)

        # API Key
        self._titulo(body, "Chave da API  —  console.groq.com")
        api_row = tk.Frame(body, bg=BG)
        api_row.pack(fill="x", pady=(4, 0))
        self.var_apikey = tk.StringVar()
        self.ent_apikey = self._entry(api_row, show="•", var=self.var_apikey)
        self.ent_apikey.pack(side="left", fill="x", expand=True)
        self._btn(api_row, "👁", self._toggle_apikey, padx=10).pack(side="left", padx=(6, 0))

        self._sep(body)

        # Descrição
        self._titulo(body, "Descrição do contato  (Ctrl+Enter para processar)")
        self.txt_desc = tk.Text(
            body, height=7, bg=INPUT, fg=TEXT, insertbackground=TEXT,
            font=("Segoe UI", 10), bd=0, relief="flat", wrap="word",
            padx=10, pady=8,
            highlightthickness=1, highlightbackground=BORDER,
            highlightcolor=ACCENT,
        )
        self.txt_desc.pack(fill="x", pady=(4, 10))

        btn_proc_row = tk.Frame(body, bg=BG)
        btn_proc_row.pack(fill="x")
        self.btn_proc = self._btn(btn_proc_row, "Processar com IA",
                                  self._processar, accent=True, padx=20, pady=9)
        self.btn_proc.pack(side="right")

        self._sep(body)

        # Campos extraídos
        self._titulo(body, "Dados extraídos  —  revise antes de salvar")

        self.var_nome      = tk.StringVar()
        self.var_sobrenome = tk.StringVar()
        self.var_cidade    = tk.StringVar()
        self.var_telefone  = tk.StringVar()
        self.var_fonte     = tk.StringVar(value=FONTES[0])
        self.var_status    = tk.StringVar(value=STATUSES[0])
        self.var_obs       = tk.StringVar()

        grid = tk.Frame(body, bg=BG)
        grid.pack(fill="x", pady=(6, 0))
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)

        self._campo(grid, "Nome",      self.var_nome,      0, 0)
        self._campo(grid, "Sobrenome", self.var_sobrenome, 0, 1)
        self._campo(grid, "Cidade",    self.var_cidade,    1, 0)
        self._campo(grid, "Telefone",  self.var_telefone,  1, 1)
        self._campo_combo(grid, "Fonte",  self.var_fonte,  FONTES,    2, 0)
        self._campo_combo(grid, "Status", self.var_status, STATUSES,  2, 1)

        # Indicador de aba
        self.lbl_aba = tk.Label(body, text="", bg=BG, fg=TEXT2,
                                font=("Segoe UI", 8, "italic"))
        self.lbl_aba.pack(anchor="e", pady=(4, 0))
        self.var_status.trace_add("write", lambda *_: self._atualizar_aba())
        self._atualizar_aba()

        # Observações
        tk.Label(body, text="Observações", bg=BG, fg=TEXT2,
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(10, 2))
        self._entry(body, var=self.var_obs).pack(fill="x")

        # Botões de ação
        acao = tk.Frame(body, bg=BG)
        acao.pack(fill="x", pady=(22, 0))
        self._btn(acao, "Abrir Excel", self._abrir_excel,
                  padx=16, pady=8).pack(side="left")
        self._btn(acao, "Salvar na planilha", self._salvar,
                  accent=True, padx=20, pady=8).pack(side="right")

        # Importar planilha antiga
        self._sep(body)
        self._titulo(body, "Importar dados de planilha antiga")
        tk.Label(body,
                 text="Selecione um .xlsx com contatos anteriores — a IA mapeia as colunas automaticamente.",
                 bg=BG, fg=TEXT2, font=("Segoe UI", 8),
                 wraplength=640, justify="left").pack(anchor="w", pady=(2, 8))
        self.btn_importar = self._btn(body, "Selecionar arquivo .xlsx e importar",
                                      self._importar_planilha, padx=16, pady=8)
        self.btn_importar.pack(anchor="w")

        # Rodapé fixo — duas linhas
        rodape = tk.Frame(self, bg=CARD)
        rodape.pack(fill="x", side="bottom")

        # Linha 1: mensagem de status (feedback de operações)
        tk.Frame(rodape, bg=BORDER, height=1).pack(fill="x")
        self.var_barra = tk.StringVar(value="Pronto.")
        tk.Label(rodape, textvariable=self.var_barra, bg=CARD, fg=TEXT2,
                 font=("Segoe UI", 8), anchor="w",
                 padx=14, pady=4).pack(fill="x")

        # Linha 2: caminho do arquivo | contador do dia
        info = tk.Frame(rodape, bg="#12122a")
        info.pack(fill="x")
        tk.Label(info, text=str(EXCEL_PATH), bg="#12122a", fg="#3d3d5c",
                 font=("Segoe UI", 7), anchor="w",
                 padx=14, pady=4).pack(side="left")
        self.var_contador = tk.StringVar(value="")
        tk.Label(info, textvariable=self.var_contador, bg="#12122a", fg=ACCENT,
                 font=("Segoe UI", 8, "bold"), anchor="e",
                 padx=14, pady=4).pack(side="right")

    # ── Helpers de widget ────────────────────────────────────────────────────

    def _titulo(self, parent, texto):
        tk.Label(parent, text=texto, bg=BG, fg=TEXT,
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10, 0))

    def _sep(self, parent):
        tk.Frame(parent, bg=BORDER, height=1).pack(fill="x", pady=14)

    def _entry(self, parent, show="", var=None):
        return tk.Entry(
            parent, textvariable=var, show=show,
            bg=INPUT, fg=TEXT, insertbackground=TEXT,
            font=("Segoe UI", 10), bd=0, relief="flat",
            highlightthickness=1, highlightbackground=BORDER,
            highlightcolor=ACCENT,
        )

    def _btn(self, parent, texto, cmd, accent=False, padx=14, pady=7):
        bg_n  = ACCENT  if accent else SURFACE
        bg_h  = ACCENT2 if accent else BORDER
        bold  = "bold"  if accent else "normal"
        b = tk.Button(parent, text=texto, command=cmd,
                      bg=bg_n, fg=TEXT, activebackground=bg_h,
                      activeforeground=TEXT, font=("Segoe UI", 10, bold),
                      bd=0, relief="flat", cursor="hand2",
                      padx=padx, pady=pady)
        b.bind("<Enter>", lambda _e: b.config(bg=bg_h))
        b.bind("<Leave>", lambda _e: b.config(bg=bg_n))
        return b

    def _campo(self, grid, label, var, row, col):
        pad = (0, 10) if col == 0 else (10, 0)
        f = tk.Frame(grid, bg=BG)
        f.grid(row=row, column=col, sticky="ew", padx=pad, pady=4)
        tk.Label(f, text=label, bg=BG, fg=TEXT2,
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 2))
        self._entry(f, var=var).pack(fill="x")

    def _campo_combo(self, grid, label, var, valores, row, col):
        pad = (0, 10) if col == 0 else (10, 0)
        f = tk.Frame(grid, bg=BG)
        f.grid(row=row, column=col, sticky="ew", padx=pad, pady=4)
        tk.Label(f, text=label, bg=BG, fg=TEXT2,
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(0, 2))
        ttk.Combobox(f, textvariable=var, values=valores,
                     style="Dark.TCombobox", font=("Segoe UI", 10),
                     state="normal").pack(fill="x")

    # ── Lógica ───────────────────────────────────────────────────────────────

    def _atualizar_contador(self):
        n = contar_registros_hoje(str(EXCEL_PATH))
        s = "registro salvo hoje" if n == 1 else "registros salvos hoje"
        self.var_contador.set(f"{n} {s}")

    def _toggle_apikey(self):
        self.ent_apikey.config(
            show="" if self.ent_apikey.cget("show") == "•" else "•"
        )

    def _atualizar_aba(self):
        aba = _aba_para_status(self.var_status.get())
        self.lbl_aba.config(text=f'→ será salvo em: "{aba}"')

    def _preencher_campos(self, dados: dict):
        self.var_nome.set(dados.get("nome", ""))
        self.var_sobrenome.set(dados.get("sobrenome", ""))
        self.var_cidade.set(dados.get("cidade", ""))
        self.var_telefone.set(dados.get("telefone", ""))
        self.var_obs.set(dados.get("observacoes", ""))
        self.var_fonte.set(dados.get("fonte", "") or FONTES[0])
        self.var_status.set(dados.get("status", "") or STATUSES[0])

    def _processar(self):
        api_key = self.var_apikey.get().strip()
        texto   = self.txt_desc.get("1.0", "end").strip()

        if not api_key:
            messagebox.showwarning("API Key", "Informe a chave do Groq.\nObtenha gratuitamente em: console.groq.com")
            return
        if not texto:
            messagebox.showwarning("Descrição", "Descreva o contato antes de processar.")
            return

        try:
            config.salvar_api_key(ROOT, api_key)
        except Exception:
            pass  # falha ao persistir a chave não deve impedir o processamento
        self.btn_proc.config(state="disabled", text="Processando…")
        self.var_barra.set("⏳  Chamando a IA — aguarde...")

        def _worker():
            try:
                dados = extrair_dados(texto, api_key)
                self.after(0, self._preencher_campos, dados)
                self.after(0, self.var_barra.set,
                           "✓  Dados extraídos. Revise os campos e salve.")
            except Exception as exc:
                self.after(0, messagebox.showerror, "Erro na IA", str(exc))
                self.after(0, self.var_barra.set, "✗  Erro ao chamar a IA.")
            finally:
                self.after(0, lambda: self.btn_proc.config(
                    state="normal", text="Processar com IA"))

        threading.Thread(target=_worker, daemon=True).start()

    def _salvar(self):
        if not self.var_nome.get().strip():
            messagebox.showwarning("Campo obrigatório", "O campo Nome está vazio.")
            return

        dados = {
            "nome":        self.var_nome.get().strip(),
            "sobrenome":   self.var_sobrenome.get().strip(),
            "cidade":      self.var_cidade.get().strip(),
            "telefone":    self.var_telefone.get().strip(),
            "fonte":       self.var_fonte.get().strip(),
            "status":      self.var_status.get().strip(),
            "observacoes": self.var_obs.get().strip(),
            "aba":         _aba_para_status(self.var_status.get()),
        }

        try:
            adicionar_linha(str(EXCEL_PATH), dados)
            aba  = dados["aba"]
            nome = f'{dados["nome"]} {dados["sobrenome"]}'.strip()
            self.var_barra.set(f'✓  "{nome}" salvo em "{aba}".')
            self._atualizar_contador()
            self._limpar()
        except PermissionError as exc:
            messagebox.showerror("Arquivo bloqueado", str(exc))
        except Exception as exc:
            messagebox.showerror("Erro ao salvar", str(exc))

    def _abrir_excel(self):
        garantir_excel(str(EXCEL_PATH))
        try:
            os.startfile(str(EXCEL_PATH))
        except Exception as exc:
            messagebox.showerror("Erro", str(exc))

    def _importar_planilha(self):
        api_key = self.var_apikey.get().strip()
        if not api_key:
            messagebox.showwarning("API Key", "Informe a chave da API antes de importar.")
            return

        path = filedialog.askopenfilename(
            title="Selecionar planilha antiga",
            filetypes=[("Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
        )
        if not path:
            return

        self.btn_importar.config(state="disabled", text="Importando…")
        self.var_barra.set("⏳  Lendo planilha — aguarde...")

        def _worker():
            try:
                def progresso(msg):
                    self.after(0, self.var_barra.set, f"⏳  {msg}")

                importados, erros = importar_planilha(
                    path, str(EXCEL_PATH), api_key, progresso
                )

                msg = f"✓  {importados} registro(s) importado(s) com sucesso."
                if erros:
                    msg += f"  ({len(erros)} linha(s) ignorada(s))"
                self.after(0, self.var_barra.set, msg)
                self.after(0, self._atualizar_contador)

                if erros:
                    self.after(0, messagebox.showwarning,
                               "Importação concluída com avisos",
                               "\n".join(erros[:10]))
            except Exception as exc:
                self.after(0, messagebox.showerror, "Erro na importação", str(exc))
                self.after(0, self.var_barra.set, "✗  Erro ao importar planilha.")
            finally:
                self.after(0, lambda: self.btn_importar.config(
                    state="normal", text="Selecionar arquivo .xlsx e importar"))

        threading.Thread(target=_worker, daemon=True).start()

    def _limpar(self):
        self.txt_desc.delete("1.0", "end")
        for v in (self.var_nome, self.var_sobrenome,
                  self.var_cidade, self.var_telefone, self.var_obs):
            v.set("")
        self.var_fonte.set(FONTES[0])
        self.var_status.set(STATUSES[0])


if __name__ == "__main__":
    app = AnhangaRadar()
    app.mainloop()
