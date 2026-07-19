import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk

from excel import carregar_planilha, detectar_colunas
from matcher import carregar_configuracao, processar_matching
from relatorio import gerar_relatorio_excel
from normalizador import carregar_sinonimos

# Configuração do CustomTkinter
ctk.set_appearance_mode("Dark")  # Inicia no modo escuro elegante
ctk.set_default_color_theme("blue")

class AppPatrimonioMatch(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Patrimônio Match")
        self.geometry("960x700")
        self.minimum_size = (900, 650)
        
        self.caminho_gemat = tk.StringVar()
        self.caminho_sam = tk.StringVar()
        self.caminho_saida = tk.StringVar()
        
        # Carregar configurações iniciais
        self.config = carregar_configuracao("config.json")
        self.sinonimos = carregar_sinonimos("sinonimos.json")
        
        # Fila de comunicação da thread de matching com a interface principal
        self.fila_comunicacao = queue.Queue()
        
        self.montar_interface()
        
        # Iniciar monitoramento da fila de mensagens da thread secundária
        self.after(100, self.processar_fila)
        
    def montar_interface(self):
        # Grid Principal do App
        self.grid_rowconfigure(0, weight=0) # Header
        self.grid_rowconfigure(1, weight=1) # Content (Esquerda e Direita)
        self.grid_rowconfigure(2, weight=0) # Footer (Console e Ação)
        self.grid_columnconfigure(0, weight=1)
        
        # ================= HEADER =================
        self.frame_header = ctk.CTkFrame(self, corner_radius=0, fg_color=("#E2E8F0", "#1A202C"))
        self.frame_header.grid(row=0, column=0, sticky="nsew", padx=0, pady=0)
        
        self.lbl_titulo = ctk.CTkLabel(
            self.frame_header, 
            text="PATRIMÔNIO MATCH", 
            font=ctk.CTkFont(family="Segoe UI", size=24, weight="bold"),
            text_color=("#1A365D", "#90CDF4")
        )
        self.lbl_titulo.pack(pady=(15, 2))
        
        self.lbl_subtitulo = ctk.CTkLabel(
            self.frame_header, 
            text="Conciliação Automatizada e Saneamento Patrimonial URE (GEMAT x SAM)", 
            font=ctk.CTkFont(family="Segoe UI", size=12, slant="italic"),
            text_color=("#4A5568", "#A0AEC0")
        )
        self.lbl_subtitulo.pack(pady=(0, 15))
        
        # ================= CONTENT =================
        self.frame_content = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_content.grid(row=1, column=0, sticky="nsew", padx=20, pady=15)
        self.frame_content.grid_columnconfigure(0, weight=3) # Arquivos
        self.frame_content.grid_columnconfigure(1, weight=2) # Parâmetros
        self.frame_content.grid_rowconfigure(0, weight=1)
        
        # --- LADO ESQUERDO: Arquivos de Entrada e Saída ---
        self.frame_arquivos = ctk.CTkFrame(self.frame_content, corner_radius=10)
        self.frame_arquivos.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)
        
        self.lbl_sec_arquivos = ctk.CTkLabel(
            self.frame_arquivos, 
            text="Arquivos do Processamento", 
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")
        )
        self.lbl_sec_arquivos.pack(anchor="w", padx=20, pady=(15, 10))
        
        # Seleção GEMAT
        self.lbl_gemat = ctk.CTkLabel(self.frame_arquivos, text="Planilha GEMAT (.xlsx ou .xls):")
        self.lbl_gemat.pack(anchor="w", padx=20, pady=(10, 2))
        
        self.frame_gemat_input = ctk.CTkFrame(self.frame_arquivos, fg_color="transparent")
        self.frame_gemat_input.pack(fill="x", padx=20, pady=0)
        self.entry_gemat = ctk.CTkEntry(self.frame_gemat_input, textvariable=self.caminho_gemat, placeholder_text="Selecione o arquivo do GEMAT...")
        self.entry_gemat.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.btn_gemat = ctk.CTkButton(self.frame_gemat_input, text="Procurar...", width=90, command=self.buscar_gemat)
        self.btn_gemat.pack(side="right")
        
        # Seleção SAM
        self.lbl_sam = ctk.CTkLabel(self.frame_arquivos, text="Planilha SAM (.xlsx ou .xls):")
        self.lbl_sam.pack(anchor="w", padx=20, pady=(15, 2))
        
        self.frame_sam_input = ctk.CTkFrame(self.frame_arquivos, fg_color="transparent")
        self.frame_sam_input.pack(fill="x", padx=20, pady=0)
        self.entry_sam = ctk.CTkEntry(self.frame_sam_input, textvariable=self.caminho_sam, placeholder_text="Selecione o arquivo do SAM...")
        self.entry_sam.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.btn_sam = ctk.CTkButton(self.frame_sam_input, text="Procurar...", width=90, command=self.buscar_sam)
        self.btn_sam.pack(side="right")
        
        # Pasta de Saída
        self.lbl_saida = ctk.CTkLabel(self.frame_arquivos, text="Diretório de Destino do Relatório:")
        self.lbl_saida.pack(anchor="w", padx=20, pady=(15, 2))
        
        self.frame_saida_input = ctk.CTkFrame(self.frame_arquivos, fg_color="transparent")
        self.frame_saida_input.pack(fill="x", padx=20, pady=0)
        self.entry_saida = ctk.CTkEntry(self.frame_saida_input, textvariable=self.caminho_saida, placeholder_text="Selecione a pasta onde salvar o relatório...")
        self.entry_saida.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.btn_saida = ctk.CTkButton(self.frame_saida_input, text="Procurar...", width=90, command=self.buscar_saida)
        self.btn_saida.pack(side="right")
        
        # --- LADO DIREITO: Configurações do Match ---
        self.frame_params = ctk.CTkFrame(self.frame_content, corner_radius=10)
        self.frame_params.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=0)
        
        self.lbl_sec_params = ctk.CTkLabel(
            self.frame_params, 
            text="Parâmetros de Confiança", 
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold")
        )
        self.lbl_sec_params.pack(anchor="w", padx=20, pady=(15, 10))
        
        # Confiança Confirmada
        self.lbl_conf_conf = ctk.CTkLabel(self.frame_params, text="Min. Confirmada (%):")
        self.lbl_conf_conf.pack(anchor="w", padx=20, pady=(5, 2))
        self.spin_conf_conf = ctk.CTkSlider(self.frame_params, from_=80, to=100, number_of_steps=20, command=self.atualizar_lbl_conf)
        self.spin_conf_conf.set(self.config.get("confianca_confirmada", 95))
        self.spin_conf_conf.pack(fill="x", padx=20, pady=2)
        self.lbl_val_conf_conf = ctk.CTkLabel(self.frame_params, text=f"{int(self.spin_conf_conf.get())}%", font=ctk.CTkFont(weight="bold"))
        self.lbl_val_conf_conf.pack(anchor="e", padx=25)
        
        # Confiança Provável
        self.lbl_conf_prov = ctk.CTkLabel(self.frame_params, text="Min. Provável (%):")
        self.lbl_conf_prov.pack(anchor="w", padx=20, pady=(5, 2))
        self.spin_conf_prov = ctk.CTkSlider(self.frame_params, from_=60, to=85, number_of_steps=25, command=self.atualizar_lbl_conf)
        self.spin_conf_prov.set(self.config.get("confianca_provavel", 85))
        self.spin_conf_prov.pack(fill="x", padx=20, pady=2)
        self.lbl_val_conf_prov = ctk.CTkLabel(self.frame_params, text=f"{int(self.spin_conf_prov.get())}%", font=ctk.CTkFont(weight="bold"))
        self.lbl_val_conf_prov.pack(anchor="e", padx=25)

        # Regras de Comparação (Checkboxes)
        self.lbl_regras = ctk.CTkLabel(self.frame_params, text="Regras Ativas na Comparação:")
        self.lbl_regras.pack(anchor="w", padx=20, pady=(15, 2))
        
        self.chk_marca = ctk.CTkCheckBox(self.frame_params, text="Considerar Marca", hover=True)
        self.chk_marca.select() if self.config.get("usar_marca", True) else self.chk_marca.deselect()
        self.chk_marca.pack(anchor="w", padx=25, pady=4)
        
        self.chk_modelo = ctk.CTkCheckBox(self.frame_params, text="Considerar Modelo", hover=True)
        self.chk_modelo.select() if self.config.get("usar_modelo", True) else self.chk_modelo.deselect()
        self.chk_modelo.pack(anchor="w", padx=25, pady=4)
        
        self.chk_medidas = ctk.CTkCheckBox(self.frame_params, text="Considerar Medidas/Dimensões", hover=True)
        self.chk_medidas.select() if self.config.get("usar_medidas", True) else self.chk_medidas.deselect()
        self.chk_medidas.pack(anchor="w", padx=25, pady=4)
        
        # ================= FOOTER =================
        self.frame_footer = ctk.CTkFrame(self, corner_radius=10)
        self.frame_footer.grid(row=2, column=0, sticky="nsew", padx=20, pady=(5, 20))
        
        self.lbl_console = ctk.CTkLabel(
            self.frame_footer, 
            text="Progresso do Processamento:", 
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold")
        )
        self.lbl_console.pack(anchor="w", padx=20, pady=(10, 2))
        
        # Caixa de Texto para Console Log
        self.txt_console = ctk.CTkTextbox(self.frame_footer, height=140, font=ctk.CTkFont(family="Consolas", size=11))
        self.txt_console.pack(fill="x", padx=20, pady=5)
        self.txt_console.configure(state="disabled")
        
        # Barra de Progresso
        self.progress_bar = ctk.CTkProgressBar(self.frame_footer)
        self.progress_bar.pack(fill="x", padx=20, pady=5)
        self.progress_bar.set(0)
        
        # Painel inferior com botão e toggle de tema
        self.frame_botoes = ctk.CTkFrame(self.frame_footer, fg_color="transparent")
        self.frame_botoes.pack(fill="x", padx=20, pady=(5, 10))
        
        self.switch_tema = ctk.CTkSwitch(self.frame_botoes, text="Modo Escuro", command=self.alternar_tema)
        self.switch_tema.select()
        self.switch_tema.pack(side="left")
        
        self.btn_iniciar = ctk.CTkButton(
            self.frame_botoes, 
            text="Iniciar Processamento", 
            font=ctk.CTkFont(size=14, weight="bold"), 
            height=36,
            command=self.iniciar_matching
        )
        self.btn_iniciar.pack(side="right")
        
    def alternar_tema(self):
        if self.switch_tema.get() == 1:
            ctk.set_appearance_mode("Dark")
        else:
            ctk.set_appearance_mode("Light")
            
    def atualizar_lbl_conf(self, val):
        self.lbl_val_conf_conf.configure(text=f"{int(self.spin_conf_conf.get())}%")
        self.lbl_val_conf_prov.configure(text=f"{int(self.spin_conf_prov.get())}%")
        
    def buscar_gemat(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar Planilha GEMAT",
            filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
        )
        if caminho:
            self.caminho_gemat.set(caminho)
            
    def buscar_sam(self):
        caminho = filedialog.askopenfilename(
            title="Selecionar Planilha SAM",
            filetypes=[("Planilhas Excel", "*.xlsx *.xls")]
        )
        if caminho:
            self.caminho_sam.set(caminho)
            
    def buscar_saida(self):
        caminho = filedialog.askdirectory(title="Selecionar Pasta de Saída")
        if caminho:
            self.caminho_saida.set(caminho)
            
    def log(self, mensagem):
        self.txt_console.configure(state="normal")
        self.txt_console.insert("end", f"{mensagem}\n")
        self.txt_console.see("end")
        self.txt_console.configure(state="disabled")
        
    def iniciar_matching(self):
        gemat = self.caminho_gemat.get()
        sam = self.caminho_sam.get()
        saida = self.caminho_saida.get()
        
        if not gemat or not os.path.exists(gemat):
            messagebox.showerror("Erro de Validação", "Por favor, selecione uma planilha GEMAT válida.")
            return
        if not sam or not os.path.exists(sam):
            messagebox.showerror("Erro de Validação", "Por favor, selecione uma planilha SAM válida.")
            return
        if not saida or not os.path.exists(saida):
            messagebox.showerror("Erro de Validação", "Por favor, selecione um diretório de saída válido.")
            return
            
        # Bloquear botões
        self.btn_iniciar.configure(state="disabled")
        self.btn_gemat.configure(state="disabled")
        self.btn_sam.configure(state="disabled")
        self.btn_saida.configure(state="disabled")
        self.progress_bar.set(0)
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()
        
        # Limpar console
        self.txt_console.configure(state="normal")
        self.txt_console.delete("1.0", "end")
        self.txt_console.configure(state="disabled")
        
        # Atualizar configurações com os valores da interface
        self.config["confianca_confirmada"] = int(self.spin_conf_conf.get())
        self.config["confianca_provavel"] = int(self.spin_conf_prov.get())
        self.config["usar_marca"] = self.chk_marca.get() == 1
        self.config["usar_modelo"] = self.chk_modelo.get() == 1
        self.config["usar_medidas"] = self.chk_medidas.get() == 1
        
        # Iniciar thread do processamento secundário
        thread = threading.Thread(target=self.executar_processo_thread, args=(gemat, sam, saida))
        thread.daemon = True
        thread.start()
        
    def executar_processo_thread(self, gemat_path, sam_path, saida_dir):
        try:
            self.fila_comunicacao.put(("STATUS", "Iniciando conciliação patrimonial..."))
            
            # Carregar planilhas
            self.fila_comunicacao.put(("STATUS", "Carregando planilha GEMAT..."))
            df_gemat = carregar_planilha(gemat_path)
            colunas_gemat = detectar_colunas(df_gemat)
            
            self.fila_comunicacao.put(("STATUS", "Carregando planilha SAM..."))
            df_sam = carregar_planilha(sam_path)
            colunas_sam = detectar_colunas(df_sam)
            
            # Validação rápida de colunas obrigatórias
            if "patrimonio" not in colunas_gemat or "descricao" not in colunas_gemat:
                self.fila_comunicacao.put(("ERRO", "Não foi possível mapear automaticamente as colunas cruciais do GEMAT (Patrimônio e Descrição)."))
                return
            if "patrimonio" not in colunas_sam or "descricao" not in colunas_sam:
                self.fila_comunicacao.put(("ERRO", "Não foi possível mapear automaticamente as colunas cruciais do SAM (Patrimônio e Descrição)."))
                return
                
            self.fila_comunicacao.put(("STATUS", f"Colunas identificadas no GEMAT: {colunas_gemat}"))
            self.fila_comunicacao.put(("STATUS", f"Colunas identificadas no SAM: {colunas_sam}"))
            
            # Executar casamento
            def callback_progresso(msg):
                self.fila_comunicacao.put(("STATUS", msg))
                
            resultados = processar_matching(
                df_gemat, colunas_gemat,
                df_sam, colunas_sam,
                self.config, self.sinonimos,
                callback_progresso
            )
            
            # Salvar Relatório
            relatorio_path = os.path.join(saida_dir, "Relatorio_PatrimonioMatch.xlsx")
            self.fila_comunicacao.put(("STATUS", f"Gerando planilha formatada em: {relatorio_path}..."))
            gerar_relatorio_excel(resultados, relatorio_path)
            
            self.fila_comunicacao.put(("STATUS", "Processamento finalizado com sucesso!"))
            self.fila_comunicacao.put(("CONCLUIDO", relatorio_path))
            
        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            self.fila_comunicacao.put(("ERRO", f"Ocorreu um erro no processamento:\n{str(e)}\n\nDetalhes:\n{tb}"))
            
    def processar_fila(self):
        while not self.fila_comunicacao.empty():
            try:
                tipo, msg = self.fila_comunicacao.get_nowait()
                if tipo == "STATUS":
                    self.log(msg)
                elif tipo == "ERRO":
                    self.log(f"[ERRO CRÍTICO] {msg}")
                    self.finalizar_processo()
                    if "PermissionError" in msg or "aberto no Excel" in msg or "[Errno 13]" in msg:
                        messagebox.showwarning(
                            "Arquivo de Relatório Aberto", 
                            "O arquivo de relatório final está aberto no Excel ou em outro programa.\n\n"
                            "Por favor, FECHE O EXCEL (e a planilha do relatório) e clique em 'Iniciar Processamento' novamente."
                        )
                    else:
                        messagebox.showerror("Erro no Processamento", msg)
                elif tipo == "CONCLUIDO":
                    self.finalizar_processo()
                    messagebox.showinfo("Sucesso", f"O relatório de matching foi gerado com sucesso no arquivo:\n{msg}")
            except queue.Empty:
                break
        self.after(100, self.processar_fila)
        
    def finalizar_processo(self):
        self.progress_bar.stop()
        self.progress_bar.configure(mode="determinate")
        self.progress_bar.set(1.0)
        # Reativar botões
        self.btn_iniciar.configure(state="normal")
        self.btn_gemat.configure(state="normal")
        self.btn_sam.configure(state="normal")
        self.btn_saida.configure(state="normal")

if __name__ == "__main__":
    app = AppPatrimonioMatch()
    app.mainloop()
