import tkinter as tk
from tkinter import messagebox
import pyautogui
import time
import sys
import os
import threading

# Importação opcional para manipulação de janelas (não obrigatória para este script funcionar)
try:
    import pygetwindow as gw
except ImportError:
    gw = None


class AppAgendasAvancada:
    def __init__(self, janela_principal):
        # Mudei de 'root' para 'janela_principal' para evitar conflito de nomes
        self.root = janela_principal
        self.root.title("Controlador de Impressão - Frente/Verso")
        self.root.geometry("460x650")

        # --- VARIÁVEIS DE ESTADO ---
        self.lotes_dados = []
        self.widgets_lotes = []
        self.indice_atual = 0
        self.etapa_atual = 'frente'
        self.tem_inicio_personalizado = tk.BooleanVar()

        # ========================================================
        # 1. ÁREA DE INPUTS (TOPO)
        # ========================================================
        frame_inputs = tk.Frame(self.root, pady=10)
        frame_inputs.pack(side='top', fill='x', padx=20)

        tk.Label(frame_inputs, text="Total de Páginas (PDF):").grid(row=0, column=0, sticky='w')
        self.entry_total = tk.Entry(frame_inputs, width=10, bg="#e0e0e0")
        self.entry_total.insert(0, "370")
        self.entry_total.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(frame_inputs, text="Folhas por Lote (Físicas):").grid(row=1, column=0, sticky='w')
        self.entry_folhas = tk.Entry(frame_inputs, width=10, bg="#e0e0e0")
        self.entry_folhas.insert(0, "10")
        self.entry_folhas.grid(row=1, column=1, padx=10, pady=5)

        tk.Checkbutton(frame_inputs, text="Início Personalizado?", variable=self.tem_inicio_personalizado,
                       command=self.toggle_inicio).grid(row=2, column=0, sticky='w')

        self.frame_inicio_vals = tk.Frame(frame_inputs)
        self.frame_inicio_vals.grid(row=2, column=1, sticky='w')
        tk.Label(self.frame_inicio_vals, text="1 a").pack(side='left')
        self.entry_fim_inicio = tk.Entry(self.frame_inicio_vals, width=5, bg="#e0e0e0")
        self.entry_fim_inicio.insert(0, "24")
        self.entry_fim_inicio.pack(side='left', padx=5)

        # Chama uma vez para garantir o estado correto inicial
        self.toggle_inicio()

        tk.Button(frame_inputs, text="GERAR TABELA", bg="#cccccc", command=self.calcular_lotes).grid(row=3, column=0,
                                                                                                     columnspan=2,
                                                                                                     pady=10,
                                                                                                     sticky='we')

        # ========================================================
        # 2. BOTÕES DE AÇÃO (RODAPÉ)
        # ========================================================
        frame_action = tk.Frame(self.root, pady=15, bg="#f0f0f0")
        frame_action.pack(side='bottom', fill='x')

        self.btn_reset = tk.Button(frame_action, text="RESETAR", bg="#ff9999", font=("Arial", 10, "bold"), height=2,
                                   width=15, command=self.resetar_tudo)
        self.btn_reset.pack(side='left', padx=40)

        self.btn_continuar = tk.Button(frame_action, text="INICIAR / CONTINUAR", bg="#99ff99",
                                       font=("Arial", 10, "bold"), height=2, width=20, command=self.acao_continuar)
        self.btn_continuar.pack(side='right', padx=40)

        # ========================================================
        # 3. ÁREA DE LISTAGEM (CENTRO)
        # ========================================================
        frame_header = tk.Frame(self.root)
        frame_header.pack(side='top', fill='x', padx=30, pady=(10, 0))
        tk.Label(frame_header, text="Lote", width=10, anchor='w', font=('Arial', 9, 'bold')).pack(side='left')
        tk.Label(frame_header, text="Intervalo", width=15, anchor='center', font=('Arial', 9, 'bold')).pack(side='left')
        tk.Label(frame_header, text="Frente", width=10, font=('Arial', 9, 'bold')).pack(side='left')
        tk.Label(frame_header, text="Verso", width=10, font=('Arial', 9, 'bold')).pack(side='left')

        self.canvas_scroll = tk.Canvas(self.root, borderwidth=1, relief="sunken", background="#ffffff")
        self.frame_lista = tk.Frame(self.canvas_scroll, background="#ffffff")
        self.vsb = tk.Scrollbar(self.root, orient="vertical", command=self.canvas_scroll.yview)
        self.canvas_scroll.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas_scroll.pack(side="top", fill="both", expand=True, padx=20, pady=5)
        self.canvas_scroll.create_window((4, 4), window=self.frame_lista, anchor="nw", tags="self.frame_lista")
        self.frame_lista.bind("<Configure>", self.on_frame_configure)

    # --- MÉTODOS AUXILIARES ---

    def on_frame_configure(self, _):
        """Atualiza a barra de rolagem quando a lista cresce."""
        self.canvas_scroll.configure(scrollregion=self.canvas_scroll.bbox("all"))

    def toggle_inicio(self):
        """Ativa/Desativa o campo de início personalizado."""
        if self.tem_inicio_personalizado.get():
            self.entry_fim_inicio.config(state='normal')
        else:
            self.entry_fim_inicio.config(state='disabled')

    @staticmethod
    def caminho_img(nome_arquivo):
        """
        Retorna o caminho correto da imagem, tanto rodando como script
        quanto rodando como executável (.exe).
        """
        if hasattr(sys, '_MEIPASS'):
            # Se estiver rodando como .exe, pega da pasta temporária
            return os.path.join(sys._MEIPASS, nome_arquivo)

        # Se estiver rodando como script normal, pega da pasta atual
        return os.path.join(os.getcwd(), nome_arquivo)

    # --- LÓGICA DE CÁLCULO DE PÁGINAS ---

    def calcular_lotes(self):
        # Limpa a tela anterior
        for widget in self.frame_lista.winfo_children():
            widget.destroy()
        self.lotes_dados = []
        self.widgets_lotes = []

        try:
            total = int(self.entry_total.get())
            limite_folhas = int(self.entry_folhas.get())
            passo = limite_folhas * 4  # 1 folha dobra = 4 páginas

            inicio = 1
            if self.tem_inicio_personalizado.get():
                fim_1 = int(self.entry_fim_inicio.get())
                self.lotes_dados.append(f"1-{fim_1}")
                inicio = fim_1 + 1

            while inicio <= total:
                fim = inicio + passo - 1
                if fim > total: fim = total
                self.lotes_dados.append(f"{inicio}-{fim}")
                inicio = fim + 1

            # Cria os itens na tela
            for i, intervalo in enumerate(self.lotes_dados):
                row_frame = tk.Frame(self.frame_lista, bg="white", pady=2)
                row_frame.pack(fill='x')
                tk.Label(row_frame, text=f"Miolo {i + 1:02d}", width=10, bg="white", anchor='w').pack(side='left')
                tk.Label(row_frame, text=intervalo, width=15, bg="#f0f0f0").pack(side='left', padx=5)

                cv_frente = tk.Canvas(row_frame, width=40, height=20, bg="red", highlightthickness=0)
                cv_frente.pack(side='left', padx=15)

                cv_verso = tk.Canvas(row_frame, width=40, height=20, bg="red", highlightthickness=0)
                cv_verso.pack(side='left', padx=15)

                self.widgets_lotes.append({'frente': cv_frente, 'verso': cv_verso})

            self.indice_atual = 0
            self.etapa_atual = 'frente'
            messagebox.showinfo("Pronto", "Tabela gerada! Clique em CONTINUAR para começar.")

        except ValueError:
            messagebox.showerror("Erro", "Verifique os números inseridos.")

    def resetar_tudo(self):
        if messagebox.askyesno("Resetar", "Deseja apagar o progresso?"):
            self.indice_atual = 0
            self.etapa_atual = 'frente'
            for widgets in self.widgets_lotes:
                widgets['frente'].config(bg='red')
                widgets['verso'].config(bg='red')

    # ========================================================
    # FUNÇÕES DE VISÃO COMPUTACIONAL (O CÉREBRO DO ROBÔ)
    # ========================================================

    def localizar_ancora(self, nome_imagem):
        caminho = self.caminho_img(nome_imagem)

        # --- DEBUG ---
        print(f"Tentando ler imagem em: {caminho}")
        # Isso vai aparecer no console do PyCharm. Verifique se o caminho faz sentido!

        if not os.path.exists(caminho):
            print("ERRO CRÍTICO: O arquivo de imagem NÃO existe nesse caminho.")
            return None

        try:
            # confidence=0.9 exige que você tenha instalado: pip install opencv-python
            posicao = pyautogui.locateOnScreen(caminho, confidence=0.9, grayscale=True)

            if posicao:
                # Calcula o centro manualmente para evitar erros de tipo
                cx = posicao.left + int(posicao.width / 2)
                cy = posicao.top + int(posicao.height / 2)
                return cx, cy

            return None

        except Exception as e:
            print(f"Erro ao buscar imagem (Verifique se instalou opencv): {e}")
            return None

    # ========================================================
    # EXECUÇÃO DA AUTOMAÇÃO
    # ========================================================

    def acao_continuar(self):
        if not self.lotes_dados:
            messagebox.showwarning("Aviso", "Gere a tabela primeiro.")
            return

        if self.indice_atual >= len(self.lotes_dados):
            messagebox.showinfo("Fim", "Todos os lotes foram impressos!")
            return

        lote_texto = self.lotes_dados[self.indice_atual]
        tipo_impressao = self.etapa_atual

        msg = f"LOTE: {lote_texto}\nLADO: {tipo_impressao.upper()}\n\nA impressora está pronta?"
        if messagebox.askyesno("Confirmar", msg):
            threading.Thread(target=self.rodar_robo, args=(lote_texto, tipo_impressao)).start()

    def rodar_robo(self, intervalo, tipo):
        try:
            self.btn_continuar.config(state='disabled')

            # Chama a função principal de cliques
            sucesso = self.template_cliques(intervalo, tipo)

            if not sucesso:
                messagebox.showerror("Erro",
                                     "Falha ao encontrar a janela do Adobe.\nVerifique se ela está aberta e visível.")
                return

            # Atualiza a interface gráfica (Cores)
            if tipo == 'frente':
                self.widgets_lotes[self.indice_atual]['frente'].config(bg='#00FF00')
                self.etapa_atual = 'verso'
                aviso = "Frente finalizada! Vire o papel e clique em Continuar."
            else:
                self.widgets_lotes[self.indice_atual]['verso'].config(bg='#00FF00')
                self.etapa_atual = 'frente'
                self.indice_atual += 1
                aviso = "Verso finalizado! Prepare o próximo lote."

            messagebox.showinfo("Sucesso", aviso)

        except Exception as e:
            print(f"Erro detalhado: {e}")
            messagebox.showerror("Erro Crítico", f"O robô parou:\n{str(e)}")
        finally:
            self.btn_continuar.config(state='normal')

    # ========================================================
    # [IMPORTANTE] ÁREA DE CONFIGURAÇÃO DOS CLIQUES
    # ========================================================
    def template_cliques(self, texto_paginas, modo):
        """
        Aqui definimos ONDE clicar.
        Usamos 'ref_x' e 'ref_y' como ponto zero (a âncora).
        """
        print(f"--- Iniciando {modo} do lote {texto_paginas} ---")

        # 1. Troca para a janela do Adobe
        pyautogui.hotkey('alt', 'tab')
        time.sleep(1.5)

        # 2. Abre o menu de impressão
        pyautogui.hotkey('ctrl', 'p')
        print("Aguardando janela de impressão abrir...")
        time.sleep(3.0)  # Tempo para a janela aparecer (aumente se o PC for lento)

        # 3. Localiza a Âncora (Ponto Zero)
        # Certifique-se de ter o arquivo 'ancora_paginas.png' na pasta!
        ponto_zero = self.localizar_ancora('ancora_paginas.png')

        if not ponto_zero:
            print("Não encontrei a âncora visual!")
            return False  # Falha

        # Desempacota as coordenadas da âncora
        ref_x, ref_y = ponto_zero
        print(f"Âncora localizada em X={ref_x}, Y={ref_y}. Calculando cliques relativos...")

        # --------------------------------------------------------------------------
        # CONFIGURAÇÃO DOS CLIQUES RELATIVOS
        # Use o Calibrador para descobrir os valores de OFFSET (Diferença)
        # Fórmula: pyautogui.click(ref_x + OFFSET_X, ref_y + OFFSET_Y)
        # --------------------------------------------------------------------------

        # --- CAMPO: PÁGINAS ---
        offset_paginas_x = 59
        offset_paginas_y = 0

        pyautogui.click(ref_x + offset_paginas_x, ref_y + offset_paginas_y)
        time.sleep(0.5)

        # --- CAMPO: NÚMERO DE PÁGINAS ---
        offset_numero_paginas_x = 206  # Exemplo: 150px pra direita
        offset_numero_paginas_y = -2  # Exemplo: Mesma altura

        pyautogui.click(ref_x + offset_numero_paginas_x, ref_y + offset_numero_paginas_y)
        time.sleep(0.5)

        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')
        pyautogui.write(texto_paginas)
        time.sleep(1)

        # --- CONFIGURAÇÕES ESPECÍFICAS DE FRENTE E VERSO ---
        if modo == 'frente':
            # 1. Botão "Mais Opções"
            off_mais_opcoes_x = -83
            off_mais_opcoes_y = 25
            pyautogui.click(ref_x + off_mais_opcoes_x, ref_y + off_mais_opcoes_y)
            pyautogui.click()
            time.sleep(1)

            # 2. Botão "Par ou Ímpar"
            off_par_impar_x = 150
            off_par_impar_y = 67
            pyautogui.click(ref_x + off_par_impar_x, ref_y + off_par_impar_y)
            time.sleep(1)

            # Seleciona Ímpar (Usando teclado é mais seguro que clique)
            pyautogui.press('down')
            pyautogui.press('enter')
            time.sleep(1)

            # 3. Botão "Múltiplo"
            off_multiplo_x = 150
            off_multiplo_y = 135
            pyautogui.click(ref_x + off_multiplo_x, ref_y + off_multiplo_y)
            time.sleep(1)

            # 4. Botão "Ordem de Páginas"
            off_ordem_x = 93
            off_ordem_y = 202
            pyautogui.click(ref_x + off_ordem_x, ref_y + off_ordem_y)
            time.sleep(1)

            # Seleciona Horizontal (Usando teclado é mais seguro que clique)
            pyautogui.press('up')
            pyautogui.press('enter')
            time.sleep(1)

        elif modo == 'verso':

            # 1. Botão "Mais Opções"
            off_mais_opcoes_x = -83
            off_mais_opcoes_y = 25
            pyautogui.click(ref_x + off_mais_opcoes_x, ref_y + off_mais_opcoes_y)
            pyautogui.click()
            time.sleep(1)

            # 2. Botão "Par ou Ímpar"
            off_par_impar_x = 150
            off_par_impar_y = 67
            pyautogui.click(ref_x + off_par_impar_x, ref_y + off_par_impar_y)
            time.sleep(1)

            # Seleciona Par (Geralmente 2 setas pra baixo)
            pyautogui.press('down', 2)
            pyautogui.press('enter')
            time.sleep(1)

            # 3. Botão "Múltiplo"
            off_multiplo_x = 150
            off_multiplo_y = 135
            pyautogui.click(ref_x + off_multiplo_x, ref_y + off_multiplo_y)
            time.sleep(1)

            # 4. Botão "Ordem de Páginas"
            off_ordem_x = 93
            off_ordem_y = 202
            pyautogui.click(ref_x + off_ordem_x, ref_y + off_ordem_y)
            time.sleep(1)

            # Seleciona Horizontal (Usando teclado é mais seguro que clique)
            pyautogui.press('down')
            pyautogui.press('enter')
            time.sleep(1)

        # --- FINALIZAR IMPRESSÃO ---
        time.sleep(1)

        pyautogui.press('enter')

        return True


if __name__ == "__main__":
    app = tk.Tk()
    interface = AppAgendasAvancada(app)
    app.mainloop()