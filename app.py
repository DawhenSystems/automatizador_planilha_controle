import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import NamedStyle
import traceback


# Função de log para exibir mensagens
def log(msg):
    print(msg)

# Função principal para processar os arquivos
def processar_arquivos(path_controle, path_ticklog, path_maxifrota, log_callback=None):
    """
    Lê os três arquivos, importa dados, aplica máximo por placa e salva cópia final.
    log_callback = função opcional para enviar mensagens para a UI.
    """
    def log(msg):
        if log_callback:
            log_callback(msg)  # Usa o callback fornecido para log
        else:
            print(msg)  # Caso contrário, usa o print padrão
    # LEITURA
    log("LENDO ARQUIVOS...")
    try:
        # Usando openpyxl para ler as planilhas com a formatação
        wb_controle = load_workbook(path_controle)

        # Verifica as abas disponíveis e printa no log
        for nome_aba in wb_controle.sheetnames:
            abas = wb_controle[nome_aba]
        log("ABAS ENCONTRADAS NA PLANILHA CONTROLE: " + ", ".join(wb_controle.sheetnames))
        
        wb_ticklog = load_workbook(path_ticklog)
        ws_ticklog = wb_ticklog.active  # Assume-se que seja a primeira aba
        
        wb_maxifrota = load_workbook(path_maxifrota)
        ws_maxifrota = wb_maxifrota.active  # Assume-se que seja a primeira aba da Maxi Frota
    except Exception as e:
        raise RuntimeError(f"ERRO AO LER PLANILHAS: {e}")

    # Cor para células não encontradas
    vermelho = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # ------------------- Passo 1: Mover Dados de H para G (KM final para KM inicial) -------------------
    log("MOVENDO DADOS DE KM/HR FINAL PARA KM/HR INICIAL...")

    # Função para verificar se a célula é mesclada
    def is_merged(cell):
        return cell.coordinate in ws_controle.merged_cells

    # Seleciona aba por aba na planilha Controle
    for aba in wb_controle.sheetnames:
        ws_controle = wb_controle[aba]

        # Copiar os dados da coluna H para a coluna G
        for row in range(7, ws_controle.max_row + 1):  # Começando da linha 7
            # Pula se a célula H estiver vazia
            if ws_controle[f"H{row}"].value is None:
                continue  # Pula se a célula H estiver vazia

            cell_h = ws_controle[f"H{row}"]
            cell_g = ws_controle[f"G{row}"]

            # Verifica se a célula H é mesclada
            if is_merged(cell_h):
                continue  # Se for mesclada, pula para a próxima linha

            # Copia o valor de H para G
            ws_controle[f"G{row}"] = cell_h.value
            ws_controle[f"H{row}"].value = None  # Limpa a coluna H

    log("DADOS DE KM/HR FINAL MOVIDOS COM SUCESSO PARA KM/HR INICIAL.")

    # ------------------- Passo 2: Processar Ticket Log -------------------

    log("PROCESSANDO PLACA DO TICKET LOG...")

    # Função para converter valores com vírgula para float
    def converter_para_float(valor):
        """Converte valores com vírgula (separador de milhar) para float"""
        if valor is None:
            return None
        try:
            # Se for string, remove a vírgula (separador de milhar)
            if isinstance(valor, str):
                valor = valor.replace(',', '').rstrip('0').rstrip('.')  # Remove a vírgula ao invés de substituir por ponto
            return float(valor)
        except (ValueError, TypeError):
            return None

    def converter_para_texto(valor):
        """Converte valores para texto, tratando None"""
        if valor is None:
            return 

    # Coletando placas e valores da planilha Ticket Log
    placas_ticklog = [ws_ticklog[f"F{row}"].value for row in range(2, ws_ticklog.max_row + 1)]  # Placa
    km_values = [converter_para_float(ws_ticklog[f"Q{row}"].value) for row in range(2, ws_ticklog.max_row + 1)]  # KM
    litros_values = [converter_para_float(ws_ticklog[f"O{row}"].value) for row in range(2, ws_ticklog.max_row + 1)]  # Litros
    valor_emissao_values = [converter_para_float(ws_ticklog[f"T{row}"].value) for row in range(2, ws_ticklog.max_row + 1)]  # Valor da emissão

    # ------------------- Passo 3: Preencher o Ticket Log -------------------
    for aba in wb_controle.sheetnames:
        ws_controle = wb_controle[aba]
        log(f"PROCESSANDO ABA: {aba}")

        # Coletando placas da planilha Controle (aba ativa)
        placas_controle = [ws_controle[f"A{row}"].value for row in range(7, ws_controle.max_row + 1)] 

        for i, placa in enumerate(placas_controle):
            if placa is None:
                continue  # Se não houver placa, pula para a próxima linha

            # Verifica se a operadora/plataforma é "TICKET LOG" na coluna E
            operadora_plataforma = ws_controle[f"E{i+7}"].value
            if operadora_plataforma != "TICKET LOG":
                continue  # Pula a linha se não for "TICKET LOG"

            # Tentar encontrar todas as ocorrências da placa na planilha Ticket Log
            if placa in placas_ticklog:
                # Obter todas as linhas onde a placa aparece
                indices = [index for index, value in enumerate(placas_ticklog) if value == placa]
                
                # Filtrar valores None
                km_validos = [km_values[idx] for idx in indices if km_values[idx] is not None]
                litros_validos = [litros_values[idx] for idx in indices if litros_values[idx] is not None]
                valor_emissao_validos = [valor_emissao_values[idx] for idx in indices if valor_emissao_values[idx] is not None]
                
                # Calcular apenas se houver valores válidos
                max_km = max(km_validos) if km_validos else None
                total_litros = sum(litros_validos) if litros_validos else 0
                total_valor_emissao = sum(valor_emissao_validos) if valor_emissao_validos else 0
                
                # Preencher as células nas colunas J, L e H
                ws_controle[f"J{i+7}"].value = total_litros
                ws_controle[f"L{i+7}"].value = total_valor_emissao
                ws_controle[f"H{i+7}"].value = str(max_km).rstrip('0').rstrip('.') if max_km is not None else None
                log(f"Placa {placa} atualizada: KM={max_km}, Litros={total_litros}, Valor Emissão={total_valor_emissao}")

            else:
                # Se não encontrar a placa, preencher a célula da coluna H com vermelho
                ws_controle[f"H{i+7}"].fill = vermelho
                log(f"Placa {placa} NÃO ENCONTRADA na planilha Ticket Log.")

                # -------------------- Escrever as placas não encontradas na coluna B --------------------
                ultima_linha = None
                for row in range(7, ws_controle.max_row + 1):
                    if ws_controle[f"B{row}"].value == "DEVOLUÇÃO:":
                        ultima_linha = row + 1
                        break
                
                if ultima_linha is not None:
                    # Encontrar próxima célula vazia
                    while ws_controle[f"B{ultima_linha}"].value:
                        ultima_linha += 1

                    ws_controle[f"B{ultima_linha}"].value = f"Placa {placa} não encontrada na planilha."
                    log(f"Placa {placa} registrada na coluna B como não encontrada.")

        log("PROCESSAMENTO DO TICKET LOG CONCLUÍDO.")

        # ------------------- Passo 4: Processar Maxi Frota -------------------

        log("PROCESSANDO PLACA DA MAXI FROTA...")

        # Coletando placas e valores da planilha Maxi Frota
        placas_maxifrota = [ws_maxifrota[f"E{row}"].value for row in range(2, ws_maxifrota.max_row + 1)]
        # Coletando placas e valores da planilha Maxi Frota
        placas_maxifrota = [ws_maxifrota[f"E{row}"].value for row in range(2, ws_maxifrota.max_row + 1)]
        hodrometro_values = [converter_para_float(ws_maxifrota[f"J{row}"].value) for row in range(2, ws_maxifrota.max_row + 1)]
        litros_maxifrota_values = [converter_para_float(ws_maxifrota[f"G{row}"].value) for row in range(2, ws_maxifrota.max_row + 1)]
        valor_emissao_maxifrota_values = [converter_para_float(ws_maxifrota[f"K{row}"].value) for row in range(2, ws_maxifrota.max_row + 1)]

        # ------------------- Passo 5: Preencher o Maxi Frota -------------------
        for i, placa in enumerate(placas_controle):
            if placa is None:
                continue

            # Verifica se a operadora/plataforma é "MAXI FROTA" na coluna E
            operadora_plataforma = ws_controle[f"E{i+7}"].value
            if operadora_plataforma != "MAXI FROTA":
                continue

            # Tentar encontrar todas as ocorrências da placa na planilha Maxi Frota
            if placa in placas_maxifrota:
                # Obter todas as linhas onde a placa aparece
                indices = [index for index, value in enumerate(placas_maxifrota) if value == placa]
                
                # Filtrar valores None
                hodrometro_validos = [hodrometro_values[idx] for idx in indices if hodrometro_values[idx] is not None]
                litros_validos = [litros_maxifrota_values[idx] for idx in indices if litros_maxifrota_values[idx] is not None]
                valor_emissao_validos = [valor_emissao_maxifrota_values[idx] for idx in indices if valor_emissao_maxifrota_values[idx] is not None]
                
                # Calcular apenas se houver valores válidos
                max_hodrometro = max(hodrometro_validos) if hodrometro_validos else None
                total_litros_maxifrota = sum(litros_validos) if litros_validos else 0
                total_valor_emissao_maxifrota = sum(valor_emissao_validos) if valor_emissao_validos else 0
                
                # Preencher as células nas colunas J, L e H
                ws_controle[f"J{i+7}"].value = total_litros_maxifrota
                ws_controle[f"L{i+7}"].value = total_valor_emissao_maxifrota
                
                ws_controle[f"H{i+7}"].value = str(max_hodrometro).replace('.', '') if max_hodrometro is not None else None
                # Compara a quantidade de dígitos do hodômetro inicial e final
                if max_hodrometro is not None:
                    hodometro_inicial = str(ws_controle[f"G{i+7}"].value).replace('.', '')
                    hodometro_final = str(max_hodrometro).replace('.', '')

                    # Se o hodômetro final tiver menos ou igual dígitos que o inicial, preencher com amarelo
                    if len(hodometro_final) + 1 < len(hodometro_inicial) or len(hodometro_final) > len(hodometro_inicial) + 1:
                        ws_controle[f"H{i+7}"].fill = amarelo
                    elif len(hodometro_final) + 1 == len(hodometro_inicial):
                        ws_controle[f"H{i+7}"].value = ws_controle[f"H{i+7}"].value + "0"  # Adiciona um zero ao final

                log(f"Placa {placa} atualizada: Hodômetro={max_hodrometro}, Litros={total_litros_maxifrota}, Valor Emissão={total_valor_emissao_maxifrota}")

            else:
                # Se não encontrar a placa, preencher a célula da coluna H com vermelho
                ws_controle[f"H{i+7}"].fill = vermelho
                log(f"Placa {placa} NÃO ENCONTRADA na planilha Maxi Frota.")

                # -------------------- Escrever as placas não encontradas na coluna E --------------------
                ultima_linha = None
                for row in range(7, ws_controle.max_row + 1):
                    if ws_controle[f"B{row}"].value == "DEVOLUÇÃO:":
                        ultima_linha = row + 1
                        break

                if ultima_linha is not None:
                    # Encontrar próxima célula vazia
                    while ws_controle[f"E{ultima_linha}"].value:
                        ultima_linha += 1

                    ws_controle[f"E{ultima_linha}"].value = f"Placa {placa} não encontrada na planilha."
                    log(f"Placa {placa} registrada na coluna E como não encontrada.")

        log("PROCESSAMENTO DA MAXI FROTA CONCLUÍDO.")

    # SALVANDO ARQUIVO MANTENDO A FORMATAÇÃO
    log("SALVANDO ARQUIVO COM AS ALTERAÇÕES...")
    try:
        wb_controle.save(path_controle)
        log(f"ARQUIVO SALVO COM SUCESSO: {path_controle}")
    except PermissionError:
        log("ERRO: NÃO FOI POSSÍVEL SALVAR O ARQUIVO. VERIFIQUE SE ELE ESTÁ ABERTO EM OUTRO PROGRAMA.")
    except Exception as e:
        log(f"ERRO AO SALVAR ARQUIVO FINAL: {e}")


# ----------------- INTERFACE TKINTER ----------------- #
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from openpyxl import load_workbook

class App:
    def __init__(self, root):
        self.root = root
        root.title("CONTROLE DE KM - IMPORTADOR TICKLOG")
        root.geometry("760x520")
        root.resizable(True, True)

        # FRAME TOP
        top = ttk.Frame(root, padding=10)
        top.pack(fill="x")

        # SELEÇÃO ARQUIVO CONTROLE
        ttk.Label(top, text="ARQUIVO CONTROLE (BASE):").grid(row=0, column=0, sticky="w")
        self.entry_controle = ttk.Entry(top, width=70)  # Definindo a variável entry_controle aqui
        self.entry_controle.grid(row=0, column=1, padx=5)
        ttk.Button(top, text="SELECIONAR", command=self.selecionar_controle).grid(row=0, column=2, padx=5)

        # SELEÇÃO ARQUIVO TICKLOG
        ttk.Label(top, text="ARQUIVO TICKLOG:").grid(row=1, column=0, sticky="w", pady=(8,0))
        self.entry_tick = ttk.Entry(top, width=70)  # Definindo a variável entry_tick aqui
        self.entry_tick.grid(row=1, column=1, padx=5, pady=(8,0))
        ttk.Button(top, text="SELECIONAR", command=self.selecionar_tick).grid(row=1, column=2, padx=5, pady=(8,0))

        # SELEÇÃO ARQUIVO MAXI FROTA
        ttk.Label(top, text="ARQUIVO MAXI FROTA:").grid(row=2, column=0, sticky="w", pady=(8,0))
        self.entry_maxifrota = ttk.Entry(top, width=70)  # Definindo a variável entry_maxifrota aqui
        self.entry_maxifrota.grid(row=2, column=1, padx=5, pady=(8,0))
        ttk.Button(top, text="SELECIONAR", command=self.selecionar_maxifrota).grid(row=2, column=2, padx=5, pady=(8,0))

        # BOTÕES DE AÇÃO
        actions = ttk.Frame(root, padding=10)
        actions.pack(fill="x")

        self.btn_process = ttk.Button(actions, text="IMPORTAR E ATUALIZAR (EXECUTAR)", command=self.executar_processamento)
        self.btn_process.grid(row=0, column=0, padx=5)

        self.btn_relatorio = ttk.Button(actions, text="VER RELATÓRIO", command=self.mostrar_relatorio, state="disabled")
        self.btn_relatorio.grid(row=0, column=1, padx=5)

        self.btn_abrir_pasta = ttk.Button(actions, text="ABRIR PASTA DO RESULTADO", command=self.abrir_pasta_saida, state="disabled")
        self.btn_abrir_pasta.grid(row=0, column=2, padx=5)

        # LOG
        log_frame = ttk.Frame(root, padding=10)
        log_frame.pack(fill="both", expand=True)

        ttk.Label(log_frame, text="LOG / MENSAGENS:").pack(anchor="w")
        self.txt_log = tk.Text(log_frame, height=18, wrap="word")
        self.txt_log.pack(fill="both", expand=True)

    # Função para selecionar o arquivo de controle
    def selecionar_controle(self):
        path = filedialog.askopenfilename(title="Selecione a planilha CONTROLE (Excel)", filetypes=[("Excel files","*.xlsx *.xls")])
        if path:
            self.entry_controle.delete(0, tk.END)
            self.entry_controle.insert(0, path)

    # Função para selecionar o arquivo Ticklog
    def selecionar_tick(self):
        path = filedialog.askopenfilename(title="Selecione o arquivo TICKLOG (Excel)", filetypes=[("Excel files","*.xlsx *.xls")])
        if path:
            self.entry_tick.delete(0, tk.END)
            self.entry_tick.insert(0, path)

    # Função para selecionar o arquivo Maxi Frota
    def selecionar_maxifrota(self):
        path = filedialog.askopenfilename(title="Selecione o arquivo MAXI FROTA (Excel)", filetypes=[("Excel files","*.xlsx *.xls")])
        if path:
            self.entry_maxifrota.delete(0, tk.END)
            self.entry_maxifrota.insert(0, path)

    # Função para logar as mensagens
    def log(self, msg):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.txt_log.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.txt_log.see(tk.END)
        self.txt_log.update_idletasks()

    # Função para executar o processamento
    def executar_processamento(self):
        path_controle = self.entry_controle.get().strip()
        path_tick = self.entry_tick.get().strip()
        path_maxifrota = self.entry_maxifrota.get().strip()

        if not path_controle or not os.path.exists(path_controle):
            messagebox.showerror("ERRO", "SELECIONE UM ARQUIVO DE CONTROLE VÁLIDO.")
            return
        if not path_tick or not os.path.exists(path_tick):
            messagebox.showerror("ERRO", "SELECIONE UM ARQUIVO TICKLOG VÁLIDO.")
            return
        if not path_maxifrota or not os.path.exists(path_maxifrota):
            messagebox.showerror("ERRO", "SELECIONE UM ARQUIVO MAXI FROTA VÁLIDO.")
            return

        self.txt_log.delete("1.0", tk.END)
        self.log("INICIANDO PROCESSAMENTO...")

        try:
            resultado = processar_arquivos(path_controle, path_tick, path_maxifrota, log_callback=self.log)
        except Exception as e:
            erro_completo = traceback.format_exc()
            print(erro_completo)  # Imprime no console
            messagebox.showerror("ERRO AO PROCESSAR", f"{str(e)}\n\nVeja o log para detalhes completos.")
            self.log("=" * 80)
            self.log("ERRO COMPLETO:")
            self.log(erro_completo)
            self.log("=" * 80)
            return

        self.result_info = resultado
        self.log("PROCESSAMENTO CONCLUÍDO COM SUCESSO.")
        self.btn_relatorio.config(state="normal")
        self.btn_abrir_pasta.config(state="normal")
        messagebox.showinfo("SUCESSO", "PROCESSAMENTO FINALIZADO. ARQUIVO SALVO:\n" + resultado["path_saida"])

    # Função para mostrar o relatório
    def mostrar_relatorio(self):
        if not self.result_info:
            messagebox.showinfo("RELATÓRIO", "NENHUM PROCESSAMENTO REALIZADO AINDA.")
            return
        rel = self.result_info["relatorio"]
        texto = (
            f"TOTAL DE REGISTROS NA BASE: {rel['total_controle']}\n"
            f"TOTAL DE REGISTROS NO TICKLOG: {rel['total_ticklog_registros']}\n"
            f"PLACAS ÚNICAS NA BASE: {rel['placas_controle_unicas']}\n"
            f"PLACAS ÚNICAS NO TICKLOG: {rel['placas_tick_unicas']}\n\n"
            f"DUPLICADOS NA BASE: {rel['duplicados_controle']}\n"
            f"DUPLICADOS NO TICKLOG: {rel['duplicados_tick']}\n\n"
            f"PLACAS DO TICKLOG NÃO ENCONTRADAS NA BASE (FALTANTES):\n{rel['placas_faltantes_na_base']}\n"
        )
        messagebox.showinfo("RELATÓRIO DE CONSISTÊNCIA", texto)

    # Função para abrir a pasta de saída
    def abrir_pasta_saida(self):
        if not self.result_info:
            return
        path_saida = self.result_info["path_saida"]
        pasta = os.path.dirname(path_saida) or "."
        try:
            if os.name == "nt":
                os.startfile(pasta)
            elif os.name == "posix":
                os.system(f'xdg-open "{pasta}"')
            else:
                messagebox.showinfo("PASTA DO RESULTADO", pasta)
        except Exception as e:
            messagebox.showerror("ERRO", f"NÃO FOI POSSÍVEL ABRIR A PASTA: {e}")

# ----------------- RODA A APLICAÇÃO ----------------- #
if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
