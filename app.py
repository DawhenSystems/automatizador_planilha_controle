import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.styles import NamedStyle
import traceback
from access_validator import verificar_acesso_com_excecao


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

    def converter_para_texto(valor):
        """Converte valores para texto, tratando None"""
        if valor is None:
            return 

    # Coletando placas e valores da planilha Ticket Log
    placas_ticklog = [str(ws_ticklog[f"F{row}"].value).strip() if ws_ticklog[f"F{row}"].value else None for row in range(2, ws_ticklog.max_row + 1)]  # Placa
    km_ticklog = [(ws_ticklog[f"Q{row}"].value) for row in range(2, ws_ticklog.max_row + 1)]  # KM
    litros_ticklog = [(ws_ticklog[f"O{row}"].value) for row in range(2, ws_ticklog.max_row + 1)]  # Litros
    valor_emissao_ticklog = [(ws_ticklog[f"T{row}"].value) for row in range(2, ws_ticklog.max_row + 1)]  # Valor da emissão
    contrato_ticklog = [str(ws_ticklog[f"AB{row}"].value) if ws_ticklog[f"AB{row}"].value else None for row in range(2, ws_ticklog.max_row + 1)]  # Contrato

    # Agrupar valores por placa (soma por placa única)
    from collections import defaultdict
    valores_por_placa = defaultdict(float)
    
    for i, placa in enumerate(placas_ticklog):
        if placa and valor_emissao_ticklog[i]:
            valores_por_placa[placa] += valor_emissao_ticklog[i]
    
    # Soma total considerando apenas placas únicas
    valor_total_emissao_ticklog = sum(valores_por_placa.values())

    # Coletando placas e valores da planilha Maxi Frota
    placas_maxifrota = [str(ws_maxifrota[f"E{row}"].value).strip() if ws_maxifrota[f"E{row}"].value else None for row in range(2, ws_maxifrota.max_row + 1)]
    hodrometro_values = [(ws_maxifrota[f"P{row}"].value) for row in range(2, ws_maxifrota.max_row + 1)]
    litros_maxifrota_values = [(ws_maxifrota[f"Q{row}"].value) for row in range(2, ws_maxifrota.max_row + 1)]
    valor_emissao_maxifrota_values = [(ws_maxifrota[f"R{row}"].value) for row in range(2, ws_maxifrota.max_row + 1)]
    
    # Agrupar valores por placa (soma por placa única)
    valor_total_emissao_maxifrota = (sum(v for v in valor_emissao_maxifrota_values if isinstance(v, (int, float))))
    
    # Coletar dados do contrato de cada placa na Maxi Frota
    contrato_maxifrota = [str(ws_maxifrota[f"X{row}"].value).strip() if ws_maxifrota[f"X{row}"].value else None for row in range(2, ws_maxifrota.max_row + 1)]

    # Formatar dados para retirar texto desnecessário
    for i in range(len(placas_maxifrota)):
        placa = placas_maxifrota[i]
        if placa and "ALVES DA CUNHA " in placa:
            placas_maxifrota[i] = placa.split("ALVES DA CUNHA ", 1)[1]  # Mantém apenas a parte após "ALVES DA CUNHA "
        # Se não contiver "ALVES DA CUNHA ", mantém o valor original

    # ------------------- Passo 3: Preencher o Ticket Log -------------------
    for aba in wb_controle.sheetnames:
        ws_controle = wb_controle[aba]
        log(f"PROCESSANDO ABA: {aba}")

        # Inicia lista de placas alteradas de contrato para cada aba
        placas_alteradas_de_contrato = []

        # Coletando placas da planilha Controle (aba ativa)
        placas_controle = [str(ws_controle[f"A{row}"].value).strip() if ws_controle[f"A{row}"].value else None for row in range(7, ws_controle.max_row + 1)]

        # Dicionário para armazenar placas não encontradas por aba
        if not hasattr(processar_arquivos, 'placas_nao_encontradas_ticklog'):
            processar_arquivos.placas_nao_encontradas_ticklog = {}
        if not hasattr(processar_arquivos, 'placas_nao_encontradas_maxifrota'):
            processar_arquivos.placas_nao_encontradas_maxifrota = {}
        
        processar_arquivos.placas_nao_encontradas_ticklog[aba] = []
        processar_arquivos.placas_nao_encontradas_maxifrota[aba] = []

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
                km_validos = [km_ticklog[idx] for idx in indices if km_ticklog[idx] is not None]
                litros_validos = [litros_ticklog[idx] for idx in indices if litros_ticklog[idx] is not None]
                valor_emissao_validos = [valor_emissao_ticklog[idx] for idx in indices if valor_emissao_ticklog[idx] is not None]
                
                # Calcular apenas se houver valores válidos
                max_km = max(km_validos) if km_validos else None
                total_litros = sum(litros_validos) if litros_validos else 0
                total_valor_emissao = sum(valor_emissao_validos) if valor_emissao_validos else 0
                
                # Preencher as células nas colunas J, L e H
                ws_controle[f"J{i+7}"].value = total_litros
                ws_controle[f"L{i+7}"].value = total_valor_emissao
                ws_controle[f"H{i+7}"].value = str(max_km).rstrip('0').rstrip('.') if max_km is not None else None
                # log(f"Placa {placa} atualizada: KM={max_km}, Litros={total_litros}, Valor Emissão={total_valor_emissao}")

                # Verifica se o hodômetro inicial é zero ou "-"
                hodometro_inicial = str(ws_controle[f"G{i+7}"].value).strip() if ws_controle[f"G{i+7}"].value else ""
                if max_km is not None and hodometro_inicial in ["0", "-", ""]:
                    # Adiciona placa a uma lista (evitando duplicatas)
                    if placa not in placas_alteradas_de_contrato:
                        placas_alteradas_de_contrato.append(placa)

                    # Marca a célula com amarelo
                    ws_controle[f"H{i+7}"].fill = amarelo

            else:
                # Se não encontrar a placa, preencher a célula da coluna H com vermelho
                ws_controle[f"H{i+7}"].fill = vermelho
                log(f"  Placa {placa} NÃO ENCONTRADA na planilha Ticket Log.")

                # Apagar os valores das colunas J e L
                ws_controle[f"J{i+7}"].value = None
                ws_controle[f"L{i+7}"].value = None

                # Adicionar à lista de placas não encontradas (sem escrever ainda)
                processar_arquivos.placas_nao_encontradas_ticklog[aba].append(placa)

        # ------------------- Passo 4: Processar Maxi Frota -------------------

        log("PROCESSANDO PLACA DA MAXI FROTA...")

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
                
                # Verifica se o hodômetro inicial é zero ou "-"
                hodometro_inicial = str(ws_controle[f"G{i+7}"].value).strip() if ws_controle[f"G{i+7}"].value else ""
                if max_hodrometro is not None and hodometro_inicial in ["0", "-", ""]:
                    # Adiciona placa a uma lista (evitando duplicatas)
                    if placa not in placas_alteradas_de_contrato:
                        placas_alteradas_de_contrato.append(placa)
                    
                    # Marca a célula com amarelo
                    ws_controle[f"H{i+7}"].fill = amarelo

                # log(f"Placa {placa} atualizada: Hodômetro={max_hodrometro}, Litros={total_litros_maxifrota}, Valor Emissão={total_valor_emissao_maxifrota}")

            else:
                # Se não encontrar a placa, preencher a célula da coluna H com vermelho
                ws_controle[f"H{i+7}"].fill = vermelho
                log(f"  Placa {placa} NÃO ENCONTRADA na planilha Maxi Frota.")

                # Apagar os valores das colunas J e L
                ws_controle[f"J{i+7}"].value = None
                ws_controle[f"L{i+7}"].value = None

                # Adicionar à lista de placas não encontradas (sem escrever ainda)
                processar_arquivos.placas_nao_encontradas_maxifrota[aba].append(placa)

        # Registrar placas alteradas de contrato na coluna J APÓS processar todas as placas da aba
        if placas_alteradas_de_contrato:
            ultima_linha_contrato = None
            for row in range(7, ws_controle.max_row + 1):
                if ws_controle[f"J{row}"].value == "PLACAS ALTERADAS DE CONTRATO:":
                    ultima_linha_contrato = row + 1
                    break
            
            if ultima_linha_contrato is not None:
                for placa in placas_alteradas_de_contrato:
                    # Encontrar próxima célula vazia
                    while ws_controle[f"J{ultima_linha_contrato}"].value:
                        ultima_linha_contrato += 1

                    ws_controle[f"J{ultima_linha_contrato}"].value = f"Placa {placa} com possível alteração de contrato."
                    ultima_linha_contrato += 1

    # Agora escrever as placas não encontradas uma única vez
    for aba in wb_controle.sheetnames:
        ws_controle = wb_controle[aba]
        
        # Escrever placas não encontradas do Ticket Log
        if aba in processar_arquivos.placas_nao_encontradas_ticklog:
            placas_ticklog_aba = processar_arquivos.placas_nao_encontradas_ticklog[aba]
            if placas_ticklog_aba:
                ultima_linha = None
                for row in range(7, ws_controle.max_row + 1):
                    if ws_controle[f"B{row}"].value == "DEVOLUÇÃO:":
                        ultima_linha = row + 1
                        break
                
                if ultima_linha is not None:
                    for placa in placas_ticklog_aba:
                        while ws_controle[f"B{ultima_linha}"].value:
                            ultima_linha += 1
                        ws_controle[f"B{ultima_linha}"].value = f"Placa {placa} não encontrada na planilha."
                        log(f"  Placa {placa} registrada na coluna B como não encontrada.")
        
        # Escrever placas não encontradas da Maxi Frota
        if aba in processar_arquivos.placas_nao_encontradas_maxifrota:
            placas_maxifrota_aba = processar_arquivos.placas_nao_encontradas_maxifrota[aba]
            if placas_maxifrota_aba:
                ultima_linha = None
                for row in range(7, ws_controle.max_row + 1):
                    if ws_controle[f"B{row}"].value == "DEVOLUÇÃO:":
                        ultima_linha = row + 1
                        break
                
                if ultima_linha is not None:
                    for placa in placas_maxifrota_aba:
                        while ws_controle[f"E{ultima_linha}"].value:
                            ultima_linha += 1
                        ws_controle[f"E{ultima_linha}"].value = f"Placa {placa} não encontrada na planilha."
                        log(f"  Placa {placa} registrada na coluna E como não encontrada.")

    log(f"PROCESSAMENTO DO TICKET LOG CONCLUÍDO. VALOR TOTAL DE REGISTROS ÚNICOS PROCESSADOS: {len(set(placas_ticklog))} -> VALOR REAL TOTAL: {valor_total_emissao_ticklog:.2f}")
    log(f"PROCESSAMENTO DA MAXI FROTA CONCLUÍDO. VALOR TOTAL DE REGISTROS ÚNICOS PROCESSADOS: {len(set(placas_maxifrota))} -> VALOR REAL TOTAL: {valor_total_emissao_maxifrota:.2f}")

    # ------------------- Passo 6: Identificar placas novas -------------------
    log("IDENTIFICANDO PLACAS NOVAS NOS RELATÓRIOS...")

    # Dicionário de equivalências entre nomes de abas e possíveis variações nos relatórios
    equivalencias = {
        "BOLANDEIRA.PIRAJÁ": ["BOLANDEIRA", "DL SALVADOR"],
        "BONFIM": ["SENHOR DO BONFIM"],
        "CAERN-NATAL-RN": ["CAERN"],
        "COPASA ESGOTO CATAGUASES-MG": ["CATAGUASES"],
        "COPASA LAFAIETE-MG  - ÁGUA": ["COPASA AGUA", "COPASA ÁGUA"],
        "EMBASA - ALAGOINHAS-BA": ["ALAGOINHAS"],
        "EMBASA - FEIRA DE SANTANA-B": ["FEIRA DE SANTANA"],
        "EMBASA -SAJ COMERCIAL": ["SAJ COMERCIAL"],
        "EMBASA -SAJ MRR": ["SAJ MRR"],
        "IGUÁ SERGIPE": ["IGUÁ SERGIPE"],
        "ITAPARICA": ["ITAPARICA", "ILHA ITAPARICA"],
        "JAGUAQUARA": ["JAGUAQUARA"],
        "SEDE": ["SEDE", "FROTAS", "SEDE MATRIZ"]
    }

    def identificar_aba_por_contrato(contrato_texto, wb_controle):
        """
        Identifica a aba correta com base no texto do contrato usando equivalências.
        Retorna o nome da aba encontrada ou None.
        """
        if not contrato_texto:
            return None
        
        contrato_upper = contrato_texto.upper()
        
        # Percorrer cada aba e suas equivalências
        for nome_aba, variacoes in equivalencias.items():
            for variacao in variacoes:
                variacao_upper = variacao.upper()
                # Verifica se a variação está contida no contrato ou vice-versa
                if variacao_upper in contrato_upper or contrato_upper in variacao_upper:
                    # Verificar se a aba existe no workbook
                    for aba_real in wb_controle.sheetnames:
                        if nome_aba.upper() in aba_real.upper() or aba_real.upper() in nome_aba.upper():
                            return aba_real
        
        return None

    # Coletar todas as placas da planilha de controle (todas as abas)
    placas_controle_todas = set()
    for aba in wb_controle.sheetnames:
        ws_controle = wb_controle[aba]
        for row in range(7, ws_controle.max_row + 1):
            placa = str(ws_controle[f"A{row}"].value).strip() if ws_controle[f"A{row}"].value else None
            if placa and placa != "None":
                placas_controle_todas.add(placa)

    # Identificar placas novas do Ticket Log
    placas_ticklog_unicas = set(p for p in placas_ticklog if p and p != "None")
    placas_novas_ticklog = placas_ticklog_unicas - placas_controle_todas

    # Identificar placas novas da Maxi Frota
    placas_maxifrota_unicas = set(p for p in placas_maxifrota if p and p != "None")
    placas_novas_maxifrota = placas_maxifrota_unicas - placas_controle_todas

    # Processar placas novas do Ticket Log
    if placas_novas_ticklog:
        log(f"ENCONTRADAS {len(placas_novas_ticklog)} PLACA(S) NOVA(S) NO TICKET LOG")
        
        for placa_nova in sorted(placas_novas_ticklog):
            # Buscar dados da placa no Ticket Log
            indices = [i for i, p in enumerate(placas_ticklog) if p == placa_nova]
            
            # Obter contratos únicos para esta placa
            contratos = [contrato_ticklog[idx] for idx in indices if contrato_ticklog[idx] is not None]
            
            # Tentar identificar a aba usando cada contrato encontrado
            aba_encontrada = None
            contrato_usado = None
            for contrato in contratos:
                aba_encontrada = identificar_aba_por_contrato(contrato, wb_controle)
                if aba_encontrada:
                    contrato_usado = contrato
                    break
            
            if aba_encontrada:
                ws_aba = wb_controle[aba_encontrada]
                
                # Encontrar a célula "PLACAS NOVAS:" na coluna C da aba correspondente
                linha_ticklog = None
                for row in range(7, ws_aba.max_row + 1):
                    if ws_aba[f"C{row}"].value == "PLACAS NOVAS:":
                        linha_ticklog = row + 1
                        break
                
                if linha_ticklog:
                    # Encontrar próxima linha vazia
                    while linha_ticklog <= ws_aba.max_row and ws_aba[f"C{linha_ticklog}"].value:
                        linha_ticklog += 1
                    
                    # Calcular dados agregados da placa
                    km_validos = [km_ticklog[idx] for idx in indices if km_ticklog[idx] is not None]
                    litros_validos = [litros_ticklog[idx] for idx in indices if litros_ticklog[idx] is not None]
                    valor_validos = [valor_emissao_ticklog[idx] for idx in indices if valor_emissao_ticklog[idx] is not None]
                    
                    max_km = max(km_validos) if km_validos else 0
                    total_litros = sum(litros_validos) if litros_validos else 0
                    total_valor = sum(valor_validos) if valor_validos else 0
                    contratos_texto = ", ".join(set(contratos))
                    
                    ws_aba[f"C{linha_ticklog}"].value = (
                        f"{placa_nova} - {contratos_texto} - KM: {max_km} | Litros: {total_litros:.2f} | Valor: R$ {total_valor:.2f}"
                    )
                else:
                    log(f"  AVISO: Célula 'PLACAS NOVAS:' não encontrada na coluna C da aba '{aba_encontrada}' para placa {placa_nova}")
            else:
                # Se não encontrar aba, registrar na primeira aba na coluna C
                ws_primeira_aba = wb_controle[wb_controle.sheetnames[0]]
                
                linha_ticklog = None
                for row in range(7, ws_primeira_aba.max_row + 1):
                    if ws_primeira_aba[f"C{row}"].value == "PLACAS NOVAS:":
                        linha_ticklog = row + 1
                        break
                
                if linha_ticklog:
                    # Encontrar próxima linha vazia
                    while linha_ticklog <= ws_primeira_aba.max_row and ws_primeira_aba[f"C{linha_ticklog}"].value:
                        linha_ticklog += 1
                    
                    km_validos = [km_ticklog[idx] for idx in indices if km_ticklog[idx] is not None]
                    litros_validos = [litros_ticklog[idx] for idx in indices if litros_ticklog[idx] is not None]
                    valor_validos = [valor_emissao_ticklog[idx] for idx in indices if valor_emissao_ticklog[idx] is not None]
                    
                    max_km = max(km_validos) if km_validos else 0
                    total_litros = sum(litros_validos) if litros_validos else 0
                    total_valor = sum(valor_validos) if valor_validos else 0
                    contratos_texto = ", ".join(set(contratos))
                    
                    ws_primeira_aba[f"C{linha_ticklog}"].value = (
                        f"{placa_nova} - {contratos_texto} - KM: {max_km} | Litros: {total_litros:.2f} | Valor: R$ {total_valor:.2f}"
                    )
                    log(f"  Placa {placa_nova} registrada na coluna C da primeira aba (aba não encontrada para contrato: {contratos_texto})")

    # Processar placas novas da Maxi Frota
    if placas_novas_maxifrota:
        log(f"ENCONTRADAS {len(placas_novas_maxifrota)} PLACA(S) NOVA(S) NA MAXI FROTA")
        
        for placa_nova in sorted(placas_novas_maxifrota):
            # Buscar dados da placa na Maxi Frota
            indices = [i for i, p in enumerate(placas_maxifrota) if p == placa_nova]
            
            # Obter contratos únicos para esta placa
            contratos = [contrato_maxifrota[idx] for idx in indices if contrato_maxifrota[idx] is not None]
            
            # Tentar identificar a aba usando cada contrato encontrado
            aba_encontrada = None
            contrato_usado = None
            for contrato in contratos:
                aba_encontrada = identificar_aba_por_contrato(contrato, wb_controle)
                if aba_encontrada:
                    contrato_usado = contrato
                    break
            
            if aba_encontrada:
                ws_aba = wb_controle[aba_encontrada]
                
                # Encontrar a célula "PLACAS NOVAS:" na coluna F da aba correspondente
                linha_maxifrota = None
                for row in range(7, ws_aba.max_row + 1):
                    if ws_aba[f"F{row}"].value == "PLACAS NOVAS:":
                        linha_maxifrota = row + 1
                        break
                
                if linha_maxifrota:
                    # Encontrar próxima linha vazia
                    while linha_maxifrota <= ws_aba.max_row and ws_aba[f"F{linha_maxifrota}"].value:
                        linha_maxifrota += 1
                    
                    # Calcular dados agregados da placa
                    hodro_validos = [hodrometro_values[idx] for idx in indices if hodrometro_values[idx] is not None]
                    litros_validos = [litros_maxifrota_values[idx] for idx in indices if litros_maxifrota_values[idx] is not None]
                    valor_validos = [valor_emissao_maxifrota_values[idx] for idx in indices if valor_emissao_maxifrota_values[idx] is not None]
                    
                    max_hodro = max(hodro_validos) if hodro_validos else 0
                    total_litros = sum(litros_validos) if litros_validos else 0
                    total_valor = sum(valor_validos) if valor_validos else 0
                    contratos_texto = ", ".join(set(contratos))
                    
                    ws_aba[f"F{linha_maxifrota}"].value = (
                        f"{placa_nova} - {contratos_texto} - KM: {max_hodro} | Litros: {total_litros:.2f} | Valor: R$ {total_valor:.2f}"
                    )
                else:
                    log(f"  AVISO: Célula 'PLACAS NOVAS:' não encontrada na coluna F da aba '{aba_encontrada}' para placa {placa_nova}")
            else:
                # Se não encontrar aba, registrar na primeira aba na coluna F
                ws_primeira_aba = wb_controle[wb_controle.sheetnames[0]]
                
                linha_maxifrota = None
                for row in range(7, ws_primeira_aba.max_row + 1):
                    if ws_primeira_aba[f"F{row}"].value == "PLACAS NOVAS:":
                        linha_maxifrota = row + 1
                        break
                
                if linha_maxifrota:
                    # Encontrar próxima linha vazia
                    while linha_maxifrota <= ws_primeira_aba.max_row and ws_primeira_aba[f"F{linha_maxifrota}"].value:
                        linha_maxifrota += 1
                    
                    hodro_validos = [hodrometro_values[idx] for idx in indices if hodrometro_values[idx] is not None]
                    litros_validos = [litros_maxifrota_values[idx] for idx in indices if litros_maxifrota_values[idx] is not None]
                    valor_validos = [valor_emissao_maxifrota_values[idx] for idx in indices if valor_emissao_maxifrota_values[idx] is not None]
                    
                    max_hodro = max(hodro_validos) if hodro_validos else 0
                    total_litros = sum(litros_validos) if litros_validos else 0
                    total_valor = sum(valor_validos) if valor_validos else 0
                    contratos_texto = ", ".join(set(contratos))
                    
                    ws_primeira_aba[f"F{linha_maxifrota}"].value = (
                        f"{placa_nova} - {contratos_texto} - KM: {max_hodro} | Litros: {total_litros:.2f} | Valor: R$ {total_valor:.2f}"
                    )

    if not placas_novas_ticklog and not placas_novas_maxifrota:
        log("  Nenhuma placa nova identificada nos relatórios.")

    # ------------------- Passo Final: Salvar o arquivo final -------------------
    # SALVANDO ARQUIVO MANTENDO A FORMATAÇÃO
    log("SALVANDO ARQUIVO COM SUCESSO...")
    try:
        wb_controle.save(path_controle)
        log(f"ARQUIVO SALVO COM SUCESSO: {path_controle}")
        
        # Retornar informações do processamento
        return {
            "path_saida": path_controle,
            "relatorio": {
                "total_controle": len(placas_controle_todas),
                "total_ticklog_registros": len(placas_ticklog),
                "placas_controle_unicas": len(placas_controle_todas),
                "placas_tick_unicas": len(set(placas_ticklog)),
                "duplicados_controle": 0,  # Pode calcular se necessário
                "duplicados_tick": len(placas_ticklog) - len(set(placas_ticklog)),
                "placas_faltantes_na_base": f"{len(placas_novas_ticklog)} do Ticket Log, {len(placas_novas_maxifrota)} da Maxi Frota"
            }
        }
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
    # Valida o accesso antes de iniciar a aplicação
    verificar_acesso_com_excecao()
    
    root = tk.Tk()
    app = App(root)
    root.mainloop()
