# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from datetime import datetime
import traceback
from collections import defaultdict

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from access_validator import verificar_acesso_com_excecao


# =========================
# VALIDACAO DA SEMANA (LINHA 4)
# =========================
def localizar_coluna_semana(ws, semana_num: int, header_row: int = 3):
    """
    PROCURA "SEMANA N" NA LINHA header_row.
    ACEITA TEXTO DO TIPO: "SEMANA 1 - 02.01 - 08.01"
    """
    alvo = f"SEMANA {semana_num}"
    for col in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=col).value
        if v is None:
            continue
        v_str = str(v).strip().upper()
        if alvo in v_str:
            return col
    return None


def validar_semana_para_preencher(ws, semana_num: int, header_row: int = 4):
    """
    1) ACHAR A SEMANA NA LINHA 4
    2) DESCER 1 LINHA (MESMA COLUNA) E VERIFICAR "VALOR"
        - SE NAO -> ERRO "NAO FOI ENCONTRADO 'VALOR' LOGO ABAIXO..."
    3) DESCER MAIS 1 LINHA E VERIFICAR SE TEM ALGO PREENCHIDO
        - SE SIM -> DESCER MAIS 1 LINHA E VERIFICAR NOVAMENTE
            - SE SIM -> ERRO "SEMANA SELECIONADA JA PREENCHIDA..."
            - SE NAO -> BLOQUEIA POR SEGURANCA (INDICIO DE PREENCHIMENTO)
        - SE NAO -> OK PARA PREENCHER
    """
    col = localizar_coluna_semana(ws, semana_num, header_row=header_row)
    if col is None:
        return {"ok": False, "msg": f'NÃO FOI ENCONTRADA "SEMANA {semana_num}" NA LINHA {header_row}.', "col": None}

    v_valor = ws.cell(row=header_row + 1, column=col).value
    if v_valor is None or str(v_valor).strip().upper() != "VALOR":
        letra = get_column_letter(col)
        return {
            "ok": False,
            "msg": f'NÃO FOI ENCONTRADO "VALOR" LOGO ABAIXO DA SEMANA {semana_num} (COLUNA {letra}). ALTERAÇÃO DE LAYOUT.',
            "col": col
        }

    v_l1 = ws.cell(row=header_row + 2, column=col).value
    if v_l1 is not None and str(v_l1).strip() != "":
        v_l2 = ws.cell(row=header_row + 3, column=col).value
        if v_l2 is not None and str(v_l2).strip() != "":
            return {
                "ok": False,
                "msg": "SEMANA SELECIONADA JÁ PREENCHIDA. SELECIONE OUTRA SEMANA OU VERIFIQUE A PLANILHA BASE.",
                "col": col
            }
        else:
            return {
                "ok": False,
                "msg": "FOI IDENTIFICADO CONTEÚDO NA SEMANA SELECIONADA. SELECIONE OUTRA SEMANA OU VERIFIQUE A PLANILHA BASE.",
                "col": col
            }

    return {"ok": True, "msg": "OK", "col": col}


def aplicar_validacao_semana_em_todas_abas(
    wb_controle,
    semana_num: int,
    log_fn=print,
    header_row: int = 3
):
    abas_validadas = 0

    for aba in wb_controle.sheetnames:
        ws = wb_controle[aba]

        r = validar_semana_para_preencher(ws, semana_num, header_row=header_row)

        # 👉 CASO NÃO ENCONTRE A SEMANA
        if not r["ok"] and "NÃO FOI ENCONTRADA \"SEMANA" in r["msg"]:
            if abas_validadas >= 2:
                log_fn(
                    f"[{aba}] SEMANA {semana_num} NÃO ENCONTRADA. "
                    "FINAL DO BLOCO DE ABAS SEMANAIS. SEGUINDO PROCESSAMENTO."
                )
                return True, "FIM DO BLOCO DE SEMANAS"
            else:
                log_fn(f"[{aba}] {r['msg']}")
                return False, f"[{aba}] {r['msg']}"

        # 👉 OUTROS ERROS (VALOR, SEMANA JÁ PREENCHIDA, ETC.)
        if not r["ok"]:
            log_fn(f"[{aba}] {r['msg']}")
            return False, f"[{aba}] {r['msg']}"

        # 👉 OK
        abas_validadas += 1
        log_fn(f"[{aba}] SEMANA {semana_num} OK PARA PREENCHER.")

    return True, "OK"

# =========================
# MAPEAMENTO DE SEMANAS PARA COLUNAS
# =========================
def obter_colunas_da_semana(semana_num):
    """
    Retorna as colunas correspondentes à semana selecionada.
    
    Args:
        semana_num (int): Número da semana (1-5)
    
    Returns:
        dict: Dicionário com as colunas para VALOR e QUANTIDADE (L)
    """
    mapeamento_semanas = {
        1: {"VALOR": "E", "QUANTIDADE (L)": "F"},
        2: {"VALOR": "H", "QUANTIDADE (L)": "I"},
        3: {"VALOR": "K", "QUANTIDADE (L)": "M"},
        4: {"VALOR": "N", "QUANTIDADE (L)": "O"},
        5: {"VALOR": "Q", "QUANTIDADE (L)": "R"}
    }
    
    if semana_num not in mapeamento_semanas:
        raise ValueError(f"SEMANA {semana_num} INVÁLIDA. SELECIONE ENTRE 1 E 5.")
    
    return mapeamento_semanas[semana_num]

# =========================
# PROCESSAMENTO PRINCIPAL
# =========================
def processar_arquivos(path_controle, path_ticketlog, path_maxifrota, semana_num, header_row_semana=4, log_callback=None):
    """
    LÊ OS TRÊS ARQUIVOS, VALIDA SEMANA NA PLANILHA CONTROLE E PROCESSA OS DADOS.
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("LENDO ARQUIVOS...")

    try:
        wb_controle = load_workbook(path_controle)
        log("ABAS ENCONTRADAS NA PLANILHA CONTROLE: " + ", ".join(wb_controle.sheetnames))

        wb_ticketlog = load_workbook(path_ticketlog)
        ws_ticketlog = wb_ticketlog.active

        wb_maxifrota = load_workbook(path_maxifrota)
        ws_maxifrota = wb_maxifrota.active
    except Exception as e:
        raise RuntimeError(f"ERRO AO LER PLANILHAS: {e}")

    # =========================
    # PASSO 1: VALIDAR SEMANA
    # =========================
    log("VERIFICANDO SE A SEMANA SELECIONADA JÁ FOI PREENCHIDA...")
    ok_semana, msg_semana = aplicar_validacao_semana_em_todas_abas(
        wb_controle, semana_num, log_fn=log, header_row=header_row_semana
    )
    if not ok_semana:
        raise RuntimeError(msg_semana)

    # =========================
    # ESTILOS
    # =========================
    vermelho = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    amarelo = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # =========================
    # VARIÁVEL PARA REGISTRAR COLUNAS DA SEMANA SELECIONADA
    # =========================
    try:
        colunas_da_semana = obter_colunas_da_semana(semana_num)
        log(f"COLUNAS DA SEMANA {semana_num}: {colunas_da_semana}")
    except ValueError as ve:
        raise RuntimeError(str(ve))

    def converter_para_texto(valor):
        if valor is None:
            return ""
        return str(valor)

    # =========================
    # PASSO 2: PROCESSAR TICKETLOG
    # =========================
    log("PROCESSANDO PLACA DO TICKET LOG...")

    placas_ticketlog = [
        str(ws_ticketlog[f"F{row}"].value).strip() if ws_ticketlog[f"F{row}"].value else None
        for row in range(2, ws_ticketlog.max_row + 1)
    ]
    litros_ticketlog = [(ws_ticketlog[f"O{row}"].value) for row in range(2, ws_ticketlog.max_row + 1)]
    valor_emissao_ticketlog = [(ws_ticketlog[f"T{row}"].value) for row in range(2, ws_ticketlog.max_row + 1)]
    contrato_ticketlog = [
        str(ws_ticketlog[f"AB{row}"].value) if ws_ticketlog[f"AB{row}"].value else None
        for row in range(2, ws_ticketlog.max_row + 1)
    ]

    valores_por_placa = defaultdict(float)
    for i, placa in enumerate(placas_ticketlog):
        if placa and isinstance(valor_emissao_ticketlog[i], (int, float)):
            valores_por_placa[placa] += valor_emissao_ticketlog[i]

    valor_total_emissao_ticketlog = sum(valores_por_placa.values())

    # =========================
    # PASSO 3: PROCESSAR MAXI FROTA
    # =========================
    placas_maxifrota = [
        str(ws_maxifrota[f"E{row}"].value).strip() if ws_maxifrota[f"E{row}"].value else None
        for row in range(2, ws_maxifrota.max_row + 1)
    ]
    litros_maxifrota_values = [(ws_maxifrota[f"AC{row}"].value) for row in range(2, ws_maxifrota.max_row + 1)]
    valor_emissao_maxifrota_values = [(ws_maxifrota[f"Q{row}"].value) for row in range(2, ws_maxifrota.max_row + 1)]

    valor_total_emissao_maxifrota = sum(
        v for v in valor_emissao_maxifrota_values if isinstance(v, (int, float))
    )

    contrato_maxifrota = [
        str(ws_maxifrota[f"AE{row}"].value).strip() if ws_maxifrota[f"AE{row}"].value else None
        for row in range(2, ws_maxifrota.max_row + 1)
    ]

    for i in range(len(placas_maxifrota)):
        placa = placas_maxifrota[i]
        if placa and "ALVES DA CUNHA " in placa:
            placas_maxifrota[i] = placa.split("ALVES DA CUNHA ", 1)[1]

    # =========================
    # PASSO 4/5: PREENCHER NA PLANILHA CONTROLE 
    # =========================
    for aba in wb_controle.sheetnames:
        ws_controle = wb_controle[aba]
        log(f"PROCESSANDO ABA: {aba}")

        placas_alteradas_de_contrato = []

        placas_controle = [
            str(ws_controle[f"A{row}"].value).strip() if ws_controle[f"A{row}"].value else None
            for row in range(7, ws_controle.max_row + 1)
        ]

        if not hasattr(processar_arquivos, "placas_nao_encontradas_ticketlog"):
            processar_arquivos.placas_nao_encontradas_ticketlog = {}
        if not hasattr(processar_arquivos, "placas_nao_encontradas_maxifrota"):
            processar_arquivos.placas_nao_encontradas_maxifrota = {}

        processar_arquivos.placas_nao_encontradas_ticketlog[aba] = []
        processar_arquivos.placas_nao_encontradas_maxifrota[aba] = []

        # -------- TICKET LOG --------
        for i, placa in enumerate(placas_controle):
            if placa is None:
                continue

            if placa in placas_ticketlog:
                indices = [index for index, value in enumerate(placas_ticketlog) if value == placa]

                valor_validos = [valor_emissao_ticketlog[idx] for idx in indices if valor_emissao_ticketlog[idx] is not None]
                litros_validos = [litros_ticketlog[idx] for idx in indices if litros_ticketlog[idx] is not None]

                total_valor = sum(valor_validos) if valor_validos else 0
                total_litros = sum(litros_validos) if litros_validos else 0

                # USAR AS COLUNAS DA SEMANA SELECIONADA
                row_num = i + 5 # Para começar na linha 5
                ws_controle[f"{colunas_da_semana['QUANTIDADE (L)']}{row_num}"].value = total_litros
                ws_controle[f"{colunas_da_semana['VALOR']}{row_num}"].value = total_valor

                if total_litros == 0 or total_litros is None:
                    if placa not in placas_alteradas_de_contrato:
                        placas_alteradas_de_contrato.append(placa)
                    ws_controle[f"{colunas_da_semana['QUANTIDADE (L)']}{row_num}"].fill = amarelo
                    log(f"  PLACA {placa}: LITROS ZERADO OU NULO")

                if total_valor == 0 or total_valor is None:
                    if placa not in placas_alteradas_de_contrato:
                        placas_alteradas_de_contrato.append(placa)
                    ws_controle[f"{colunas_da_semana['VALOR']}{row_num}"].fill = amarelo
                    log(f"  PLACA {placa}: VALOR ZERADO OU NULO")
            else:
                row_num = i + 5 # Para começar na linha 5
                ws_controle[f"{colunas_da_semana['QUANTIDADE (L)']}{row_num}"].fill = vermelho
                log(f"  PLACA {placa} NÃO ENCONTRADA NA PLANILHA TICKET LOG.")
                ws_controle[f"{colunas_da_semana['VALOR']}{row_num}"].value = None
                ws_controle[f"{colunas_da_semana['QUANTIDADE (L)']}{row_num}"].value = None
                processar_arquivos.placas_nao_encontradas_ticketlog[aba].append(placa)

        # -------- MAXI FROTA --------
        log("PROCESSANDO PLACA DA MAXI FROTA...")

        for i, placa in enumerate(placas_controle):
            if placa is None:
                continue

            if placa in placas_maxifrota:
                indices = [index for index, value in enumerate(placas_maxifrota) if value == placa]

                valor_validos = [valor_emissao_maxifrota_values[idx] for idx in indices if valor_emissao_maxifrota_values[idx] is not None]
                litros_validos = [litros_maxifrota_values[idx] for idx in indices if litros_maxifrota_values[idx] is not None]

                total_valor = sum(valor_validos) if valor_validos else 0
                total_litros = sum(litros_validos) if litros_validos else 0

                row_num = i + 5 # Para começar na linha 5
                ws_controle[f"{colunas_da_semana['QUANTIDADE (L)']}{row_num}"].value = total_litros
                ws_controle[f"{colunas_da_semana['VALOR']}{row_num}"].value = total_valor

                if total_litros == 0 or total_litros is None:
                    if placa not in placas_alteradas_de_contrato:
                        placas_alteradas_de_contrato.append(placa)
                    ws_controle[f"{colunas_da_semana['QUANTIDADE (L)']}{row_num}"].fill = amarelo
                    log(f"  PLACA {placa}: LITROS ZERADO OU NULO")
                
                if total_valor == 0 or total_valor is None:
                    if placa not in placas_alteradas_de_contrato:
                        placas_alteradas_de_contrato.append(placa)
                    ws_controle[f"{colunas_da_semana['VALOR']}{row_num}"].fill = amarelo
                    log(f"  PLACA {placa}: VALOR ZERADO OU NULO")
            else:
                row_num = i + 5 # Para começar na linha 5
                ws_controle[f"{colunas_da_semana['QUANTIDADE (L)']}{row_num}"].fill = vermelho
                log(f"  PLACA {placa} NÃO ENCONTRADA NA PLANILHA MAXI FROTA.")
                ws_controle[f"{colunas_da_semana['VALOR']}{row_num}"].value = None
                ws_controle[f"{colunas_da_semana['QUANTIDADE (L)']}{row_num}"].value = None
                processar_arquivos.placas_nao_encontradas_maxifrota[aba].append(placa)

        # REGISTRAR PLACAS ALTERADAS
        if placas_alteradas_de_contrato:
            # Seleciona aba DIVERGÊNCIAS
            ws_divergencias = wb_controle['DIVERGÊNCIAS']

            if ws_divergencias[f"A{row}"].value == "PLACAS ALTERADAS DE CONTRATO:":
                ultima_linha_contrato = row + 1
                break

            if ultima_linha_contrato is not None:
                for placa in placas_alteradas_de_contrato:
                    while ws_divergencias[f"J{ultima_linha_contrato}"].value:
                        ultima_linha_contrato += 1
                    ws_divergencias[f"J{ultima_linha_contrato}"].value = f"PLACA {placa} COM POSSÍVEL ALTERAÇÃO DE CONTRATO."
                    ultima_linha_contrato += 1

    # ESCREVER NA DEVOLUCAO
    for aba in wb_controle.sheetnames:
        # ESCREVER NA DEVOLUCAO
        ws_divergencias_devolucao = wb_controle['DIVERGÊNCIAS']
        
        # -------- PROCESSAR PLACAS NÃO ENCONTRADAS DO TICKET LOG --------
        for aba in wb_controle.sheetnames:
            if aba in processar_arquivos.placas_nao_encontradas_ticketlog:
                placas_ticketlog_aba = processar_arquivos.placas_nao_encontradas_ticketlog[aba]
                if placas_ticketlog_aba:
                    # ENCONTRAR A LINHA DO CABEÇALHO "DEVOLUÇÃO:" NA COLUNA H
                    ultima_linha = None
                    for row in range(1, ws_divergencias_devolucao.max_row + 1):
                        if ws_divergencias_devolucao[f"H{row}"].value == "DEVOLUÇÃO:":
                            ultima_linha = row + 1  # COMEÇAR NA LINHA ABAIXO DO CABEÇALHO
                            break
                    
                    if ultima_linha is not None:
                        for placa in placas_ticketlog_aba:
                            # PULAR LINHAS JÁ PREENCHIDAS
                            while ws_divergencias_devolucao[f"H{ultima_linha}"].value:
                                ultima_linha += 1
                            ws_divergencias_devolucao[f"H{ultima_linha}"].value = f"PLACA {placa} NÃO ENCONTRADA NA PLANILHA TICKET LOG."
                            log(f"  PLACA {placa} ({aba}) REGISTRADA NA COLUNA H COMO NÃO ENCONTRADA.")
                            ultima_linha += 1
                    else:
                        log(f"  AVISO: CABEÇALHO 'DEVOLUÇÃO:' NÃO ENCONTRADO NA COLUNA H DA ABA DIVERGÊNCIAS.")

        # -------- PROCESSAR PLACAS NÃO ENCONTRADAS DA MAXI FROTA --------
        for aba in wb_controle.sheetnames:
            if aba in processar_arquivos.placas_nao_encontradas_maxifrota:
                placas_maxifrota_aba = processar_arquivos.placas_nao_encontradas_maxifrota[aba]
                if placas_maxifrota_aba:
                    # ENCONTRAR A LINHA DO CABEÇALHO "DEVOLUÇÃO:" NA COLUNA H
                    ultima_linha = None
                    for row in range(1, ws_divergencias.max_row + 1):
                        if ws_divergencias[f"H{row}"].value == "DEVOLUÇÃO:":
                            ultima_linha = row + 1  # COMEÇAR NA LINHA ABAIXO DO CABEÇALHO
                            break
                    
                    if ultima_linha is not None:
                        for placa in placas_maxifrota_aba:
                            # PULAR LINHAS JÁ PREENCHIDAS
                            while ws_divergencias[f"H{ultima_linha}"].value:
                                ultima_linha += 1
                            ws_divergencias[f"H{ultima_linha}"].value = f"PLACA {placa} NÃO ENCONTRADA NA PLANILHA MAXI FROTA."
                            log(f"  PLACA {placa} ({aba}) REGISTRADA NA COLUNA H COMO NÃO ENCONTRADA.")
                            ultima_linha += 1
                    else:
                        log(f"  AVISO: CABEÇALHO 'DEVOLUÇÃO:' NÃO ENCONTRADO NA COLUNA H DA ABA DIVERGÊNCIAS.")

    log(f"PROCESSAMENTO DO TICKET LOG CONCLUÍDO. REGISTROS ÚNICOS: {len(set(p for p in placas_ticketlog if p))} -> VALOR REAL TOTAL: {valor_total_emissao_ticketlog:.2f}")
    log(f"PROCESSAMENTO DA MAXI FROTA CONCLUÍDO. REGISTROS ÚNICOS: {len(set(p for p in placas_maxifrota if p))} -> VALOR REAL TOTAL: {valor_total_emissao_maxifrota:.2f}")

    # =========================
    # IDENTIFICAR PLACAS NOVAS
    # =========================
    log("IDENTIFICANDO PLACAS NOVAS NOS RELATÓRIOS...")

    equivalencias = {
        "DL SALVADOR": ["BOLANDEIRA", "DL SALVADOR"],
        "BONFIM": ["SENHOR DO BONFIM", "BONFIM"],
        "CAERN": ["CAERN-NATAL-RN", "CAERN"],
        "CATAGUASES": ["CATAGUASES", "COPASA ESGOTO CATAGUASES-MG"],
        "COPASA ESGOTO": ["COPASA ESGOTO"],
        "COPASA ÁGUA": ["COPASA AGUA", "COPASA ÁGUA", "COPASA LAFAIETE-MG  - ÁGUA"],
        "ALAGOINHAS": ["EMBASA - ALAGOINHAS-BA", "ALAGOINHAS"],
        "FEIRA DE SANTANA": ["FEIRA DE SANTANA", "EMBASA - FEIRA DE SANTANA-B"],
        "SAJ COMERCIAL": ["SAJ COMERCIAL", "EMBASA -SAJ COMERCIAL"],
        "SAJ MRR": ["EMBASA -SAJ MRR", "SAJ MRR"],
        "IGUÁ": ["IGUÁ SERGIPE", "IGUÁ"],
        "ITAPARICA": ["ITAPARICA", "ILHA ITAPARICA"],
        "JAGUAQUARA": ["JAGUAQUARA"],
        "SEDE": ["SEDE", "FROTAS", "SEDE MATRIZ"]
    }

    def identificar_aba_por_contrato(contrato_texto, wb):
        '''
        Identifica a aba correta na planilha de controle com base no texto do contrato.

        Args:
            contrato_texto (str): Texto do contrato a ser verificado.
            wb (Workbook): Objeto Workbook da planilha de controle.
        
        Returns:
            str or None: Nome da aba correspondente ou None se não encontrado.
        '''
        if not contrato_texto:
            return None

        contrato_upper = contrato_texto.upper()

        for nome_aba, variacoes in equivalencias.items():
            for variacao in variacoes:
                variacao_upper = variacao.upper()
                if variacao_upper in contrato_upper or contrato_upper in variacao_upper:
                    for aba_real in wb.sheetnames:
                        if nome_aba.upper() in aba_real.upper() or aba_real.upper() in nome_aba.upper():
                            return aba_real
        return None

    placas_controle_todas = set()
    for aba in wb_controle.sheetnames:
        ws_controle = wb_controle[aba]
        for row in range(7, ws_controle.max_row + 1):
            placa = str(ws_controle[f"A{row}"].value).strip() if ws_controle[f"A{row}"].value else None
            if placa and placa != "None":
                placas_controle_todas.add(placa)

    placas_ticketlog_unicas = set(p for p in placas_ticketlog if p and p != "None")
    placas_novas_ticketlog = placas_ticketlog_unicas - placas_controle_todas

    placas_maxifrota_unicas = set(p for p in placas_maxifrota if p and p != "None")
    placas_novas_maxifrota = placas_maxifrota_unicas - placas_controle_todas

    # ✅ ACESSAR ABA DIVERGÊNCIAS UMA ÚNICA VEZ
    ws_divergencias = wb_controle['DIVERGÊNCIAS']

    # ✅ ENCONTRAR O CABEÇALHO "PLACAS NOVAS:" NA COLUNA M
    linha_placas_novas = None
    for row in range(1, ws_divergencias.max_row + 1):
        if ws_divergencias[f"M{row}"].value == "PLACAS NOVAS:":
            linha_placas_novas = row + 1  # COMEÇAR NA LINHA ABAIXO DO CABEÇALHO
            break

    if linha_placas_novas is None:
        log("  AVISO: CABEÇALHO 'PLACAS NOVAS:' NÃO ENCONTRADO NA COLUNA M DA ABA DIVERGÊNCIAS.")
    else:
        # ✅ PROCESSAR PLACAS NOVAS DO TICKET LOG
        if placas_novas_ticketlog:
            log(f"ENCONTRADAS {len(placas_novas_ticketlog)} PLACA(S) NOVA(S) NO TICKET LOG")
            for placa_nova in sorted(placas_novas_ticketlog):
                indices = [i for i, p in enumerate(placas_ticketlog) if p == placa_nova]
                contratos = [contrato_ticketlog[idx] for idx in indices if contrato_ticketlog[idx] is not None]

                aba_encontrada = None
                for contrato in contratos:
                    aba_encontrada = identificar_aba_por_contrato(contrato, wb_controle)
                    if aba_encontrada:
                        break

                if not aba_encontrada:
                    aba_encontrada = "SEM ABA IDENTIFICADA"

                # PULAR LINHAS JÁ PREENCHIDAS
                while ws_divergencias[f"M{linha_placas_novas}"].value:
                    linha_placas_novas += 1

                km_validos = [km_ticketlog[idx] for idx in indices if km_ticketlog[idx] is not None]
                litros_validos = [litros_ticketlog[idx] for idx in indices if litros_ticketlog[idx] is not None]
                valor_validos = [valor_emissao_ticketlog[idx] for idx in indices if valor_emissao_ticketlog[idx] is not None]

                max_km = max(km_validos) if km_validos else 0
                total_litros = sum(litros_validos) if litros_validos else 0
                total_valor = sum(valor_validos) if valor_validos else 0
                contratos_texto = ", ".join(set(contratos)) if contratos else "SEM CONTRATO"

                ws_divergencias[f"M{linha_placas_novas}"].value = (
                    f"[TICKET LOG] {placa_nova} - {aba_encontrada} - {contratos_texto} - KM: {max_km} | LITROS: {total_litros:.2f} | VALOR: R$ {total_valor:.2f}"
                )
                log(f"  PLACA {placa_nova} REGISTRADA NA COLUNA M (TICKET LOG)")
                linha_placas_novas += 1

        # ✅ PROCESSAR PLACAS NOVAS DA MAXI FROTA
        if placas_novas_maxifrota:
            log(f"ENCONTRADAS {len(placas_novas_maxifrota)} PLACA(S) NOVA(S) NA MAXI FROTA")
            for placa_nova in sorted(placas_novas_maxifrota):
                indices = [i for i, p in enumerate(placas_maxifrota) if p == placa_nova]
                contratos = [contrato_maxifrota[idx] for idx in indices if contrato_maxifrota[idx] is not None]

                aba_encontrada = None
                for contrato in contratos:
                    aba_encontrada = identificar_aba_por_contrato(contrato, wb_controle)
                    if aba_encontrada:
                        break

                if not aba_encontrada:
                    aba_encontrada = "SEM ABA IDENTIFICADA"

                # PULAR LINHAS JÁ PREENCHIDAS
                while ws_divergencias[f"M{linha_placas_novas}"].value:
                    linha_placas_novas += 1

                hodro_validos = [hodrometro_values[idx] for idx in indices if hodrometro_values[idx] is not None]
                litros_validos = [litros_maxifrota_values[idx] for idx in indices if litros_maxifrota_values[idx] is not None]
                valor_validos = [valor_emissao_maxifrota_values[idx] for idx in indices if valor_emissao_maxifrota_values[idx] is not None]

                max_hodro = max(hodro_validos) if hodro_validos else 0
                total_litros = sum(litros_validos) if litros_validos else 0
                total_valor = sum(valor_validos) if valor_validos else 0
                contratos_texto = ", ".join(set(contratos)) if contratos else "SEM CONTRATO"

                ws_divergencias[f"M{linha_placas_novas}"].value = (
                    f"[MAXI FROTA] {placa_nova} - {aba_encontrada} - {contratos_texto} - KM: {max_hodro} | LITROS: {total_litros:.2f} | VALOR: R$ {total_valor:.2f}"
                )
                log(f"  PLACA {placa_nova} REGISTRADA NA COLUNA M (MAXI FROTA)")
                linha_placas_novas += 1

        if not placas_novas_ticketlog and not placas_novas_maxifrota:
            log("  NENHUMA PLACA NOVA IDENTIFICADA NOS RELATÓRIOS.")

    # =========================
    # SALVAR
    # =========================
    log("SALVANDO ARQUIVO COM SUCESSO...")
    try:
        wb_controle.save(path_controle)
        log(f"ARQUIVO SALVO COM SUCESSO: {path_controle}")

        return {
            "path_saida": path_controle,
            "relatorio": {
                "total_controle": len(placas_controle_todas),
                "total_ticketlog_registros": len(placas_ticketlog),
                "placas_controle_unicas": len(placas_controle_todas),
                "placas_tick_unicas": len(set(p for p in placas_ticketlog if p)),
                "duplicados_controle": 0,
                "duplicados_tick": len([p for p in placas_ticketlog if p]) - len(set(p for p in placas_ticketlog if p)),
                "placas_faltantes_na_base": f"{len(placas_novas_ticketlog)} DO TICKET LOG, {len(placas_novas_maxifrota)} DA MAXI FROTA"
            }
        }
    except PermissionError:
        raise RuntimeError("ERRO: NÃO FOI POSSÍVEL SALVAR O ARQUIVO. VERIFIQUE SE ELE ESTÁ ABERTO EM OUTRO PROGRAMA.")
    except Exception as e:
        raise RuntimeError(f"ERRO AO SALVAR ARQUIVO FINAL: {e}")


# =========================
# INTERFACE TKINTER
# =========================
class App:
    def __init__(self, root):
        self.root = root
        root.title("CONTROLE SEMANAL")
        root.geometry("760x520")
        root.resizable(True, True)

        self.semana_selecionada = None
        self.result_info = None

        # FRAME TOP
        top = ttk.Frame(root, padding=10)
        top.pack(fill="x")

        # SELEÇÃO ARQUIVO CONTROLE
        ttk.Label(top, text="ARQUIVO CONTROLE (BASE):").grid(row=0, column=0, sticky="w")
        self.entry_controle = ttk.Entry(top, width=70)
        self.entry_controle.grid(row=0, column=1, padx=5)
        ttk.Button(top, text="SELECIONAR", command=self.selecionar_controle).grid(row=0, column=2, padx=5)

        # SELEÇÃO ARQUIVO TICKETLOG
        ttk.Label(top, text="ARQUIVO TICKETLOG:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.entry_ticketlog = ttk.Entry(top, width=70)
        self.entry_ticketlog.grid(row=1, column=1, padx=5, pady=(8, 0))
        ttk.Button(top, text="SELECIONAR", command=self.selecionar_ticketlog).grid(row=1, column=2, padx=5, pady=(8, 0))

        # SELEÇÃO ARQUIVO MAXI FROTA
        ttk.Label(top, text="ARQUIVO MAXI FROTA:").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.entry_maxifrota = ttk.Entry(top, width=70)
        self.entry_maxifrota.grid(row=2, column=1, padx=5, pady=(8, 0))
        ttk.Button(top, text="SELECIONAR", command=self.selecionar_maxifrota).grid(row=2, column=2, padx=5, pady=(8, 0))

        # SELEÇÃO SEMANA
        ttk.Label(top, text="SELECIONAR SEMANA:").grid(row=4, column=0, sticky="w", pady=(8, 0))
        frame_semana = ttk.Frame(top)
        frame_semana.grid(row=4, column=1, columnspan=2, sticky="w", pady=(8, 0))

        ttk.Button(frame_semana, text="SEMANA 1", command=lambda: self.selecionar_semana(1)).grid(row=0, column=0, padx=5)
        ttk.Button(frame_semana, text="SEMANA 2", command=lambda: self.selecionar_semana(2)).grid(row=0, column=1, padx=5)
        ttk.Button(frame_semana, text="SEMANA 3", command=lambda: self.selecionar_semana(3)).grid(row=0, column=2, padx=5)
        ttk.Button(frame_semana, text="SEMANA 4", command=lambda: self.selecionar_semana(4)).grid(row=0, column=3, padx=5)
        ttk.Button(frame_semana, text="SEMANA 5", command=lambda: self.selecionar_semana(5)).grid(row=0, column=4, padx=5)

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

    def selecionar_semana(self, num):
        self.semana_selecionada = num
        self.log(f"SEMANA {num} SELECIONADA.")

    def selecionar_controle(self):
        path = filedialog.askopenfilename(
            title="SELECIONE A PLANILHA CONTROLE SEMANAL (EXCEL)",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if path:
            self.entry_controle.delete(0, tk.END)
            self.entry_controle.insert(0, path)

    def selecionar_ticketlog(self):
        path = filedialog.askopenfilename(
            title="SELECIONE O ARQUIVO TICKETLOG (EXCEL)",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if path:
            self.entry_ticketlog.delete(0, tk.END)
            self.entry_ticketlog.insert(0, path)

    def selecionar_maxifrota(self):
        path = filedialog.askopenfilename(
            title="SELECIONE O ARQUIVO MAXI FROTA (EXCEL)",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("Todos os arquivos", "*.*")]
        )
        if path:
            self.entry_maxifrota.delete(0, tk.END)
            self.entry_maxifrota.insert(0, path)

    def log(self, msg):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.txt_log.insert(tk.END, f"[{timestamp}] {msg}\n")
        self.txt_log.see(tk.END)
        self.txt_log.update_idletasks()

    def executar_processamento(self):
        path_controle = self.entry_controle.get().strip()
        path_ticketlog = self.entry_ticketlog.get().strip()
        path_maxifrota = self.entry_maxifrota.get().strip()

        if not path_controle or not os.path.exists(path_controle):
            messagebox.showerror("ERRO", "SELECIONE UM ARQUIVO DE CONTROLE VÁLIDO.")
            return
        if not path_ticketlog or not os.path.exists(path_ticketlog):
            messagebox.showerror("ERRO", "SELECIONE UM ARQUIVO TICKETLOG VÁLIDO.")
            return
        if not path_maxifrota or not os.path.exists(path_maxifrota):
            messagebox.showerror("ERRO", "SELECIONE UM ARQUIVO MAXI FROTA VÁLIDO.")
            return

        if self.semana_selecionada is None:
            messagebox.showerror("ERRO", "SELECIONE UMA SEMANA ANTES DE EXECUTAR.")
            return

        self.txt_log.delete("1.0", tk.END)
        self.log("INICIANDO PROCESSAMENTO...")

        try:
            # ATENÇÃO: SE NA SUA PLANILHA A SEMANA ESTIVER NA LINHA 3, TROQUE header_row_semana=3
            resultado = processar_arquivos(
                path_controle, path_ticketlog, path_maxifrota,
                semana_num=self.semana_selecionada,
                header_row_semana=3,
                log_callback=self.log
            )
        except Exception as e:
            erro_completo = traceback.format_exc()
            print(erro_completo)
            messagebox.showerror("ERRO AO PROCESSAR", f"{str(e)}\n\nVEJA O LOG PARA DETALHES.")
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

    def mostrar_relatorio(self):
        if not self.result_info:
            messagebox.showinfo("RELATÓRIO", "NENHUM PROCESSAMENTO REALIZADO AINDA.")
            return
        rel = self.result_info["relatorio"]
        texto = (
            f"TOTAL DE REGISTROS NA BASE: {rel['total_controle']}\n"
            f"TOTAL DE REGISTROS NO TICKETLOG: {rel['total_ticketlog_registros']}\n"
            f"PLACAS ÚNICAS NA BASE: {rel['placas_controle_unicas']}\n"
            f"PLACAS ÚNICAS NO TICKETLOG: {rel['placas_ticketlog_unicas']}\n\n"
            f"DUPLICADOS NA BASE: {rel['duplicados_controle']}\n"
            f"DUPLICADOS NO TICKETLOG: {rel['duplicados_ticketlog']}\n\n"
            f"PLACAS DO TICKETLOG NÃO ENCONTRADAS NA BASE (FALTANTES):\n{rel['placas_faltantes_na_base']}\n"
        )
        messagebox.showinfo("RELATÓRIO DE CONSISTÊNCIA", texto)

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


# =========================
# RODAR APP
# =========================
if __name__ == "__main__":
    verificar_acesso_com_excecao()

    root = tk.Tk()
    app = App(root)
    root.mainloop()
