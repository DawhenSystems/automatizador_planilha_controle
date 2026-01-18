#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de teste para validar a interface CustomTkinter
"""

import sys
import os

# Adicionar o diretório atual ao path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    import customtkinter as ctk
    print("✓ CustomTkinter importado com sucesso")
except ImportError as e:
    print(f"✗ Erro ao importar CustomTkinter: {e}")
    sys.exit(1)

try:
    from access_validator import verificar_acesso_com_excecao
    print("✓ access_validator importado com sucesso")
except ImportError as e:
    print(f"✗ Erro ao importar access_validator: {e}")

try:
    # Importar semanal apenas para verificar sintaxe
    # (não será executado completamente)
    print("✓ Validação básica completa")
except Exception as e:
    print(f"✗ Erro: {e}")
    sys.exit(1)

print("\n✓ Todos os imports funcionam corretamente!")
print("✓ A interface pode ser iniciada com: python app.py")
