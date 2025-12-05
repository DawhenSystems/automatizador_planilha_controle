"""
Módulo de validação de acesso para o automatizador de planilha de controle.
Verifica o arquivo access_config.json no repositório GitHub antes de permitir execução.
"""

import json
import sys
from tkinter import messagebox
import urllib.request
import urllib.error

# Configurações
GITHUB_REPO = "DawhenSystems/automatizador_planilha_controle"
CONFIG_FILE = "access_config.json"
GITHUB_RAW_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{CONFIG_FILE}"
VERSION_ATUAL = "0.1.0"
CHAVE_ACESSO = "c46e3b3ac80d03a49fb7c54a027723fdced9b5fc65093489f06e0eea85548e8f"


def baixar_config_github():
    """
    Baixa o arquivo de configuração do GitHub.
    Retorna uma tupla (config_dict, mensagem_erro) onde:
    - config_dict: dicionário com a configuração ou None em caso de erro
    - mensagem_erro: string descritiva do erro ou None se sucesso
    """
    try:
        with urllib.request.urlopen(GITHUB_RAW_URL, timeout=10) as response:
            if response.status != 200:
                return None, f"Servidor retornou código {response.status}"
            
            conteudo = response.read().decode('utf-8')
            config = json.loads(conteudo)
            return config, None
            
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None, (
                "Arquivo de configuração não encontrado no repositório.\n"
                f"Certifique-se de que '{CONFIG_FILE}' existe na branch 'main'."
            )
        elif e.code == 403:
            return None, "Acesso negado ao repositório (403 Forbidden)."
        elif e.code == 500:
            return None, "Erro no servidor do GitHub (500)."
        else:
            return None, f"Erro HTTP {e.code}: {e.reason}"
            
    except urllib.error.URLError as e:
        erro_str = str(e.reason).lower()
        if "timed out" in erro_str or "timeout" in erro_str:
            return None, "Tempo de conexão esgotado. Verifique sua internet."
        elif "certificate" in erro_str or "ssl" in erro_str:
            return None, "Erro de certificado SSL. Verifique as configurações de segurança."
        elif "name resolution" in erro_str or "getaddrinfo" in erro_str or "nodename" in erro_str:
            return None, "Não foi possível resolver o nome do servidor. Verifique sua conexão com a internet."
        elif "refused" in erro_str or "connection refused" in erro_str:
            return None, "Conexão recusada pelo servidor. Tente novamente mais tarde."
        elif "reset" in erro_str or "connection reset" in erro_str:
            return None, "Conexão interrompida. Verifique sua rede e tente novamente."
        elif "unreachable" in erro_str:
            return None, "Servidor inacessível. Verifique sua conexão de rede."
        else:
            return None, f"Erro de conexão: {e.reason}"
            
    except json.JSONDecodeError as e:
        return None, (
            f"Arquivo de configuração com formato JSON inválido.\n"
            f"Erro na linha {e.lineno}, coluna {e.colno}: {e.msg}"
        )
        
    except UnicodeDecodeError as e:
        return None, f"Erro ao decodificar o arquivo. Encoding inválido: {e.encoding}"
        
    except TimeoutError:
        return None, "Tempo de conexão esgotado após 10 segundos. Verifique sua internet."
        
    except ConnectionError as e:
        return None, f"Erro de conexão com o servidor: {str(e)}"
        
    except Exception as e:
        return None, f"Erro inesperado: {type(e).__name__} - {str(e)}"


def validar_versao(versao_requerida):
    """
    Valida se a versão atual do aplicativo é compatível com a versão requerida.
    Retorna True se a versão atual é >= versão requerida.
    """
    try:
        versao_atual_parts = [int(x) for x in VERSION_ATUAL.split('.')]
        versao_req_parts = [int(x) for x in versao_requerida.split('.')]
        
        # Compara versões (major.minor.patch)
        for i in range(max(len(versao_atual_parts), len(versao_req_parts))):
            atual = versao_atual_parts[i] if i < len(versao_atual_parts) else 0
            req = versao_req_parts[i] if i < len(versao_req_parts) else 0
            
            if atual > req:
                return True
            elif atual < req:
                return False
        
        return True
    except (ValueError, AttributeError, IndexError) as e:
        print(f"Erro ao validar versão: {e}")
        return False


def validar_acesso():
    """
    Função principal de validação de accesso.
    Retorna True se o accesso for permitido, False caso contrário.
    """
    print("Validando accesso ao aplicativo...")
    
    # Baixa configuração do GitHub
    config, erro = baixar_config_github()
    
    if config is None:
        messagebox.showerror(
            "ERRO DE VALIDAÇÃO",
            f"Não foi possível validar o acesso:\n\n{erro}\n\n"
            "Entre em contato com o suporte se o problema persistir."
        )
        return False
    
    # Valida estrutura do arquivo JSON
    campos_obrigatorios = ["access_enabled", "version_required", "access_key"]
    campos_faltantes = [campo for campo in campos_obrigatorios if campo not in config]
    
    if campos_faltantes:
        messagebox.showerror(
            "ERRO DE CONFIGURAÇÃO",
            f"Arquivo de configuração incompleto.\n"
            f"Campos faltantes: {', '.join(campos_faltantes)}\n\n"
            "Entre em contato com o suporte."
        )
        return False
    
    # Verifica se o accesso está habilitado
    access_enabled = config.get("access_enabled")
    if not isinstance(access_enabled, bool):
        messagebox.showerror(
            "ERRO DE CONFIGURAÇÃO",
            f"Campo 'access_enabled' inválido.\n"
            f"Esperado: true/false, Recebido: {type(access_enabled).__name__}\n\n"
            "Entre em contato com o suporte."
        )
        return False
        
    if not access_enabled:
        messagebox.showerror(
            "ACESSO NEGADO",
            "O aplicativo está temporariamente desabilitado.\n"
            "Entre em contato com o suporte."
        )
        return False
    
    # Valida a versão
    versao_requerida = config.get("version_required", "0.1.0")
    if not isinstance(versao_requerida, str):
        messagebox.showerror(
            "ERRO DE CONFIGURAÇÃO",
            f"Campo 'version_required' inválido: {versao_requerida}\n\n"
            "Entre em contato com o suporte."
        )
        return False
        
    if not validar_versao(versao_requerida):
        messagebox.showerror(
            "VERSÃO DESATUALIZADA",
            f"Esta versão do aplicativo ({VERSION_ATUAL}) está desatualizada.\n"
            f"Versão mínima requerida: {versao_requerida}\n\n"
            "Baixe a versão mais recente para continuar."
        )
        return False
    
    # Valida a chave de accesso única
    chave_config = config.get("access_key", "")
    
    if not isinstance(chave_config, str):
        messagebox.showerror(
            "ERRO DE CONFIGURAÇÃO",
            f"Campo 'access_key' inválido: {type(chave_config).__name__}\n\n"
            "Entre em contato com o suporte."
        )
        return False
        
    if not chave_config:
        messagebox.showerror(
            "ERRO DE CONFIGURAÇÃO",
            "Chave de acesso vazia no arquivo de configuração.\n\n"
            "Entre em contato com o suporte."
        )
        return False
    
    if CHAVE_ACESSO != chave_config:
        messagebox.showerror(
            "ACESSO NEGADO",
            "Chave de acesso inválida.\n\n"
            "Este executável não está autorizado.\n"
            "Entre em contato com o suporte para obter a versão correta."
        )
        return False
    
    print("Acesso validado com sucesso!")
    return True


def verificar_acesso_com_excecao():
    """
    Verifica o acesso e encerra o programa se não for autorizado.
    """
    if not validar_acesso():
        sys.exit(1)
