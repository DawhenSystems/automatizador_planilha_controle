"""
Script auxiliar para gerar nova versão do executável usando build.spec
"""
import subprocess
import os
import sys
from datetime import datetime

# Configurações do aplicativo
NOME_APP = "Automatizador Planilha SEMANAL de Controle"
DATA_COMPILACAO = datetime.now().strftime("%H%M%S%d%m%y")
NOME_EXECUTAVEL = f"AutomatizadorPlanilhaSEMANALControle_c{DATA_COMPILACAO}"
VERSAO = "1"

# Tenta obter o hash do Git (opcional)
try:
    result = subprocess.run(
        ["git", "rev-parse", "--short", "HEAD"],
        capture_output=True,
        text=True,
        timeout=5
    )
    GIT_HASH = result.stdout.strip() if result.returncode == 0 else None
except:
    GIT_HASH = None

def update_spec_file():
    """Atualiza o nome do executável no arquivo build.spec"""
    spec_path = "build.spec"
    
    if not os.path.exists(spec_path):
        print(f"❌ Arquivo {spec_path} não encontrado!")
        print(f"💡 Crie o arquivo build.spec primeiro usando o comando:")
        print(f"   pyi-makespec --onefile --windowed --name {NOME_EXECUTAVEL} planilha_semanal.py")
        return False
    
    # Lê o arquivo spec
    with open(spec_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Substitui o nome do executável
    import re
    new_content = re.sub(
        r"name='[^']*'",
        f"name='{NOME_EXECUTAVEL}'",
        content
    )
    
    # Salva o arquivo atualizado
    with open(spec_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print(f"✅ Arquivo {spec_path} atualizado com nome: {NOME_EXECUTAVEL}")
    return True

def check_dependencies():
    """Verifica se as dependências necessárias estão instaladas"""
    try:
        import openpyxl
        import tkinter
        print("✅ Dependências básicas encontradas")
        return True
    except ImportError as e:
        print(f"❌ Dependência faltando: {e}")
        print("💡 Instale as dependências com: pip install -r requirements.txt")
        return False

def build_exe():
    """Gera o executável usando o arquivo build.spec"""
    
    print("=" * 70)
    print(f"🔨 Gerando executável do {NOME_APP}")
    print("=" * 70)
    print(f"📦 Nome do executável: {NOME_EXECUTAVEL}.exe")
    print(f"🏷️  Versão: {VERSAO}")
    if GIT_HASH:
        print(f"🔖 Git Hash: {GIT_HASH}")
    print(f"📅 Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("=" * 70)
    
    # Verifica dependências
    if not check_dependencies():
        return
    
    # Verifica se o arquivo planilha_semanal.py existe
    if not os.path.exists("planilha_semanal.py"):
        print("❌ Arquivo planilha_semanal.py não encontrado!")
        return
    
    # Atualiza o nome no spec file
    if not update_spec_file():
        return
    
    # Limpa builds anteriores
    print("\n🧹 Limpando builds anteriores...")
    if os.path.exists("build"):
        import shutil
        shutil.rmtree("build")
        print("   ✅ Pasta build removida")
    
    # Executa o PyInstaller com o arquivo spec
    print("\n⏳ Iniciando build com PyInstaller...\n")
    result = subprocess.run(["pyinstaller", "build.spec", "--clean", "--noconfirm"])
    
    if result.returncode == 0:
        exe_path = os.path.join("dist", f"{NOME_EXECUTAVEL}.exe")
        if os.path.exists(exe_path):
            size_mb = os.path.getsize(exe_path) / (1024 * 1024)
            print("\n" + "=" * 70)
            print(f"✅ Executável gerado com sucesso!")
            print(f"📁 Local: {exe_path}")
            print(f"📊 Tamanho: {size_mb:.2f} MB")
            print("=" * 70)
            
            # Pergunta se deseja abrir a pasta
            try:
                response = input("\n❓ Deseja abrir a pasta dist? (s/n): ").strip().lower()
                if response == 's':
                    if os.name == 'nt':
                        os.startfile("dist")
                    else:
                        subprocess.run(["xdg-open", "dist"])
            except:
                pass
        else:
            print("\n❌ Executável não foi encontrado na pasta dist!")
    else:
        print("\n❌ Erro ao gerar executável!")
        print("💡 Verifique os logs acima para mais detalhes")

if __name__ == "__main__":
    try:
        build_exe()
    except KeyboardInterrupt:
        print("\n\n⚠️  Build cancelado pelo usuário")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)