#!/usr/bin/env python3
"""
Script para criar executável portável do Controle DDU usando PyInstaller
Execução: python build_executavel.py
"""

import subprocess
import sys
import os
import shutil
import json


def write_dist_env(dist_path, default_url="http://localhost:3000/upload-data"):
    """Cria arquivo .env na pasta dist ao lado do executável."""
    dist_env_path = os.path.join(dist_path, ".env")
    env_content = (
        "# URL de destino para envio do data.json\n"
        "# Pode ser dominio base (sera usado /upload-data automaticamente)\n"
        "# ou URL completa terminando com /upload-data\n"
        f"DATA_UPLOAD_URL={default_url}\n"
    )
    with open(dist_env_path, "w", encoding="utf-8") as f:
        f.write(env_content)
    return dist_env_path


def write_dist_config(current_dir, dist_path):
    """Garante config.json ao lado do executável na pasta dist."""
    source_config_path = os.path.join(current_dir, "config.json")
    dist_config_path = os.path.join(dist_path, "config.json")

    if os.path.exists(source_config_path):
        shutil.copy2(source_config_path, dist_config_path)
        return dist_config_path

    default_config = {
        "Sub-titulo": "Polícia Militar de Minas Gerais",
        "Cabecalho_PDF": ["Polícia Militar de Minas Gerais"],
    }
    with open(dist_config_path, "w", encoding="utf-8") as f:
        json.dump(default_config, f, ensure_ascii=False, indent=2)
    return dist_config_path


def remove_dist_database(dist_path):
    """Remove documents.db da pasta dist para garantir primeira execução limpa."""
    dist_db_path = os.path.join(dist_path, "documents.db")
    if os.path.exists(dist_db_path):
        os.remove(dist_db_path)
        return dist_db_path
    return None

def main():
    print("=" * 70)
    print("🔨 INICIANDO COMPILAÇÃO DO CONTROLE_DDU PARA EXECUTÁVEL (.EXE)")
    print("=" * 70)
    
    # Caminhos
    current_dir = os.path.dirname(os.path.abspath(__file__))
    app_candidates = [
        os.path.join(current_dir, "Controle_documentos.py"),
        os.path.join(current_dir, "app.py"),
    ]
    app_path = next((p for p in app_candidates if os.path.exists(p)), None)
    if not app_path:
        print("\n❌ Arquivo principal não encontrado (Controle_documentos.py/app.py)")
        return False
    dist_path = os.path.join(current_dir, "dist")
    build_path = os.path.join(current_dir, "build")
    escudo_path = os.path.join(current_dir, "ESCUDO PMMG.png")
    icon_png_path = os.path.join(current_dir, "Icone.png")
    check_icon_path = os.path.join(current_dir, "check_blue.svg")
    
    # Verificar se PyInstaller está instalado
    try:
        import PyInstaller
        print("✓ PyInstaller encontrado")
    except ImportError:
        print("\n❌ PyInstaller não está instalado!")
        print("📦 Instalando PyInstaller...")
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller instalado com sucesso")
    
    # Limpar builds anteriores
    if os.path.exists(dist_path):
        print(f"🗑️  Limpando pasta dist antiga...")
        shutil.rmtree(dist_path)
    if os.path.exists(build_path):
        print(f"🗑️  Limpando pasta build antiga...")
        shutil.rmtree(build_path)
    
    print("\n" + "=" * 70)
    print("📦 CONFIGURANDO PACOTE...")
    print("=" * 70)
    
    # Comando PyInstaller
    cmd = [
        sys.executable,
        "-m", "PyInstaller",
        "--name=Controle_DDU",
        "--onefile",              # Gerar arquivo único
        "--windowed",              # Sem console (interface gráfica)
        "--clean",                 # Limpar cache
        "--noconfirm",             # Não pedir confirmação
        "--collect-all=PySide6",   # Inclui plugins/binários Qt necessários
        "--collect-all=plotly",    # Inclui assets dinâmicos do Plotly
        "--collect-all=kaleido",   # Necessário para exportação de imagens dos gráficos
        "--collect-all=reportlab", # Garante módulos/fontes auxiliares do PDF
        "--collect-all=Crypto",    # Garante bibliotecas do pycryptodome (AES/padding)
        "--hidden-import=requests", # Dependência usada no upload do data.json
    ]
    
    # Adicionar logo se existir
    if os.path.exists(escudo_path):
        cmd.extend(["--add-data", f"{escudo_path};."])
        print(f"✓ Escudo PMMG incluído: {escudo_path}")
    else:
        print(f"⚠️  Escudo não encontrado em: {escudo_path}")

    if os.path.exists(icon_png_path):
        cmd.extend(["--add-data", f"{icon_png_path};."])
        print(f"✓ Ícone PNG incluído para runtime: {icon_png_path}")
    else:
        print(f"⚠️  Ícone PNG não encontrado em: {icon_png_path}")

    if os.path.exists(check_icon_path):
        cmd.extend(["--add-data", f"{check_icon_path};."])
        print(f"✓ Ícone de checkbox incluído: {check_icon_path}")
    else:
        print(f"⚠️  Ícone de checkbox não encontrado em: {check_icon_path}")

    # Adicionar ícone se existir
    icon_candidates = [
        os.path.join(current_dir, "Icone.ico"),
        os.path.join(current_dir, "Icone.png"),
        os.path.join(current_dir, "icon.ico"),
    ]
    icon_path = next((p for p in icon_candidates if os.path.exists(p)), None)
    if icon_path:
        cmd.extend(["--icon", icon_path])
        print(f"✓ Ícone incluído: {icon_path}")
    else:
        print("⚠️  Ícone não encontrado (Icone.ico/Icone.png/icon.ico)")
    
    cmd.append(app_path)
    
    print("\n" + "=" * 70)
    print("⚙️  COMPILANDO (Pode levar alguns minutos)...")
    print("=" * 70)
    
    # Executar PyInstaller
    result = subprocess.run(cmd, cwd=current_dir)
    
    if result.returncode == 0:
        exe_file = os.path.join(dist_path, "Controle_DDU.exe")
        
        if os.path.exists(exe_file):
            file_size = os.path.getsize(exe_file) / (1024 * 1024)  # MB
            env_file = write_dist_env(dist_path)
            config_file = write_dist_config(current_dir, dist_path)
            removed_db_path = remove_dist_database(dist_path)
            
            print("\n" + "=" * 70)
            print("✅ EXECUTÁVEL CRIADO COM SUCESSO!")
            print("=" * 70)
            print(f"📦 Arquivo: {exe_file}")
            print(f"⚙️  Configuração: {env_file}")
            print(f"⚙️  Configuração institucional: {config_file}")
            if removed_db_path:
                print(f"🗑️  Banco removido da distribuição: {removed_db_path}")
            print(f"📊 Tamanho: {file_size:.2f} MB")
            print(f"\n📋 COMO USAR:")
            print(f"   1. Localize o arquivo em: {dist_path}")
            print(f"   2. Copie 'Controle_DDU.exe' para qualquer computador Windows")
            print(f"   3. Execute diretamente - NÃO precisa Python instalado!")
            print(f"   4. O banco de dados será criado automaticamente")
            print("\n💡 OBSERVAÇÕES:")
            print(f"   • O executável é portátil e autocontido")
            print(f"   • Todas as dependências estão incluídas")
            print(f"   • O banco SQLite será criado na primeira execução")
            print(f"   • A logo PMMG está embutida no executável")
            print("=" * 70)
            
            return True
        else:
            print("\n❌ Arquivo executável não foi encontrado!")
            return False
    else:
        print("\n❌ ERRO NA COMPILAÇÃO")
        print("Verifique as mensagens de erro acima")
        return False

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
