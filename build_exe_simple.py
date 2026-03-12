#!/usr/bin/env python3
"""
Script para criar executável portável do Controle DDU usando PyInstaller
Execução: python build_exe_simple.py
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
    current_dir = os.path.dirname(os.path.abspath(__file__))
    venv_python = os.path.join(current_dir, ".venv", "Scripts", "pyinstaller.exe")
    app_path = os.path.join(current_dir, "Controle_documentos.py")
    if not os.path.exists(app_path):
        app_path = os.path.join(current_dir, "app.py")
    dist_path = os.path.join(current_dir, "dist")
    escudo_path = os.path.join(current_dir, "ESCUDO PMMG.png")
    icon_candidates = [
        os.path.join(current_dir, "Icone.ico"),
        os.path.join(current_dir, "Icone.png"),
    ]
    icon_path = next((p for p in icon_candidates if os.path.exists(p)), None)
    
    # Criar pasta dist se não existir
    os.makedirs(dist_path, exist_ok=True)
    
    print("=" * 60)
    print("🔨 Compilando Controle_DDU para executável (.exe)...")
    print("=" * 60)
    
    cmd = [
        venv_python,
        "--name", "Controle_DDU",
        "--onefile",
        "--windowed",
        "--add-data", f"{escudo_path}:.",
        "--add-data", f"{os.path.join(current_dir, 'Icone.png')}:.",
        "--distpath", dist_path,
        app_path
    ]

    if icon_path:
        cmd.extend(["--icon", icon_path])
    
    result = subprocess.run(cmd)
    
    if result.returncode == 0:
        exe_file = os.path.join(dist_path, "Controle_DDU.exe")
        env_file = write_dist_env(dist_path)
        config_file = write_dist_config(current_dir, dist_path)
        removed_db_path = remove_dist_database(dist_path)
        print("\n" + "=" * 60)
        print("✅ EXECUTÁVEL CRIADO COM SUCESSO!")
        print("=" * 60)
        print(f"📦 Arquivo: {exe_file}")
        print(f"⚙️  Configuração: {env_file}")
        print(f"⚙️  Configuração institucional: {config_file}")
        if removed_db_path:
            print(f"🗑️  Banco removido da distribuição: {removed_db_path}")
        print(f"\n📋 Próximos passos:")
        print(f"   1. Copie o arquivo Controle_DDU.exe")
        print(f"   2. Envie para outro computador")
        print(f"   3. Execute direto - não precisa Python!")
        print("=" * 60)
    else:
        print("\n❌ Erro na compilação")
        sys.exit(1)

if __name__ == "__main__":
    main()
