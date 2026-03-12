#!/usr/bin/env python3
"""
Script para criar executável do programa Controle DDU
Uso: python build_exe.py
"""

import os
import sys
import subprocess
import shutil


def write_dist_env(base_dir, default_url="http://localhost:3000/upload-data"):
    """Cria arquivo .env na pasta dist ao lado do executável."""
    dist_env_path = os.path.join(base_dir, "dist", ".env")
    env_content = (
        "# URL de destino para envio do data.json\n"
        "# Pode ser dominio base (sera usado /upload-data automaticamente)\n"
        "# ou URL completa terminando com /upload-data\n"
        f"DATA_UPLOAD_URL={default_url}\n"
    )
    with open(dist_env_path, "w", encoding="utf-8") as f:
        f.write(env_content)
    return dist_env_path


def write_dist_config(base_dir):
    """Garante config.json ao lado do executável na pasta dist."""
    source_config_path = os.path.join(base_dir, "config.json")
    dist_config_path = os.path.join(base_dir, "dist", "config.json")

    if os.path.exists(source_config_path):
        shutil.copy2(source_config_path, dist_config_path)
        return dist_config_path

    default_config = {
        "Sub-titulo": "Polícia Militar de Minas Gerais",
        "Cabecalho_PDF": ["Polícia Militar de Minas Gerais"],
    }
    with open(dist_config_path, "w", encoding="utf-8") as f:
        import json
        json.dump(default_config, f, ensure_ascii=False, indent=2)
    return dist_config_path


def remove_dist_database(base_dir):
    """Remove documents.db da pasta dist para garantir primeira execução limpa."""
    dist_db_path = os.path.join(base_dir, "dist", "documents.db")
    if os.path.exists(dist_db_path):
        os.remove(dist_db_path)
        return dist_db_path
    return None

def build_executable():
    """Cria o executável usando PyInstaller"""
    app_name = "Controle_DDU"
    app_file = "Controle_documentos.py" if os.path.exists("Controle_documentos.py") else "app.py"
    
    # Caminho base
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    icon_candidates = [
        os.path.join(base_dir, "Icone.ico"),
        os.path.join(base_dir, "Icone.png"),
        os.path.join(base_dir, "pmmg.png"),
    ]
    icon_path = next((p for p in icon_candidates if os.path.exists(p)), None)

    # Comando PyInstaller
    cmd = [
        sys.executable, "-m", "pyinstaller",
        "--name", app_name,
        "--onefile",  # Gera um único arquivo .exe
        "--windowed",  # Remove console
        "--icon", icon_path,
        "--add-data", f"{base_dir}/ESCUDO PMMG.png:.",  # Incluir logomarca
        "--add-data", f"{base_dir}/Icone.png:.",  # Incluir ícone
        "--distpath", os.path.join(base_dir, "dist"),
        "--buildpath", os.path.join(base_dir, "build"),
        "--specpath", os.path.join(base_dir, "build"),
        app_file
    ]
    
    # Remover None do comando
    cmd = [c for c in cmd if c is not None]
    
    print(f"🔨 Compilando {app_name}...")
    print(f"Comando: {' '.join(cmd)}")
    
    result = subprocess.run(cmd, cwd=base_dir)
    
    if result.returncode == 0:
        exe_path = os.path.join(base_dir, "dist", f"{app_name}.exe")
        env_path = write_dist_env(base_dir)
        config_path = write_dist_config(base_dir)
        removed_db_path = remove_dist_database(base_dir)
        print(f"\n✅ Executável criado com sucesso!")
        print(f"📦 Arquivo: {exe_path}")
        print(f"⚙️  Configuração: {env_path}")
        print(f"⚙️  Configuração institucional: {config_path}")
        if removed_db_path:
            print(f"🗑️  Banco removido da distribuição: {removed_db_path}")
        print(f"\nPara distribuir:")
        print(f"1. Copie a pasta 'dist' ou o arquivo '{app_name}.exe'")
        print(f"2. Envie para outro computador")
        print(f"3. Execute o .exe - não precisa de Python instalado!")
    else:
        print(f"\n❌ Erro ao compilar")
        sys.exit(1)

if __name__ == "__main__":
    build_executable()
