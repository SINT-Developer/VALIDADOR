"""
Script para automatizar o processo de release do Validador SINT.

Uso:
    python release.py          # Incrementa patch (1.0.0 -> 1.0.1)
    python release.py minor    # Incrementa minor (1.0.0 -> 1.1.0)
    python release.py major    # Incrementa major (1.0.0 -> 2.0.0)
    python release.py 1.2.3    # Define versao especifica

O script:
1. Atualiza APP_VERSION no validador_standalone.py
2. Gera o exe com PyInstaller
3. Cria release no GitHub com o exe anexado
4. Atualiza o Gist com a nova versao e URL de download

Configuracao inicial (apenas uma vez):
1. Crie um Personal Access Token no GitHub: https://github.com/settings/tokens
   - Permissoes necessarias: "gist" e "repo" (para criar releases)
2. Crie um arquivo .env na pasta do projeto com:
   GITHUB_TOKEN=seu_token_aqui
"""

import sys
import os
import re
import subprocess
import urllib.request
import json

# Configuracoes
GITHUB_REPO = "SINT-Developer/VALIDADOR"
GIST_ID = "a38baad856a6149526948d7c0c360ab9"
VALIDATOR_FILE = "validador_standalone.py"
EXE_NAME = "Validador SINT.exe"
EXE_PATH = os.path.join("dist", EXE_NAME)


def obter_versao_atual():
    """Le a versao atual do codigo"""
    with open(VALIDATOR_FILE, 'r', encoding='utf-8') as f:
        conteudo = f.read()

    match = re.search(r'APP_VERSION\s*=\s*"([^"]*)"', conteudo)
    if match:
        return match.group(1)
    return "0.0.0"


def incrementar_versao(versao_atual, tipo="patch"):
    """Incrementa a versao baseado no tipo (major, minor, patch)"""
    partes = [int(x) for x in versao_atual.split(".")]

    while len(partes) < 3:
        partes.append(0)

    if tipo == "major":
        partes[0] += 1
        partes[1] = 0
        partes[2] = 0
    elif tipo == "minor":
        partes[1] += 1
        partes[2] = 0
    else:  # patch
        partes[2] += 1

    return ".".join(str(p) for p in partes)


def carregar_token():
    """Carrega o token do arquivo .env"""
    env_path = os.path.join(os.path.dirname(__file__), ".env")
    if not os.path.exists(env_path):
        print("ERRO: Arquivo .env nao encontrado.")
        print("Crie um arquivo .env com: GITHUB_TOKEN=seu_token_aqui")
        print("Gere o token em: https://github.com/settings/tokens")
        print("Permissoes necessarias: 'gist' e 'repo'")
        return None

    with open(env_path, 'r') as f:
        for line in f:
            if line.startswith("GITHUB_TOKEN="):
                return line.strip().split("=", 1)[1]

    print("ERRO: GITHUB_TOKEN nao encontrado no .env")
    return None


def atualizar_versao_codigo(nova_versao):
    """Atualiza APP_VERSION no codigo"""
    with open(VALIDATOR_FILE, 'r', encoding='utf-8') as f:
        conteudo = f.read()

    novo_conteudo = re.sub(
        r'APP_VERSION\s*=\s*"[^"]*"',
        f'APP_VERSION = "{nova_versao}"',
        conteudo
    )

    if novo_conteudo == conteudo:
        print("AVISO: APP_VERSION nao encontrado ou ja esta na versao correta")
        return False

    with open(VALIDATOR_FILE, 'w', encoding='utf-8') as f:
        f.write(novo_conteudo)

    print(f"[OK] APP_VERSION atualizado para {nova_versao}")
    return True


def gerar_exe():
    """Gera o executavel com PyInstaller"""
    print("[...] Gerando executavel com PyInstaller...")

    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--icon=icon.ico",
        "--name", "Validador SINT",
        "validador_standalone.py"
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        print(f"ERRO ao gerar exe: {result.stderr}")
        return False

    print(f"[OK] Executavel gerado em {EXE_PATH}")
    return True


def criar_github_release(nova_versao, token):
    """Cria uma release no GitHub e faz upload do exe"""
    tag = f"v{nova_versao}"

    # 1. Criar a release
    print(f"[...] Criando release {tag} no GitHub...")

    release_data = {
        "tag_name": tag,
        "name": f"Validador SINT {tag}",
        "body": f"Release automatica v{nova_versao}",
        "draft": False,
        "prerelease": False
    }

    req = urllib.request.Request(
        f"https://api.github.com/repos/{GITHUB_REPO}/releases",
        data=json.dumps(release_data).encode('utf-8'),
        headers={
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
            "Content-Type": "application/json"
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(req) as response:
            release_info = json.loads(response.read().decode('utf-8'))
            upload_url = release_info["upload_url"].replace("{?name,label}", "")
            release_id = release_info["id"]
            print(f"[OK] Release {tag} criada")
    except urllib.error.HTTPError as e:
        error_body = e.read().decode('utf-8')
        if "already_exists" in error_body:
            print(f"[...] Release {tag} ja existe, atualizando...")
            # Buscar release existente
            return atualizar_github_release(nova_versao, token)
        print(f"ERRO ao criar release: {e.code} - {error_body}")
        return None

    # 2. Upload do exe
    print(f"[...] Fazendo upload do exe...")

    with open(EXE_PATH, 'rb') as f:
        exe_data = f.read()

    upload_req = urllib.request.Request(
        f"{upload_url}?name={EXE_NAME.replace(' ', '%20')}",
        data=exe_data,
        headers={
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
            "Content-Type": "application/octet-stream"
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(upload_req) as response:
            asset_info = json.loads(response.read().decode('utf-8'))
            download_url = asset_info["browser_download_url"]
            print(f"[OK] Exe uploaded: {download_url}")
            return download_url
    except urllib.error.HTTPError as e:
        print(f"ERRO ao fazer upload: {e.code} - {e.read().decode('utf-8')}")
        return None


def atualizar_github_release(nova_versao, token):
    """Atualiza uma release existente no GitHub"""
    tag = f"v{nova_versao}"

    # Buscar release pela tag
    req = urllib.request.Request(
        f"https://api.github.com/repos/{GITHUB_REPO}/releases/tags/{tag}",
        headers={
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json"
        }
    )

    try:
        with urllib.request.urlopen(req) as response:
            release_info = json.loads(response.read().decode('utf-8'))
            release_id = release_info["id"]
            upload_url = release_info["upload_url"].replace("{?name,label}", "")

            # Deletar assets antigos
            for asset in release_info.get("assets", []):
                delete_req = urllib.request.Request(
                    f"https://api.github.com/repos/{GITHUB_REPO}/releases/assets/{asset['id']}",
                    headers={
                        "Authorization": f"token {token}",
                        "Accept": "application/vnd.github.v3+json"
                    },
                    method="DELETE"
                )
                urllib.request.urlopen(delete_req)

    except urllib.error.HTTPError as e:
        print(f"ERRO ao buscar release: {e.code}")
        return None

    # Upload do novo exe
    print(f"[...] Fazendo upload do exe...")

    with open(EXE_PATH, 'rb') as f:
        exe_data = f.read()

    upload_req = urllib.request.Request(
        f"{upload_url}?name={EXE_NAME.replace(' ', '%20')}",
        data=exe_data,
        headers={
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
            "Content-Type": "application/octet-stream"
        },
        method="POST"
    )

    try:
        with urllib.request.urlopen(upload_req) as response:
            asset_info = json.loads(response.read().decode('utf-8'))
            download_url = asset_info["browser_download_url"]
            print(f"[OK] Exe uploaded: {download_url}")
            return download_url
    except urllib.error.HTTPError as e:
        print(f"ERRO ao fazer upload: {e.code} - {e.read().decode('utf-8')}")
        return None


def atualizar_gist(nova_versao, download_url, token):
    """Atualiza o Gist com a nova versao e URL"""
    print(f"[...] Atualizando Gist...")

    url = f"https://api.github.com/gists/{GIST_ID}"

    data = {
        "files": {
            "version.json": {
                "content": json.dumps({
                    "version": nova_versao,
                    "download_url": download_url
                }, indent=2)
            }
        }
    }

    req = urllib.request.Request(
        url,
        data=json.dumps(data).encode('utf-8'),
        headers={
            "Authorization": f"token {token}",
            "Accept": "application/vnd.github.v3+json",
            "Content-Type": "application/json"
        },
        method="PATCH"
    )

    try:
        with urllib.request.urlopen(req) as response:
            if response.status == 200:
                print(f"[OK] Gist atualizado para versao {nova_versao}")
                return True
    except urllib.error.HTTPError as e:
        print(f"ERRO ao atualizar Gist: {e.code} - {e.reason}")
        return False

    return False


def main():
    versao_atual = obter_versao_atual()

    # Determinar nova versao
    if len(sys.argv) < 2:
        nova_versao = incrementar_versao(versao_atual, "patch")
    elif sys.argv[1] in ("major", "minor", "patch"):
        nova_versao = incrementar_versao(versao_atual, sys.argv[1])
    elif re.match(r'^\d+\.\d+\.\d+$', sys.argv[1]):
        nova_versao = sys.argv[1]
    else:
        print("Uso:")
        print("  python release.py          # Incrementa patch (1.0.0 -> 1.0.1)")
        print("  python release.py minor    # Incrementa minor (1.0.0 -> 1.1.0)")
        print("  python release.py major    # Incrementa major (1.0.0 -> 2.0.0)")
        print("  python release.py 1.2.3    # Define versao especifica")
        sys.exit(1)

    print(f"\n=== Release v{nova_versao} (atual: v{versao_atual}) ===\n")

    # 1. Carregar token
    token = carregar_token()
    if not token:
        sys.exit(1)

    # 2. Atualizar versao no codigo
    if not atualizar_versao_codigo(nova_versao):
        sys.exit(1)

    # 3. Gerar exe
    if not gerar_exe():
        sys.exit(1)

    # 4. Criar release no GitHub
    download_url = criar_github_release(nova_versao, token)
    if not download_url:
        sys.exit(1)

    # 5. Atualizar Gist
    if not atualizar_gist(nova_versao, download_url, token):
        sys.exit(1)

    print(f"\n=== Release v{nova_versao} concluido! ===")
    print(f"\nDownload: {download_url}")


if __name__ == "__main__":
    main()
