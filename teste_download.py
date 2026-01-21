"""
Script para testar diferentes abordagens de download do OneDrive.
"""
import urllib.request
import ssl
import tempfile
import os

DOWNLOAD_URL = "https://sintsistemas-my.sharepoint.com/:u:/g/personal/erick_sint_com_br/IQD0gjWOwwbjTbSwuszn_rDdAfji_siYsNaFAN-2HZ-aiOs?e=3pCrW9&download=1"

# Tamanho minimo esperado (exe real tem ~31MB)
TAMANHO_MINIMO = 1000000  # 1MB


def verificar_arquivo(path):
    """Verifica se o arquivo baixado e um exe valido"""
    if not os.path.exists(path):
        return False, "Arquivo nao existe"

    size = os.path.getsize(path)
    if size < TAMANHO_MINIMO:
        # Ler primeiros bytes para ver se e HTML
        with open(path, 'rb') as f:
            inicio = f.read(500)

        if b'<html' in inicio.lower() or b'<!doctype' in inicio.lower():
            return False, f"Recebeu HTML em vez do exe ({size} bytes)"
        return False, f"Arquivo muito pequeno ({size} bytes)"

    # Verificar se comeca com MZ (exe valido)
    with open(path, 'rb') as f:
        magic = f.read(2)

    if magic == b'MZ':
        return True, f"Exe valido ({size} bytes)"
    return False, f"Nao e um exe valido ({size} bytes)"


def teste_1_urlretrieve_simples():
    """Abordagem original - urlretrieve sem headers"""
    print("\n=== TESTE 1: urlretrieve simples ===")
    try:
        temp_file = os.path.join(tempfile.gettempdir(), "teste1.exe")
        urllib.request.urlretrieve(DOWNLOAD_URL, temp_file)
        valido, msg = verificar_arquivo(temp_file)
        print(f"[{'OK' if valido else 'ERRO'}] {msg}")
        if os.path.exists(temp_file):
            os.remove(temp_file)
        return valido
    except Exception as e:
        print(f"[ERRO] {e}")
        return False


def teste_2_urlopen_com_headers():
    """Abordagem com headers de navegador"""
    print("\n=== TESTE 2: urlopen com headers ===")
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
        }

        req = urllib.request.Request(DOWNLOAD_URL, headers=headers)

        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

        temp_file = os.path.join(tempfile.gettempdir(), "teste2.exe")
        with urllib.request.urlopen(req, timeout=30, context=ctx) as response:
            # Mostrar URL final (pode ter redirecionado)
            print(f"  URL final: {response.url[:80]}...")
            print(f"  Content-Type: {response.headers.get('Content-Type', 'N/A')}")
            with open(temp_file, 'wb') as f:
                f.write(response.read())

        valido, msg = verificar_arquivo(temp_file)
        print(f"[{'OK' if valido else 'ERRO'}] {msg}")
        if os.path.exists(temp_file):
            os.remove(temp_file)
        return valido
    except Exception as e:
        print(f"[ERRO] {e}")
        return False


def teste_3_opener_com_headers():
    """Abordagem com opener global"""
    print("\n=== TESTE 3: opener com headers ===")
    try:
        opener = urllib.request.build_opener()
        opener.addheaders = [
            ('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'),
        ]
        urllib.request.install_opener(opener)

        temp_file = os.path.join(tempfile.gettempdir(), "teste3.exe")
        urllib.request.urlretrieve(DOWNLOAD_URL, temp_file)

        size = os.path.getsize(temp_file)
        print(f"[OK] Baixou {size} bytes para {temp_file}")
        os.remove(temp_file)
        return True
    except Exception as e:
        print(f"[ERRO] {e}")
        return False


def teste_4_urlopen_chunks():
    """Abordagem com chunks e progresso"""
    print("\n=== TESTE 4: urlopen com chunks (simula progresso) ===")
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': '*/*',
            'Accept-Language': 'pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7',
            'Accept-Encoding': 'identity',
        }

        req = urllib.request.Request(DOWNLOAD_URL, headers=headers)

        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE

        temp_file = os.path.join(tempfile.gettempdir(), "teste4.exe")

        with urllib.request.urlopen(req, timeout=60, context=ctx) as response:
            total_size = response.headers.get('Content-Length')
            if total_size:
                total_size = int(total_size)
                print(f"Tamanho total: {total_size} bytes")

            downloaded = 0
            chunk_size = 8192

            with open(temp_file, 'wb') as f:
                while True:
                    chunk = response.read(chunk_size)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total_size:
                        percent = int(downloaded * 100 / total_size)
                        print(f"\rBaixando... {percent}%", end="", flush=True)

        print()
        size = os.path.getsize(temp_file)
        print(f"[OK] Baixou {size} bytes para {temp_file}")
        os.remove(temp_file)
        return True
    except Exception as e:
        print(f"[ERRO] {e}")
        return False


def teste_5_github_release():
    """Testar download de GitHub Releases (alternativa)"""
    print("\n=== TESTE 5: GitHub Releases (alternativa) ===")
    print("  Se OneDrive nao funcionar, usar GitHub Releases e a melhor opcao.")
    print("  Comando: gh release create v1.0.0 'dist/Validador SINT.exe'")
    print("  URL seria: https://github.com/USUARIO/REPO/releases/download/v1.0.0/Validador.SINT.exe")
    return None  # Nao testavel sem o release existir


if __name__ == "__main__":
    print("Testando diferentes abordagens de download...\n")
    print(f"URL: {DOWNLOAD_URL[:80]}...")
    print(f"Tamanho esperado: > {TAMANHO_MINIMO/1000000:.1f} MB")

    resultados = []

    resultados.append(("urlretrieve simples", teste_1_urlretrieve_simples()))
    resultados.append(("urlopen com headers", teste_2_urlopen_com_headers()))
    resultados.append(("opener com headers", teste_3_opener_com_headers()))
    resultados.append(("urlopen com chunks", teste_4_urlopen_chunks()))
    teste_5_github_release()

    print("\n" + "="*50)
    print("RESUMO:")
    print("="*50)
    for nome, sucesso in resultados:
        status = "OK" if sucesso else "FALHOU"
        print(f"  {nome}: {status}")

    print("\n" + "="*50)
    print("CONCLUSAO:")
    print("="*50)
    if not any(s for _, s in resultados):
        print("  OneDrive for Business requer autenticacao.")
        print("  ALTERNATIVAS:")
        print("  1. GitHub Releases (recomendado) - funciona sem autenticacao")
        print("  2. Google Drive com link publico")
        print("  3. Servidor proprio (ex: Azure Blob Storage)")
    else:
        print("  Pelo menos uma abordagem funcionou!")
