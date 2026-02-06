import win32com.client as win32
from pathlib import Path
import re
import unicodedata
from datetime import datetime
import csv
import sys

CONTA_EMAIL = ""
PASTA_DESTINO = Path.cwd() / "Anexos"
EXTENSOES_IMAGEM = (".jpg", ".jpeg", ".png", ".bmp")
ARQUIVO_REGISTRO_CSV = PASTA_DESTINO / "registro_crachas.csv"

def limpar_texto(texto):
    texto = str(texto or "")
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    # Note: O comentário original aqui foi removido conforme solicitado
    texto = re.sub(r'[<>:"/\\|?*]', "", texto)
    return texto.strip() or "SemNome"

def extrair_secretaria(corpo):
    if not corpo:
        return "SemSecretaria"

    linhas = [ln.strip() for ln in re.split(r'\r?\n', corpo) if ln.strip()]
    secretaria_patterns = [
        r"(?:Secretaria/Departamento|Departamento|Secretaria)\s*[:\-]\s*(.+)",
    ]

    for ln in linhas:
        ln_clean = ln.strip()
        ln_clean = re.sub(r'^\[|\]$', '', ln_clean).strip()

        for pat in secretaria_patterns:
            m = re.search(pat, ln_clean, re.IGNORECASE)
            if m:
                candidate = m.group(1).strip()
                candidate = re.split(r'\bAtenciosamente\b|,|;|-{2,}', candidate, flags=re.IGNORECASE)[0].strip()
                return limpar_texto(candidate) if candidate else "SemSecretaria"
    return "SemSecretaria"

def extrair_nome_matricula(corpo):
    if not corpo:
        return "SemNome", "SemMatricula"

    linhas = [ln.strip() for ln in re.split(r'\r?\n', corpo) if ln.strip()]

    nome = None
    matricula = None

    nome_patterns = [
        r"(?:Nome completo|Nome do colaborador|Nome)\s*[:\-]\s*(.+)",
    ]
    matricula_patterns = [
        r"Mat[ríi]cula\s*[:\-]\s*([\w\/\-\.]+)",
        r"Registro\s*[:\-]\s*([\w\/\-\.]+)"
    ]

    for ln in linhas:
        ln_clean = ln.strip()
        ln_clean = re.sub(r'^\[|\]$', '', ln_clean).strip()

        if nome is None:
            for pat in nome_patterns:
                m = re.search(pat, ln_clean, re.IGNORECASE)
                if m:
                    candidate = m.group(1).strip()
                    if candidate and 'nome' not in candidate.lower():
                        nome = candidate
                        break

        if matricula is None:
            for pat in matricula_patterns:
                m = re.search(pat, ln_clean, re.IGNORECASE)
                if m:
                    matricula = m.group(1).strip().replace(" ", "")
                    break
        if nome and matricula:
            break

    if matricula is None:
        m = re.search(r"\b(\d{4,12})\b", corpo)
        if m:
            matricula = m.group(1)

    nome = limpar_texto(nome) if nome else "SemNome"
    matricula = limpar_texto(matricula) if matricula else "SemMatricula"

    return nome, matricula

def salvar_anexo(anexo, destino, matricula):
    nome_arquivo = anexo.FileName
    ext = Path(nome_arquivo).suffix.lower()

    if ext not in EXTENSOES_IMAGEM:
        return None

    matricula_sanitizada = matricula.replace(" ", "_")
    novo_nome = f"{matricula_sanitizada}{ext}"
    caminho_final = destino / novo_nome

    if caminho_final.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        novo_nome = f"{matricula_sanitizada}_{timestamp}{ext}"
        caminho_final = destino / novo_nome

    try:
        anexo.SaveAsFile(str(caminho_final))
        print(f"✅ Foto salva: {caminho_final}")
        return str(caminho_final)
    except Exception as e:
        print(f"❌ Erro ao salvar anexo {nome_arquivo}: {e}")
        return None

def registrar_dados(nome, matricula, secretaria, caminho_arquivo_foto):
    cabecalho = ["Nome", "Matricula", "Secretaria/Departamento", "Foto"]
    escrever_cabecalho = not ARQUIVO_REGISTRO_CSV.exists()

    csv_encoding = 'utf-8' if sys.platform != 'win32' else 'utf-8-sig'

    try:
        with open(ARQUIVO_REGISTRO_CSV, mode='a', newline='', encoding=csv_encoding) as file:
            writer = csv.writer(file, delimiter=';')
            if escrever_cabecalho:
                writer.writerow(cabecalho)
            writer.writerow([nome, matricula, secretaria, caminho_arquivo_foto])
        print(f"📝 Dados registrados no CSV: {ARQUIVO_REGISTRO_CSV.name}")
    except Exception as e:
        print(f"❌ Erro ao registrar no CSV: {e}")

def processar_emails():
    print("Conectando ao Outlook...")
    try:
        outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as e:
        print(f"❌ Erro ao conectar ao Outlook. Certifique-se de que está aberto. Erro: {e}")
        return

    conta = None
    for i in range(1, outlook.Folders.Count + 1):
        if outlook.Folders.Item(i).Name.lower() == CONTA_EMAIL.lower():
            conta = outlook.Folders.Item(i)
            break

    if not conta:
        raise Exception(f"Conta '{CONTA_EMAIL}' não encontrada no Outlook.")

    try:
        caixa_entrada = conta.Folders["Caixa de Entrada"]
    except Exception:
        caixa_entrada = conta.Folders["Inbox"]

    mensagens = caixa_entrada.Items
    mensagens.Sort("[ReceivedTime]", True)

    contador_lidos = 0
    contador_processados = 0
    PASTA_DESTINO.mkdir(parents=True, exist_ok=True)

    print("Iniciando varredura de e-mails não lidos...")
    for m in mensagens:
        if not getattr(m, "UnRead", False):
            continue

        contador_lidos += 1
        corpo = str(getattr(m, "Body", "") or "")
        nome, matricula = extrair_nome_matricula(corpo)

        print(f"\n📩 Assunto: {getattr(m, 'Subject', '<sem assunto>')}")
        print(f"👤 Nome: {nome} | Matrícula: {matricula}")

        if matricula == "SemMatricula":
            print("🛑 E-mail ignorado (sem matrícula).")
            continue

        secretaria = extrair_secretaria(corpo)
        print(f"🏢 Secretaria: {secretaria}")

        foto_salva = False
        nome_arquivo_foto = "NaoEncontrada"

        if getattr(m, "Attachments", None) and m.Attachments.Count > 0:
            for a in m.Attachments:
                caminho_arq_salvo = salvar_anexo(a, PASTA_DESTINO, matricula)
                if caminho_arq_salvo:
                    nome_arquivo_foto = caminho_arq_salvo
                    foto_salva = True
                    break

        if not foto_salva:
            print("⚠️ Nenhum anexo de imagem encontrado.")

        registrar_dados(nome, matricula, secretaria, nome_arquivo_foto)

        try:
            m.UnRead = False
            contador_processados += 1
            print("📧 E-mail marcado como LIDO.")
        except Exception as e:
            print(f"❌ Erro ao marcar e-mail como lido: {e}")

    print("\n--- Processo Concluído ---")
    print(f"E-mails não lidos: {contador_lidos}")
    print(f"E-mails processados: {contador_processados}")

if __name__ == "__main__":
    print("Iniciando leitura de e-mails para confecção de crachá...")
    print(f"📁 Pasta destino: {PASTA_DESTINO}")
    print(f"🗂️ CSV: {ARQUIVO_REGISTRO_CSV}\n")

    while True:
        try:
            processar_emails()
        except Exception as e:
            print(f"\n🚨 ERRO FATAL: {e}")

        resp = input("\nDeseja rodar novamente? (s/n): ").lower()
        if resp != "s":
            print("Encerrando execução.")
            break