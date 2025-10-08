import os
import sys
import time
import re
import zipfile
import shutil
from lxml import etree

# ============================================================
# AJUSTE AUTOMÁTICO DE DIRETÓRIO (FUNCIONA NO .EXE)
# ============================================================
if getattr(sys, 'frozen', False):  # se rodar como .exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ============================================================
# CONFIGURAÇÕES AUTOMÁTICAS DE PASTAS
# ============================================================
DIR_ORIGEM = os.path.join(BASE_DIR, "PARA_EDICAO")
DIR_FINAL = os.path.join(BASE_DIR, "FINALIZADOS")
DIR_TEMP = os.path.join(BASE_DIR, "_TMP_EDICAO")

for pasta in [DIR_ORIGEM, DIR_FINAL, DIR_TEMP]:
    os.makedirs(pasta, exist_ok=True)

# CNPJs
CNPJ_CACAU = "33163908010561"
CNPJ_CHOCOLATE = "33163908008583"

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
def extrair_zips(origem, destino_tmp):
    """Extrai todos os arquivos ZIP encontrados em PARA_EDICAO"""
    zips = [f for f in os.listdir(origem) if f.lower().endswith(".zip")]
    extraidos = []

    for zip_name in zips:
        zip_path = os.path.join(origem, zip_name)
        nome_pasta = os.path.splitext(zip_name)[0]
        destino_pasta = os.path.join(destino_tmp, nome_pasta)
        os.makedirs(destino_pasta, exist_ok=True)

        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(destino_pasta)
            print(f"📦 Extraído: {zip_name}")
            extraidos.append((zip_name, zip_path, destino_pasta))
        except Exception as e:
            print(f"💥 Erro ao extrair {zip_name}: {e}")
    return extraidos


def recreate_zip(folder_path, zip_dest_path):
    """Compacta novamente uma pasta em um ZIP (mantendo PDFs e XMLs editados)"""
    with zipfile.ZipFile(zip_dest_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(folder_path):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, folder_path)
                zf.write(full_path, rel_path)
    print(f"🗜️ Novo ZIP: {os.path.basename(zip_dest_path)}")


def obter_info_xml(file_path):
    """Extrai informações principais do XML (UF, nCT, CNPJ Tomador)"""
    try:
        tree = etree.parse(file_path)
        ns = {"cte": "http://www.portalfiscal.inf.br/cte"}
        uf = tree.findtext(".//cte:UFEnv", namespaces=ns)
        nct = tree.findtext(".//cte:nCT", namespaces=ns)
        cnpj_rem = tree.findtext(".//cte:rem/cte:CNPJ", namespaces=ns)
        if cnpj_rem:
            cnpj_rem = re.sub(r"\D", "", cnpj_rem)
        return uf, nct, cnpj_rem
    except Exception:
        return None, None, None


def alterar_po(file_path, novo_po):
    """Substitui ou adiciona o PO em qualquer parte do XML"""
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            xml_text = f.read()

        # Captura todos os POs antigos (mesmo que repetidos)
        encontrados = re.findall(r"4504\d{6,}/\d{5}", xml_text)
        antigo = encontrados[0] if encontrados else "(não encontrado)"

        # Substitui todos os padrões existentes por novo PO
        xml_editado = re.sub(r"4504\d{6,}/\d{5}", novo_po, xml_text)

        # Se não houver nenhum PO anterior, ainda insere o novo (mantendo XML íntegro)
        if xml_editado == xml_text:
            xml_editado = xml_text.replace("</CTe>", f"<xObs>{novo_po}</xObs></CTe>")

        # Regrava o XML editado
        with open(file_path, "w", encoding="utf-8") as f:
            f.write(xml_editado)

        return True, antigo

    except Exception as e:
        print(f"⚠️ Erro ao editar {file_path}: {e}")
        return False, None


# ============================================================
# GRADE DE POs
# ============================================================
PO_GRID_CACAU = {
    "SP": [
        ("4504819456/00010", "FRETE DE VENDAS"),
        ("4504819466/00010", "Transferencia Cacau - Omegax CROSS/ARMAZ."),
        ("4504819456/00020", "FRETE CUSTO EXT - HOSP PERN"),
    ],
    "MG": [
        ("4504819472/00010", "FRETE DE VENDAS"),
        ("4504819478/00010", "Transferencia Cacau - Omegax CROSS/ARMAZ."),
        ("4504819478/00020", "FRETE CUSTO EXT - HOSP PERN"),
    ],
}

PO_GRID_CHOCOLATE = {
    "SP": [
        ("4504820478/00010", "FRETE DE VENDAS"),
        ("4504820597/00010", "FRETE CUSTO EXT-  HOSP PER"),
    ],
    "MG": [
        ("4504820480/00010", "FRETE DE VENDAS"),
        ("4504820481/00010", "Transferencia PA - Omega X Cross/Armazenagem"),
        ("4504820600/00010", "FRETE CUSTO EXT-  HOSP PER"),
    ],
}


# ============================================================
# PROCESSAMENTO PRINCIPAL
# ============================================================
def main():
    print("\n🔍 Iniciando varredura de arquivos XML...\n")
    inicio = time.time()

    pastas_zipadas = extrair_zips(DIR_ORIGEM, DIR_TEMP)

    # XMLs soltos (não zipados)
    xml_files_soltos = [
        os.path.join(DIR_ORIGEM, f)
        for f in os.listdir(DIR_ORIGEM)
        if f.lower().endswith(".xml")
    ]

    xml_files = xml_files_soltos[:]
    for _, _, pasta in pastas_zipadas:
        for root, _, files in os.walk(pasta):
            for file in files:
                if file.lower().endswith(".xml"):
                    xml_files.append(os.path.join(root, file))

    if not xml_files:
        print("🚫 Nenhum arquivo XML encontrado (nem solto, nem zipado).")
        input("\nPressione ENTER para sair...")
        return

    arquivos_info = []
    for caminho in xml_files:
        uf, nct, cnpj = obter_info_xml(caminho)
        if uf:
            arquivos_info.append({"arquivo": caminho, "UF": uf, "nCT": nct, "CNPJ": cnpj})

    total_mg = len([a for a in arquivos_info if a["UF"] == "MG"])
    total_sp = len([a for a in arquivos_info if a["UF"] == "SP"])
    uf_dominante = "MG" if total_mg >= total_sp else "SP"

    cnpjs = [a["CNPJ"] for a in arquivos_info if a["CNPJ"]]
    cnpj_dominante = max(set(cnpjs), key=cnpjs.count)
    tomador = "CACAU" if cnpj_dominante == CNPJ_CACAU else "CHOCOLATE"

    print("📊 Resumo de detecção:")
    print(f"   ➤ MG: {total_mg} arquivo(s)")
    print(f"   ➤ SP: {total_sp} arquivo(s)")
    print(f"   ➤ Tomador: {tomador} ({cnpj_dominante})")
    print(f"\n➡️ UF dominante: {uf_dominante}\n")

    PO_GRID = PO_GRID_CACAU if tomador == "CACAU" else PO_GRID_CHOCOLATE
    print(f"📋 Opções de PO para {tomador} ({uf_dominante}):")
    for i, (po, tipo) in enumerate(PO_GRID[uf_dominante], 1):
        print(f"  {i} → {po} ({tipo})")

    opcao = int(input(f"\nDigite o Nº do PO desejado para {tomador}: ").strip())
    novo_po, tipo_po = PO_GRID[uf_dominante][opcao - 1]
    print(f"\n✅ PO selecionado: {novo_po} ({tipo_po})\n")

    alterados, ignorados = 0, 0
    detalhes_alterados, detalhes_ignorados = [], []

    total = len(arquivos_info)
    for idx, arq in enumerate(arquivos_info, 1):
        uf = arq["UF"]
        nct = arq["nCT"]
        caminho = arq["arquivo"]
        print(f"🧩 ({idx}/{total}) nCT {nct}", end="\r")

        if uf != uf_dominante:
            ignorados += 1
            detalhes_ignorados.append(f"nCT {nct}: Ignorado ({uf})")
            continue

        ok, antigo = alterar_po(caminho, novo_po)
        if ok:
            alterados += 1
            detalhes_alterados.append(f"nCT {nct}: Alterado de {antigo} → {novo_po}")
        else:
            ignorados += 1
            detalhes_ignorados.append(f"nCT {nct}: Sem alteração")

    # Recria os ZIPs
    for nome_zip, zip_original, pasta_tmp in pastas_zipadas:
        zip_dest = os.path.join(DIR_FINAL, nome_zip)
        recreate_zip(pasta_tmp, zip_dest)
        try:
            os.remove(zip_original)
        except:
            pass

    # Move XMLs soltos alterados para FINALIZADOS
    for xml in xml_files_soltos:
        nome = os.path.basename(xml)
        shutil.move(xml, os.path.join(DIR_FINAL, nome))

    shutil.rmtree(DIR_TEMP, ignore_errors=True)

    log_path = os.path.join(DIR_FINAL, "LOG_EDITOR_PO.txt")
    with open(log_path, "w", encoding="utf-8") as log:
        log.write(f"Tomador: {tomador} ({cnpj_dominante})\n")
        log.write(f"UF dominante: {uf_dominante}\n")
        log.write(f"PO aplicado: {novo_po} - {tipo_po}\n\n")
        log.write(f"✅ Alterados: {alterados}\n")
        log.write(f"🚫 Ignorados: {ignorados}\n\n")
        if detalhes_alterados:
            log.write("=== ARQUIVOS ALTERADOS ===\n")
            log.write("\n".join(detalhes_alterados) + "\n\n")
        if detalhes_ignorados:
            log.write("=== ARQUIVOS IGNORADOS ===\n")
            log.write("\n".join(detalhes_ignorados) + "\n")

    print("\n📊 RESUMO FINAL")
    print(f"✅ Alterados: {alterados}")
    print(f"🚫 Ignorados: {ignorados}")
    print(f"🏁 Arquivos salvos em: {DIR_FINAL}")
    print(f"🗂️ Log salvo em: {log_path}")

    input("\nPressione ENTER para sair...")


if __name__ == "__main__":
    main()



# Para alterar o arquivo executável, execute o seguinte comando no prompt de comando: pyinstaller --onefile automa_editor_po_barry_portatil.py