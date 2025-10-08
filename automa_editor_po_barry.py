import os
import time
import re
import zipfile
import shutil
from lxml import etree

# ============================================================
# CONFIGURAÇÕES GERAIS
# ============================================================
DIR_ORIGEM = r"C:\Users\Sidenei Silva\Desktop\PROJETO_PYTHON\EDITOR_PO_XML\PARA_EDICAO"
DIR_FINAL = r"C:\Users\Sidenei Silva\Desktop\PROJETO_PYTHON\EDITOR_PO_XML\FINALIZADOS"
DIR_TEMP = os.path.join(DIR_FINAL, "_EXTRAIDOS_TMP")

os.makedirs(DIR_FINAL, exist_ok=True)
os.makedirs(DIR_TEMP, exist_ok=True)

# CNPJs dos Tomadores
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
            print(f"📦 Extraído: {zip_name} → {destino_pasta}")
            extraidos.append(destino_pasta)
        except Exception as e:
            print(f"💥 Erro ao extrair {zip_name}: {e}")
    return extraidos


def modify_text_value(tree, old_value=None, new_value=None):
    """Substitui valores de PO em qualquer tag de texto"""
    changed = False
    for elem in tree.xpath("//*[text()]"):
        text = elem.text
        if text:
            if old_value and old_value in text:
                elem.text = text.replace(old_value, new_value)
                changed = True
            elif "4504" in text and "/000" in text:
                elem.text = new_value
                changed = True
    return changed


def modify_xTexto_value(tree, old_value=None, new_value=None):
    """Substitui valores dentro de <xTexto>"""
    changed = False
    for elem in tree.findall(".//xTexto"):
        if elem.text:
            if old_value and old_value in elem.text:
                elem.text = elem.text.replace(old_value, new_value)
                changed = True
            elif "4504" in elem.text and "/000" in elem.text:
                elem.text = new_value
                changed = True
    return changed


def obter_info_xml(file_path):
    """Lê UF, Município, nCT e CNPJ do Tomador (<rem><CNPJ>)"""
    try:
        tree = etree.parse(file_path)
        ns = {"cte": "http://www.portalfiscal.inf.br/cte"}

        uf = tree.findtext(".//cte:UFEnv", namespaces=ns)
        mun = tree.findtext(".//cte:xMunEnv", namespaces=ns)
        nct = tree.findtext(".//cte:nCT", namespaces=ns)
        cnpj_rem = tree.findtext(".//cte:rem/cte:CNPJ", namespaces=ns)
        if cnpj_rem:
            cnpj_rem = re.sub(r"\D", "", cnpj_rem)
        return uf, mun, nct, cnpj_rem
    except Exception as e:
        print(f"⚠️ Erro ao ler {file_path}: {e}")
        return None, None, None, None


def alterar_po(file_path, novo_po, destino_final):
    """Edita o PO no XML e salva no destino mantendo a estrutura"""
    try:
        tree = etree.parse(file_path)
        ns = {"cte": "http://www.portalfiscal.inf.br/cte"}

        valores_antigos = []
        for elem in tree.xpath(".//cte:xTexto | .//cte:xObs", namespaces=ns):
            if elem.text:
                encontrados = re.findall(r"4504\d{6,}/\d{5}", elem.text)
                if encontrados:
                    valores_antigos.extend(encontrados)
        valores_antigos = list(set(valores_antigos))

        alterado = modify_text_value(tree, None, novo_po) or modify_xTexto_value(tree, None, novo_po)
        if alterado:
            os.makedirs(os.path.dirname(destino_final), exist_ok=True)
            tree.write(destino_final, pretty_print=True, xml_declaration=True, encoding="UTF-8")
            antigo = valores_antigos[0] if valores_antigos else "(não encontrado)"
            return True, antigo
        else:
            return False, None
    except Exception as e:
        print(f"💥 Erro ao alterar {file_path}: {e}")
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

    pastas_extraidas = extrair_zips(DIR_ORIGEM, DIR_TEMP)
    caminhos_base = [DIR_ORIGEM] + pastas_extraidas

    xml_files = []
    for base in caminhos_base:
        for root_dir, _, files in os.walk(base):
            for file in files:
                if file.lower().endswith(".xml"):
                    xml_files.append(os.path.join(root_dir, file))

    if not xml_files:
        print("🚫 Nenhum arquivo XML encontrado (nem em ZIPs ou subpastas).")
        input("\nPressione ENTER para sair...")
        return

    arquivos_info = []
    for caminho in xml_files:
        uf, mun, nct, cnpj_tomador = obter_info_xml(caminho)
        if uf:
            arquivos_info.append({"arquivo": caminho, "UF": uf, "MUN": mun, "nCT": nct, "CNPJ": cnpj_tomador})

    total_mg = len([a for a in arquivos_info if a["UF"] == "MG"])
    total_sp = len([a for a in arquivos_info if a["UF"] == "SP"])
    uf_dominante = "MG" if total_mg >= total_sp else "SP"

    cnpjs = [a["CNPJ"] for a in arquivos_info if a["CNPJ"]]
    if not cnpjs:
        print("⚠️ Nenhum CNPJ de tomador localizado.")
        input("\nPressione ENTER para sair...")
        return
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

    alterados, nao_alterados, ignorados = 0, 0, 0
    detalhes_alterados, detalhes_nao_alterados, detalhes_ignorados = [], [], []

    total_arquivos = len(arquivos_info)
    ajustados = 0

    for idx, arq in enumerate(arquivos_info, 1):
        uf = arq["UF"]
        nct = arq["nCT"]
        caminho = arq["arquivo"]
        rel_path = os.path.basename(caminho)
        destino = os.path.join(DIR_FINAL, rel_path)

        print(f"🧩 ({ajustados}/{total_arquivos}) nCT Ajustados", end="\r")

        if uf != uf_dominante:
            ignorados += 1
            detalhes_ignorados.append(f"nCT {nct}: Ignorado ({uf})")
            continue

        ok, antigo = alterar_po(caminho, novo_po, destino)
        if ok:
            alterados += 1
            ajustados += 1
            detalhes_alterados.append(f"nCT {nct}: Alterado de {antigo} para → {novo_po}")
            try:
                os.remove(caminho)
            except Exception:
                pass
        else:
            nao_alterados += 1
            ajustados += 1
            detalhes_nao_alterados.append(f"nCT {nct}: Sem alteração")

    fim = time.time()
    duracao = round(fim - inicio, 2)

    # ===== GERAR LOG FORMATADO =====
    log_path = os.path.join(DIR_FINAL, "LOG_EDITOR_PO.txt")
    with open(log_path, "w", encoding="utf-8") as log:
        log.write("======================================================================\n")
        log.write("🧾 LOG DE EDIÇÃO DE PO XML\n")
        log.write("======================================================================\n\n")
        log.write(f"Tomador: {tomador} ({cnpj_dominante})\n")
        log.write(f"UF dominante: {uf_dominante}\n")
        log.write(f"PO aplicado: {novo_po} - {tipo_po}\n\n")
        log.write(f"✅ Alterados: {alterados}\n")
        log.write(f"⚠️ Não alterados: {nao_alterados}\n")
        log.write(f"🚫 Ignorados: {ignorados}\n\n")
        log.write("----------------------------------------------------------------------\n")

        if detalhes_alterados:
            log.write("=== ARQUIVOS ALTERADOS ===\n")
            log.write("\n".join(detalhes_alterados) + "\n\n")

        if detalhes_nao_alterados:
            log.write("=== ARQUIVOS NÃO ALTERADOS ===\n")
            log.write("\n".join(detalhes_nao_alterados) + "\n\n")

        if detalhes_ignorados:
            log.write("=== ARQUIVOS IGNORADOS ===\n")
            log.write("\n".join(detalhes_ignorados) + "\n\n")

        log.write("======================================================================\n")
        log.write(f"Tempo total de execução: {duracao} segundos\n")
        log.write(f"Arquivos salvos em: {DIR_FINAL}\n")
        log.write("======================================================================\n")

    # ===== RESUMO FINAL NO TERMINAL =====
    print("\n" + "=" * 70)
    print("📊 RESUMO FINAL")
    print("=" * 70)
    print(f"🧾 Tomador: {tomador} ({cnpj_dominante})")
    print(f"📍 UF dominante: {uf_dominante}")
    print(f"🧾 PO aplicado: {novo_po} ({tipo_po})\n")
    print(f"✅ Alterados: {alterados}")
    print(f"⚠️ Não alterados: {nao_alterados}")
    print(f"🚫 Ignorados: {ignorados}")
    print("=" * 70)
    print(f"🏁 Arquivos alterados salvos em: {DIR_FINAL}")
    print(f"🗂️ Log salvo em: {log_path}")
   
    try:
        if os.path.exists(DIR_TEMP):
            shutil.rmtree(DIR_TEMP)
    except Exception:
        pass

    input("\nPressione ENTER para sair...")


# ============================================================
if __name__ == "__main__":
    main()


# Para alterar o arquivo executável, execute o seguinte comando no prompt de comando: pyinstaller --onefile automa_editor_po_barry.py

