# ============================================================
# EDITOR_AUTOM_PO – RESUMO FUNCIONAL
# ============================================================
# Objetivo:
# Automatizar a validação e correção de PO (Pedido SAP) em CT-e (XML),
# com base na pasta onde o arquivo é colocado pelo usuário.
#
# Funcionamento geral:
# - O sistema roda automaticamente (manual, Agendador do Windows ou .exe).
# - Trabalha sempre na Área de Trabalho do usuário, na pasta:
#   EDITOR_BO_BARRY
#
# Estrutura de pastas:
# - PARA_EDICAO/
#     ├─ FRETE
#     ├─ TRANSFERENCIA
#     └─ CUSTO
# - SAIDA_FINAL/
#     ├─ FRETE
#     ├─ TRANSFERENCIA
#     └─ CUSTO
# - LOG/
#
# Regras de processamento:
# - O usuário define o TIPO do processo apenas pela pasta (FRETE / TRANSFERENCIA / CUSTO).
# - O sistema NÃO confia no PO existente no XML.
# - O PO é SEMPRE forçado conforme:
#     Tipo (pasta) + Tomador (CNPJ no XML) + UF (UFEnv no XML).
#
# Tomadores reconhecidos:
# - CACAU      → CNPJ 33163908010561
# - CHOCOLATE  → CNPJ 33163908008583
#
# Tipos de arquivo aceitos:
# - XML solto:
#     • PO é validado/alterado
#     • XML é movido para a pasta de saída correspondente
#     • PDF (DACTE), se existir, é movido junto
#
# - ZIP:
#     • ZIP é extraído em pasta temporária
#     • Todos os XMLs internos são processados
#     • PDFs e demais arquivos são preservados
#     • ZIP é recomposto com o mesmo nome
#     • ZIP final é movido para a pasta de saída correspondente
#
# Logs:
# - LOG_EDICAO_PO.txt
#     • XML solto: 1 linha por XML (PO antes e depois)
#     • ZIP: 1 linha resumida por ZIP
#       (total de XMLs + distribuição de PO antes e depois)
#
# - LOG_ERRO.txt
#     • Gerado somente se ocorrer falha
#     • Arquivos com erro permanecem na pasta de entrada
#
# Comportamento esperado:
# - Em sucesso: pasta de entrada fica vazia
# - Em erro: nada é movido e o erro é registrado
#
# Observação importante:
# - O mesmo XML pode ser processado novamente em outra pasta.
# - O PO final sempre refletirá o TIPO da pasta atual.
# ============================================================




import os
import re
import shutil
import zipfile
import tempfile
from datetime import datetime
from collections import Counter
from lxml import etree


# ============================================================
# LOCALIZAÇÃO (ÁREA DE TRABALHO)
# ============================================================
def get_desktop_path():
    return os.path.join(os.path.expanduser("~"), "Desktop")


# ============================================================
# PASTAS BASE
# ============================================================
def definir_pastas_base():
    desktop = get_desktop_path()
    base = os.path.join(desktop, "EDITOR_BO_BARRY")

    return {
        "BASE": base,

        # Entrada
        "ENTRADA_FRETE": os.path.join(base, "PARA_EDICAO", "FRETE"),
        "ENTRADA_TRANSFERENCIA": os.path.join(base, "PARA_EDICAO", "TRANSFERENCIA"),
        "ENTRADA_CUSTO": os.path.join(base, "PARA_EDICAO", "CUSTO"),

        # Saída
        "SAIDA_FRETE": os.path.join(base, "SAIDA_FINAL", "FRETE"),
        "SAIDA_TRANSFERENCIA": os.path.join(base, "SAIDA_FINAL", "TRANSFERENCIA"),
        "SAIDA_CUSTO": os.path.join(base, "SAIDA_FINAL", "CUSTO"),

        # Logs
        "LOG": os.path.join(base, "LOG"),
    }


def garantir_pastas(pastas):
    for caminho in pastas.values():
        os.makedirs(caminho, exist_ok=True)


# ============================================================
# LOGS
# ============================================================
def registrar_log_xml(pastas, tipo, arquivo, po_antigo, po_novo):
    log = os.path.join(pastas["LOG"], "LOG_EDICAO_PO.txt")
    with open(log, "a", encoding="utf-8") as f:
        f.write(
            f"{datetime.now():%Y-%m-%d %H:%M:%S} | "
            f"{tipo} | {arquivo} | "
            f"PO_ANTES={po_antigo} | PO_DEPOIS={po_novo}\n"
        )


def registrar_log_zip_resumido(pastas, tipo, zip_nome, total, antes, depois):
    def fmt(counter):
        return ", ".join([f"{qtd}x {po}" for po, qtd in counter.items()])

    log = os.path.join(pastas["LOG"], "LOG_EDICAO_PO.txt")
    with open(log, "a", encoding="utf-8") as f:
        f.write(
            f"{datetime.now():%Y-%m-%d %H:%M:%S} | "
            f"{tipo} | {zip_nome} | "
            f"TOTAL_XML={total} | "
            f"PO_ANTES=[{fmt(antes)}] | "
            f"PO_DEPOIS=[{fmt(depois)}]\n"
        )


def registrar_erro(pastas, tipo, arquivo, erro):
    log = os.path.join(pastas["LOG"], "LOG_ERRO.txt")
    with open(log, "a", encoding="utf-8") as f:
        f.write(
            f"{datetime.now():%Y-%m-%d %H:%M:%S} | "
            f"{tipo} | {arquivo} | ERRO={erro}\n"
        )


# ============================================================
# REGRAS DE PO
# ============================================================
CNPJ_CACAU = "33163908010561"
CNPJ_CHOCOLATE = "33163908008583"

PO_RULES = {
    "CUSTO": {
        "CACAU": {"SP": "4504819456/00020", "MG": "4504819478/00020"},
        "CHOCOLATE": {"SP": "4504820597/00010", "MG": "4504820600/00010"},
    },
    "FRETE": {
        "CACAU": {"SP": "4504819456/00010", "MG": "4504819472/00010"},
        "CHOCOLATE": {"SP": "4504820478/00010", "MG": "4504820480/00010"},
    },
    "TRANSFERENCIA": {
        "CACAU": {"SP": "4504819466/00010", "MG": "4504819478/00010"},
        "CHOCOLATE": {"SP": "4504820481/00010", "MG": "4504820481/00010"},
    },
}


# ============================================================
# XML
# ============================================================
def extrair_info_xml(xml_path):
    tree = etree.parse(xml_path)
    ns = {"cte": "http://www.portalfiscal.inf.br/cte"}

    uf = tree.findtext(".//cte:UFEnv", namespaces=ns)
    cnpj = tree.findtext(".//cte:rem/cte:CNPJ", namespaces=ns)

    if not uf or not cnpj:
        raise ValueError("UF ou CNPJ não encontrados")

    cnpj = re.sub(r"\D", "", cnpj)
    tomador = "CACAU" if cnpj == CNPJ_CACAU else "CHOCOLATE"

    return uf, tomador


def alterar_po_xml(xml_path, novo_po):
    with open(xml_path, "r", encoding="utf-8") as f:
        xml = f.read()

    encontrados = re.findall(r"4504\d{6,}/\d{5}", xml)
    po_antigo = encontrados[0] if encontrados else "NAO_ENCONTRADO"

    xml_editado = re.sub(r"4504\d{6,}/\d{5}", novo_po, xml)

    if xml_editado == xml:
        xml_editado = xml.replace("</CTe>", f"<xObs>{novo_po}</xObs></CTe>")

    with open(xml_path, "w", encoding="utf-8") as f:
        f.write(xml_editado)

    return po_antigo, novo_po


# ============================================================
# PROCESSAMENTO XML SOLTO
# ============================================================
def processar_xml_individual(pastas, tipo, xml_path):
    try:
        uf, tomador = extrair_info_xml(xml_path)
        novo_po = PO_RULES[tipo][tomador][uf]

        po_antigo, po_novo = alterar_po_xml(xml_path, novo_po)

        registrar_log_xml(
            pastas, tipo, os.path.basename(xml_path), po_antigo, po_novo
        )

        destino = os.path.join(pastas[f"SAIDA_{tipo}"], os.path.basename(xml_path))
        shutil.move(xml_path, destino)

        # PDF associado
        base = os.path.splitext(os.path.basename(xml_path))[0]
        pdf = os.path.join(os.path.dirname(xml_path), f"{base}.pdf")
        if os.path.exists(pdf):
            shutil.move(pdf, os.path.join(pastas[f"SAIDA_{tipo}"], os.path.basename(pdf)))

    except Exception as e:
        registrar_erro(pastas, tipo, os.path.basename(xml_path), str(e))


# ============================================================
# PROCESSAMENTO ZIP
# ============================================================
def processar_zip(pastas, tipo, zip_path):
    zip_nome = os.path.basename(zip_path)

    total = 0
    po_antes = Counter()
    po_depois = Counter()

    try:
        tmp_dir = tempfile.mkdtemp(prefix="ZIP_", dir=pastas["BASE"])

        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(tmp_dir)

        for root, _, files in os.walk(tmp_dir):
            for nome in files:
                if nome.lower().endswith(".xml"):
                    xml_path = os.path.join(root, nome)

                    uf, tomador = extrair_info_xml(xml_path)
                    novo_po = PO_RULES[tipo][tomador][uf]

                    po_antigo, po_novo = alterar_po_xml(xml_path, novo_po)

                    po_antes[po_antigo] += 1
                    po_depois[po_novo] += 1
                    total += 1

        destino_zip = os.path.join(pastas[f"SAIDA_{tipo}"], zip_nome)
        with zipfile.ZipFile(destino_zip, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, _, files in os.walk(tmp_dir):
                for file in files:
                    full = os.path.join(root, file)
                    rel = os.path.relpath(full, tmp_dir)
                    zf.write(full, rel)

        registrar_log_zip_resumido(
            pastas, tipo, zip_nome, total, po_antes, po_depois
        )

        os.remove(zip_path)
        shutil.rmtree(tmp_dir, ignore_errors=True)

    except Exception as e:
        registrar_erro(pastas, tipo, zip_nome, str(e))


# ============================================================
# EXECUÇÃO
# ============================================================
def executar(pastas):
    entradas = {
        "FRETE": pastas["ENTRADA_FRETE"],
        "TRANSFERENCIA": pastas["ENTRADA_TRANSFERENCIA"],
        "CUSTO": pastas["ENTRADA_CUSTO"],
    }

    for tipo, pasta in entradas.items():
        for nome in os.listdir(pasta):
            caminho = os.path.join(pasta, nome)

            if not os.path.isfile(caminho):
                continue

            if nome.lower().endswith(".xml"):
                processar_xml_individual(pastas, tipo, caminho)

            elif nome.lower().endswith(".zip"):
                processar_zip(pastas, tipo, caminho)


# ============================================================
# MAIN
# ============================================================
def main():
    pastas = definir_pastas_base()
    garantir_pastas(pastas)
    executar(pastas)


if __name__ == "__main__":
    main()


# ============================================================  


# Para alterar o arquivo executável, execute o seguinte comando no prompt de comando: pyinstaller --onefile EDITOR_AUTOM_PO.py

