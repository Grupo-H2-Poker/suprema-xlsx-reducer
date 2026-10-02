#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Redutor de planilha Suprema 113 (Grand Union)
=============================================
Gera uma versão enxuta do relatório da Suprema, contendo apenas o que o
importador (process_suprema_113.py) usa:

- Mantém só as abas: Grand Union Data, Grand Union Member Resume,
  Grand Union Game Statistics e Grand Union Game Detail
- Na aba Grand Union Game Detail mantém apenas as linhas de jogadores dos
  clubes H2 (todos os blocos de jogo são mantidos, mesmo sem jogadores H2)
- As demais abas são copiadas como estão, na mesma posição de linha/coluna
- Remove formatação, células vazias e linhas em branco

Uso:
    python reduz_planilha.py
        Processa todos os .xlsx da pasta pendentes/. A planilha reduzida é
        gravada em processados/ com o mesmo nome e o original é movido para
        processados/originais/. Se der erro, o arquivo fica em pendentes/.

    python reduz_planilha.py ENTRADA.xlsx [-o SAIDA.xlsx]
        Reduz um arquivo específico (padrão de saída: <nome>_reduzido.xlsx).
"""

import argparse
import os
import re
import shutil
import sys
import time
import traceback
import zipfile
import xml.etree.ElementTree as ET

import xlsxwriter

# ====================================
# CONFIGURAÇÃO
# ====================================
ABAS = [
    "Grand Union Data",
    "Grand Union Member Resume",
    "Grand Union Game Statistics",
    "Grand Union Game Detail",
]
ABA_DETALHE = "Grand Union Game Detail"

BASE_PATH = os.path.dirname(os.path.abspath(__file__))
PENDENTES_PATH = os.path.join(BASE_PATH, "pendentes")
PROCESSADOS_PATH = os.path.join(BASE_PATH, "processados")
ORIGINAIS_PATH = os.path.join(PROCESSADOS_PATH, "originais")

# Mesmo critério do PHP clubesH2(): nome LIKE '%h2 %' OR nome LIKE 'h2poker%'
# Mantém "H2poker 1", "H2poker | ITM", "H2 Club Curitiba"; descarta "XYZ ♤ H20"
PATTERN_CLUBE_H2 = re.compile(r"^h2poker|h2 ", re.IGNORECASE)

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RNS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PATTERN_REF = re.compile(r"([A-Z]+)(\d+)")


def log(msg):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}", flush=True)


# ====================================
# LEITURA EM STREAMING DO XLSX
# ====================================
class LeitorXlsx:
    """Lê as abas linha a linha direto do XML, sem carregar a planilha na memória"""

    def __init__(self, caminho):
        self.zip = zipfile.ZipFile(caminho)
        self.abas = self._mapear_abas()
        self.shared_strings = self._carregar_shared_strings()

    def _mapear_abas(self):
        wb = ET.fromstring(self.zip.read("xl/workbook.xml"))
        rels = ET.fromstring(self.zip.read("xl/_rels/workbook.xml.rels"))
        destinos = {r.get("Id"): r.get("Target") for r in rels}
        abas = {}
        for s in wb.iter(NS + "sheet"):
            destino = destinos[s.get(RNS + "id")].lstrip("/")
            abas[s.get("name")] = destino if destino.startswith("xl/") else "xl/" + destino
        return abas

    def _carregar_shared_strings(self):
        if "xl/sharedStrings.xml" not in self.zip.namelist():
            return []
        strings = []
        for _, el in ET.iterparse(self.zip.open("xl/sharedStrings.xml")):
            if el.tag == NS + "si":
                strings.append("".join(t.text or "" for t in el.iter(NS + "t")))
                el.clear()
        return strings

    @staticmethod
    def _coluna(letras):
        n = 0
        for ch in letras:
            n = n * 26 + ord(ch) - 64
        return n - 1

    @staticmethod
    def _numero(texto):
        try:
            return int(texto)
        except ValueError:
            return float(texto)

    def _valor(self, c):
        tipo = c.get("t", "n")
        if tipo == "inlineStr":
            return "".join(t.text or "" for t in c.iter(NS + "t"))
        v = c.find(NS + "v")
        if v is None or v.text is None:
            return None
        if tipo == "s":
            return self.shared_strings[int(v.text)]
        if tipo == "n":
            return self._numero(v.text)
        if tipo == "b":
            return v.text == "1"
        return v.text  # "str" (resultado de fórmula) ou "e" (erro)

    def linhas(self, aba):
        """Gera (indice_linha, {indice_coluna: valor}) começando em 0"""
        for _, el in ET.iterparse(self.zip.open(self.abas[aba])):
            if el.tag != NS + "row":
                continue
            celulas = {}
            for c in el.iter(NS + "c"):
                valor = self._valor(c)
                if valor is None or valor == "":
                    continue
                letras, _ = PATTERN_REF.match(c.get("r")).groups()
                celulas[self._coluna(letras)] = valor
            yield int(el.get("r")) - 1, celulas
            el.clear()


# ====================================
# ESCRITA
# ====================================
def escrever_linha(ws, linha, celulas):
    for col, valor in sorted(celulas.items()):
        if isinstance(valor, bool):
            ws.write_boolean(linha, col, valor)
        elif isinstance(valor, (int, float)):
            ws.write_number(linha, col, valor)
        else:
            # write_string evita que textos iniciados com "=" virem fórmula
            ws.write_string(linha, col, valor)


def copiar_aba(leitor, ws):
    total = 0
    for linha, celulas in leitor.linhas(ws.name):
        if celulas:
            escrever_linha(ws, linha, celulas)
            total += 1
    return total


def reduzir_detalhe(leitor, ws):
    """
    Cada bloco de jogo tem: título "Game Name: ...", linha em branco,
    "Common", cabeçalho "Union ID ..." e as linhas de jogadores.
    O bloco é mantido com o mesmo layout, mas só com jogadores H2,
    e o próximo bloco começa logo após a última linha escrita.
    """
    proxima_linha = 0      # próxima linha livre na saída
    inicio_bloco = None    # (linha_origem_titulo, linha_saida_titulo)
    col_clube = None       # coluna "Club Name" do cabeçalho atual
    blocos = jogadores = descartados = 0

    for linha, celulas in leitor.linhas(ABA_DETALHE):
        if not celulas:
            continue
        primeira = str(celulas.get(min(celulas), ""))

        if primeira.startswith("Game Name:"):
            inicio_bloco = (linha, proxima_linha)
            col_clube = None
            escrever_linha(ws, proxima_linha, celulas)
            proxima_linha += 1
            blocos += 1
            continue

        if inicio_bloco is None:
            continue

        if col_clube is None:
            # linhas antes dos jogadores ("Common" e cabeçalho) mantêm a distância do título
            destino = inicio_bloco[1] + (linha - inicio_bloco[0])
            escrever_linha(ws, destino, celulas)
            proxima_linha = destino + 1
            if primeira == "Union ID":
                col_clube = next((c for c, v in celulas.items() if v == "Club Name"), None)
            continue

        if PATTERN_CLUBE_H2.search(str(celulas.get(col_clube, ""))):
            escrever_linha(ws, proxima_linha, celulas)
            proxima_linha += 1
            jogadores += 1
        else:
            descartados += 1

    return blocos, jogadores, descartados


# ====================================
# PRINCIPAL
# ====================================
def reduzir(entrada, saida):
    inicio = time.time()
    log(f"Lendo {entrada}")
    leitor = LeitorXlsx(entrada)

    faltando = [a for a in ABAS if a not in leitor.abas]
    if faltando:
        raise ValueError(f"Abas não encontradas na planilha: {', '.join(faltando)}")

    # Planilhas 106/107 também têm abas Grand Union, mas o importador delas usa as abas "Union *"
    cabecalho = " ".join(
        str(v) for _, (_, celulas) in zip(range(3), leitor.linhas(ABAS[0])) for v in celulas.values()
    )
    if "Union:113" not in cabecalho:
        raise ValueError("A planilha não é da liga 113 (Grand Union); só ela deve ser reduzida.")

    wb = xlsxwriter.Workbook(saida, {"constant_memory": True})
    try:
        for aba in ABAS:
            ws = wb.add_worksheet(aba)
            if aba == ABA_DETALHE:
                blocos, jogadores, descartados = reduzir_detalhe(leitor, ws)
                log(f"{aba}: {blocos} jogos, {jogadores} linhas H2 mantidas, {descartados} descartadas")
            else:
                log(f"{aba}: {copiar_aba(leitor, ws)} linhas copiadas")
    finally:
        wb.close()

    antes = os.path.getsize(entrada) / 1024 / 1024
    depois = os.path.getsize(saida) / 1024 / 1024
    log(f"✅ Reduzida: {antes:.0f} MB → {depois:.0f} MB em {time.time() - inicio:.0f}s")


def processar_pendentes():
    for pasta in (PENDENTES_PATH, PROCESSADOS_PATH, ORIGINAIS_PATH):
        os.makedirs(pasta, exist_ok=True)

    arquivos = sorted(
        f for f in os.listdir(PENDENTES_PATH)
        if f.lower().endswith(".xlsx") and not f.startswith(("~$", ".~lock"))
    )
    if not arquivos:
        log("Nenhuma planilha em pendentes/.")
        return True

    log(f"{len(arquivos)} planilha(s) em pendentes/")
    falhas = []
    for nome in arquivos:
        entrada = os.path.join(PENDENTES_PATH, nome)
        saida = os.path.join(PROCESSADOS_PATH, nome)
        temporario = saida + ".tmp"
        try:
            reduzir(entrada, temporario)
            os.replace(temporario, saida)
            shutil.move(entrada, os.path.join(ORIGINAIS_PATH, nome))
            log(f"→ {nome} movido para processados/ (original em processados/originais/)")
        except Exception as e:
            if isinstance(e, ValueError):
                log(f"⚠️ {nome}: {e}")
            else:
                traceback.print_exc()
            if os.path.exists(temporario):
                os.remove(temporario)
            falhas.append(nome)
            log(f"❌ Falha ao reduzir {nome}, mantido em pendentes/")

    if falhas:
        log(f"Concluído com {len(falhas)} falha(s): {', '.join(falhas)}")
    else:
        log("✅ Todas as planilhas foram processadas.")
    return not falhas


def main():
    parser = argparse.ArgumentParser(description="Reduz a planilha Suprema 113 para o importador")
    parser.add_argument("entrada", nargs="?",
                        help="planilha específica (.xlsx); sem argumento processa a pasta pendentes/")
    parser.add_argument("-o", "--saida", help="arquivo de saída (padrão: <nome>_reduzido.xlsx)")
    args = parser.parse_args()

    if not args.entrada:
        sys.exit(0 if processar_pendentes() else 1)

    saida = args.saida or re.sub(r"\.xlsx$", "", args.entrada, flags=re.IGNORECASE) + "_reduzido.xlsx"
    if os.path.abspath(saida) == os.path.abspath(args.entrada):
        sys.exit("❌ A saída não pode sobrescrever a planilha original.")

    reduzir(args.entrada, saida)
    log(f"Arquivo gerado: {saida}")


if __name__ == "__main__":
    main()
