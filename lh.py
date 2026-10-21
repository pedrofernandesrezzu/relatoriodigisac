"""
Gera o Excel de resultado de campanhas do Digisac (mesmo layout do
quebra_geral_19_10_26_export.csv) e separa os números sem WhatsApp.

Uso:
  # 1) Todas as campanhas de um mês (ex.: outubro/2026)
  python exportar_campanha_digisac.py --mes 2026-10

  # 2) Todas do mês filtrando pelo nome
  python exportar_campanha_digisac.py --mes 2026-10 --nome quebra_geral

  # 3) Uma campanha por ID
  python exportar_campanha_digisac.py --campanha <ID>

  # 4) Procurando campanha pelo nome e data
  python exportar_campanha_digisac.py --nome quebra_geral --data 2026-10-19

  # 5) A partir de um CSV já baixado do Digisac
  python exportar_campanha_digisac.py --csv quebra_geral_19_10_26_export.csv

Variáveis de ambiente:
  DIGISAC_URL    (padrão: https://rezzusolucoes.digisac.biz)
  DIGISAC_TOKEN  token de acesso pessoal (Conta > API > Tokens de acesso pessoal)
"""

import argparse
import io
import json
import os
import sys
from datetime import datetime, timedelta

import pandas as pd
import requests
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

URL_BASE = "https://rezzusolucoes.digisac.biz"
TOKEN = "b91d4805700204b79038d14676155bc55394d3f6"

COL_NUMERO = "Número do contato"
COL_STATUS = "Status do disparo"
COL_MOTIVO = "Motivo da falha"
COL_CAMPANHA = "Campanha"
TEXTO_SEM_WHATSAPP = "não está registrado no WhatsApp"

# Fuso de Brasília (UTC-3): meia-noite local = 03:00 UTC
OFFSET_UTC = timedelta(hours=3)
POR_PAGINA = 50


# ---------------------------------------------------------------- API Digisac
def headers():
    if not TOKEN:
        sys.exit("Defina a variável de ambiente DIGISAC_TOKEN com o seu token.")
    return {"Authorization": f"Bearer {TOKEN}", "Content-Type": "application/json"}


def iso_utc(dt_local):
    return (dt_local + OFFSET_UTC).strftime("%Y-%m-%dT%H:%M:%S.000Z")


def listar_campanhas(inicio_local, fim_local, nome=None):
    """Lista TODAS as campanhas criadas entre inicio_local e fim_local (paginando)."""
    campanhas, pagina = [], 1
    while True:
        query = {
            "where": {
                "$and": [
                    {"createdAt": {"$gte": iso_utc(inicio_local)}},
                    {"createdAt": {"$lt": iso_utc(fim_local)}},
                ]
            },
            "order": [["createdAt", "ASC"]],
            "page": pagina,
            "perPage": POR_PAGINA,
        }
        r = requests.get(
            f"{URL_BASE}/api/v1/campaigns",
            headers=headers(),
            params={"query": json.dumps(query)},
            timeout=60,
        )
        r.raise_for_status()
        corpo = r.json()
        lote = corpo.get("data", corpo) if isinstance(corpo, dict) else corpo
        campanhas.extend(lote)
        if len(lote) < POR_PAGINA:
            break
        pagina += 1

    if nome:
        campanhas = [c for c in campanhas if nome.lower() in (c.get("name") or "").lower()]
    return campanhas


def limites_do_mes(mes):
    """'2026-10' -> (01/10 00:00, 01/11 00:00)."""
    inicio = datetime.strptime(mes, "%Y-%m")
    fim = datetime(inicio.year + (inicio.month == 12), inicio.month % 12 + 1, 1)
    return inicio, fim


def exportar_csv_campanha(campanha_id):
    """Chama o export de resultados da campanha e devolve o CSV como DataFrame."""
    r = requests.post(
        f"{URL_BASE}/api/v1/campaigns/export/csv",
        headers=headers(),
        json={"where": {"campaignId": campanha_id}, "type": "semiColon"},
        timeout=120,
    )
    r.raise_for_status()

    tipo = r.headers.get("Content-Type", "")
    if "json" in tipo:
        # Algumas versões devolvem um JSON com o link do arquivo
        corpo = r.json()
        link = None
        if isinstance(corpo, dict):
            link = corpo.get("url") or corpo.get("link") or (corpo.get("file") or {}).get("url")
        if not link:
            sys.exit(f"Resposta inesperada do export: {str(corpo)[:300]}")
        if link.startswith("/"):
            link = URL_BASE + link
        r = requests.get(link, headers=headers(), timeout=120)
        r.raise_for_status()

    texto = r.content.decode("utf-8-sig")
    return pd.read_csv(io.StringIO(texto), sep=";", dtype=str, keep_default_na=False)


def exportar_varias(campanhas):
    """Exporta cada campanha e junta tudo num DataFrame com a coluna 'Campanha'."""
    partes = []
    for i, c in enumerate(campanhas, start=1):
        nome = c.get("name") or c["id"]
        print(f"[{i}/{len(campanhas)}] {nome} ({c['id']})")
        try:
            df = exportar_csv_campanha(c["id"])
        except Exception as e:  # uma campanha com erro não derruba as outras
            print(f"   ! falhou: {e}")
            continue
        if df.empty:
            print("   (sem linhas)")
            continue
        df.insert(0, COL_CAMPANHA, nome)
        partes.append(df)
    if not partes:
        sys.exit("Nenhuma campanha com resultados para exportar.")
    return pd.concat(partes, ignore_index=True)


# ------------------------------------------------------------------- Excel
def estilizar(ws, df):
    cabecalho = PatternFill("solid", start_color="1F4E78")
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = cabecalho
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for linha in ws.iter_rows(min_row=2):
        for c in linha:
            c.font = Font(name="Arial", size=10)
            c.number_format = "@"  # tudo como texto: não perde o 55 nem vira notação científica
    for i, col in enumerate(df.columns, start=1):
        maior = max([len(str(col))] + [len(str(v)) for v in df[col].head(500)])
        ws.column_dimensions[get_column_letter(i)].width = min(max(maior + 2, 12), 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def montar_resumo(df):
    df = df.copy()
    df["_sem_wa"] = df[COL_MOTIVO].str.contains(TEXTO_SEM_WHATSAPP, case=False, na=False)
    if COL_CAMPANHA in df.columns:
        g = df.groupby(COL_CAMPANHA, sort=False)
        resumo = pd.DataFrame(
            {
                "Total de contatos": g.size(),
                "Com sucesso": g[COL_STATUS].apply(lambda s: int((s == "Com sucesso").sum())),
                "Com falha": g[COL_STATUS].apply(lambda s: int((s == "Com falha").sum())),
                "Sem WhatsApp": g["_sem_wa"].sum().astype(int),
            }
        ).reset_index()
        total = {COL_CAMPANHA: "TOTAL"}
        for col in resumo.columns[1:]:
            total[col] = int(resumo[col].sum())
        return pd.concat([resumo, pd.DataFrame([total])], ignore_index=True)

    return pd.DataFrame(
        {
            "Indicador": ["Total de contatos", "Com sucesso", "Com falha", "Sem WhatsApp"],
            "Quantidade": [
                len(df),
                int((df[COL_STATUS] == "Com sucesso").sum()),
                int((df[COL_STATUS] == "Com falha").sum()),
                int(df["_sem_wa"].sum()),
            ],
        }
    )


def gerar_excel(df, caminho):
    mask = df[COL_MOTIVO].str.contains(TEXTO_SEM_WHATSAPP, case=False, na=False)
    sem_whatsapp = df[mask]
    resumo = montar_resumo(df)

    with pd.ExcelWriter(caminho, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="Geral", index=False)
        sem_whatsapp.to_excel(xw, sheet_name="Sem WhatsApp", index=False)
        resumo.to_excel(xw, sheet_name="Resumo", index=False)
        estilizar(xw.sheets["Geral"], df)
        estilizar(xw.sheets["Sem WhatsApp"], sem_whatsapp)
        estilizar(xw.sheets["Resumo"], resumo)

    print(f"\nExcel gerado: {caminho}")
    print(resumo.to_string(index=False))


# -------------------------------------------------------------------- main
def main():
    p = argparse.ArgumentParser(description="Resultado de campanhas do Digisac em Excel")
    p.add_argument("--mes", help="Mês AAAA-MM: pega TODAS as campanhas do mês (ex.: 2026-10)")
    p.add_argument("--campanha", help="ID da campanha no Digisac")
    p.add_argument("--nome", help="Parte do nome da campanha (ex.: quebra_geral)")
    p.add_argument("--data", help="Data de criação da campanha, AAAA-MM-DD (usada com --nome)")
    p.add_argument("--csv", help="CSV já exportado do Digisac (não usa a API)")
    p.add_argument("--saida", help="Caminho do .xlsx de saída")
    a = p.parse_args()

    if a.csv:
        df = pd.read_csv(a.csv, sep=";", dtype=str, keep_default_na=False, encoding="utf-8-sig")
        base = os.path.splitext(os.path.basename(a.csv))[0]
    elif a.campanha:
        df = exportar_csv_campanha(a.campanha)
        base = f"campanha_{a.campanha[:8]}_export"
    elif a.mes:
        inicio, fim = limites_do_mes(a.mes)
        campanhas = listar_campanhas(inicio, fim, a.nome)
        if not campanhas:
            sys.exit(f"Nenhuma campanha encontrada em {a.mes}.")
        print(f"{len(campanhas)} campanha(s) em {a.mes}:")
        df = exportar_varias(campanhas)
        base = f"campanhas_{inicio.strftime('%m_%Y')}_export"
    elif a.nome:
        data = a.data or datetime.now().strftime("%Y-%m-%d")
        dia = datetime.strptime(data, "%Y-%m-%d")
        campanhas = listar_campanhas(dia, dia + timedelta(days=1), a.nome)
        if not campanhas:
            sys.exit(f"Nenhuma campanha com '{a.nome}' encontrada em {data}.")
        df = exportar_varias(campanhas)
        base = f"{a.nome}_{dia.strftime('%d_%m_%y')}_export"
    else:
        p.error("Informe --mes, --campanha, --nome ou --csv")

    gerar_excel(df, a.saida or f"{base}.xlsx")


if __name__ == "__main__":
    main()