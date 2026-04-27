#!/usr/bin/env python3
"""
Atualização automática — Dashboard Comercial FuelTech
Lê o relatório Excel mais recente da pasta e atualiza o dashboard HTML.
Uso: python3 atualizar_dashboard.py
"""

import os, sys, glob, json, re
from datetime import datetime, date
from collections import defaultdict
from calendar import monthrange

try:
    import openpyxl
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl --break-system-packages -q")
    import openpyxl

# ── Configurações ──────────────────────────────────────────────────────────────
SCRIPT_DIR   = os.path.dirname(os.path.abspath(__file__))
DASHBOARD    = os.path.join(SCRIPT_DIR, "dashboard_fueltech_v4.html")
LOG_FILE     = os.path.join(SCRIPT_DIR, "atualizacao_dashboard.log")

VENDEDORES = [
    "GIUMBELLI MARTINELLO",
    "SAULO DE OLIVEIRA MOREIRA",
    "VITORIO MUNHOZ",
    "LEONARDO ARESCO",
    "GUSTAVO SCOUTO CAMPOS",
    "OMAR MORAIS",
    "DIEGO VIEIRA DOS SANTOS",
    "JOAO BOPP",
    "FERNANDA FERNANDES",
]

# Canais internos (loja + eventos) — fonte do SITE_DATA
CANAIS_LOJA    = {"Loja 1803"}
CANAIS_EVENTOS = {"Eventos 1801", "Eventos 1802", "Eventos 1807", "Eventos 1808", "Eventos 1809"}
CANAIS_SITE    = CANAIS_LOJA | CANAIS_EVENTOS

MESES_PT = {1:"jan",2:"fev",3:"mar",4:"abr",5:"mai",6:"jun",
            7:"jul",8:"ago",9:"set",10:"out",11:"nov",12:"dez"}
MESES_NOME = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",
              6:"Junho",7:"Julho",8:"Agosto",9:"Setembro",
              10:"Outubro",11:"Novembro",12:"Dezembro"}

# ── Utilitários ────────────────────────────────────────────────────────────────
def log(msg):
    ts = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def parse_date(val):
    return datetime.strptime(str(int(val)), "%Y%m%d").date()

def fmt_br(d):
    return d.strftime("%d/%m/%Y")

def mes_key(d):
    return f"{d.year}-{d.month:02d}"

def dias_desde(d, hoje):
    return (hoje - d).days

def semana_label(d, mes, ano, last_day_data):
    day = d.day
    m   = MESES_PT[mes]
    if   day <=  7: return f"01-07/{m}"
    elif day <= 14: return f"08-14/{m}"
    elif day <= 21: return f"15-21/{m}"
    else:           return f"22-{last_day_data}/{m}"

def pareto_list(dct, top=20):
    items = sorted(dct.items(), key=lambda x: -x[1])[:top]
    total = sum(v for _, v in items)
    acum  = 0
    result = []
    for nome, valor in items:
        pct   = round(valor / total * 100, 1) if total else 0
        acum += pct
        result.append({
            "nome": nome, "valor": round(valor, 2),
            "pct": pct, "acum": round(acum, 1),
            "pareto": 1 if acum <= 80 else 0
        })
    return result

# ── Carregamento do Excel ──────────────────────────────────────────────────────
def list_complete_excels():
    """Retorna todos os arquivos rel_vendas_*.xlsx >= 1MB, ordenados pelo nome."""
    all_files = glob.glob(os.path.join(SCRIPT_DIR, "rel_vendas_*.xlsx"))
    if not all_files:
        raise FileNotFoundError("Nenhum arquivo rel_vendas_*.xlsx encontrado em: " + SCRIPT_DIR)
    complete = sorted(
        [f for f in all_files if os.path.getsize(f) >= 1_000_000],
        key=lambda f: os.path.basename(f)
    )
    skipped = [os.path.basename(f) for f in all_files if f not in complete]
    if skipped:
        log(f"  Ignorados (< 1MB): {', '.join(skipped)}")
    if not complete:
        raise FileNotFoundError("Nenhum arquivo >= 1MB encontrado.")
    return complete          # ordenados do mais antigo ao mais recente

def find_latest_excel():
    """Retorna apenas o arquivo mais recente (>= 1MB)."""
    return list_complete_excels()[-1]

def _load_one(path):
    """Carrega um único Excel e detecta automaticamente a linha do cabeçalho.
    Usa pandas quando disponível (muito mais rápido para arquivos grandes).
    """
    try:
        import pandas as pd
        import numpy as np

        # Escolher engine: calamine é muito mais rápido que openpyxl
        try:
            import python_calamine  # noqa
            engine = 'calamine'
        except ImportError:
            engine = 'openpyxl'

        # Detectar header lendo apenas as 2 primeiras linhas
        df0 = pd.read_excel(path, engine=engine, header=0, nrows=1)
        row0_strings = sum(1 for v in df0.columns
                           if isinstance(v, str) and v.strip() and not v.startswith('Unnamed'))
        header_row = 0 if row0_strings >= 5 else 1

        df = pd.read_excel(path, engine=engine, header=header_row)
        cols = {v: i for i, v in enumerate(df.columns) if v and not str(v).startswith('Unnamed')}
        # Converte para lista de tuplas (mesmo formato que openpyxl retornava)
        data = [tuple(None if (isinstance(v, float) and np.isnan(v)) else v for v in row)
                for row in df.itertuples(index=False, name=None)]
        return data, cols
    except Exception:
        # Fallback para openpyxl puro
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        wb.close()
        row0_strings = sum(1 for v in rows[0] if isinstance(v, str) and v.strip())
        if row0_strings >= 5:
            header, data = rows[0], rows[1:]
        else:
            header, data = rows[1], rows[2:]
        cols = {v: i for i, v in enumerate(header) if v}
        return data, cols

def load_excel(path):
    """Carrega o arquivo mais recente (para métricas do mês atual)."""
    log(f"Lendo: {os.path.basename(path)}")
    data, cols = _load_one(path)
    log(f"  {len(data):,} linhas · colunas: {list(cols.keys())[:6]}...")
    return data, cols

def load_all_excels():
    """Carrega e mescla TODOS os arquivos completos (histórico + atual).
    Deduplica por (Nf_saida, Item) para evitar dupla contagem em caso de overlap.
    Retorna: (all_rows, cols, latest_path)
    """
    files = list_complete_excels()
    log(f"Carregando {len(files)} arquivo(s) de dados:")

    all_rows = []
    seen     = set()   # chave de deduplicação: (nf, item)
    cols_ref = None

    for path in files:
        data, cols = _load_one(path)
        i_nf   = cols.get('Nf_saida', -1)
        i_item = cols.get('Item',     -1)
        i_data = cols.get('Data',     -1)

        added = 0
        for r in data:
            # Chave única por linha de NF
            key = (r[i_nf] if i_nf >= 0 else None,
                   r[i_item] if i_item >= 0 else None)
            if key in seen:
                continue
            seen.add(key)
            all_rows.append(r)
            added += 1

        if cols_ref is None:
            cols_ref = cols

        # Detectar período deste arquivo
        dates = []
        for r in data[:5000]:
            try:
                dates.append(parse_date(r[i_data]))
            except:
                pass
        periodo = f"{min(dates)} → {max(dates)}" if dates else "?"
        log(f"  ✓ {os.path.basename(path)}: {added:,} linhas únicas · {periodo}")

    log(f"  Total mesclado: {len(all_rows):,} linhas de {len(files)} arquivo(s)")
    return all_rows, cols_ref, files[-1]

def detect_ref_month(rows, cols):
    """Detecta mês/ano de referência = mês mais recente nos dados."""
    dates = []
    for r in rows:
        try:
            if r[cols['Ano']] and int(r[cols['Ano']]) >= 2025:
                dates.append(parse_date(r[cols['Data']]))
        except:
            pass
    max_d = max(dates)
    min_d = min(dates)
    log(f"  Período Excel: {min_d} → {max_d}")
    return max_d.month, max_d.year, max_d

# ── CARTEIRA ───────────────────────────────────────────────────────────────────
def build_carteira(rows, cols, hoje):
    log("Construindo CARTEIRA...")
    carteira = {}

    for vend in VENDEDORES:
        vend_rows = [r for r in rows if r[cols['Nome_vendedor']] == vend]
        if not vend_rows:
            log(f"  Sem dados para {vend}")
            continue

        cli_agg = defaultdict(lambda: {
            "fat":0.0,"f25":0.0,"f26":0.0,"fja25":0.0,
            "nfs":set(),"its":0,"dates":[],
            "fmes":defaultdict(float),"prods":defaultdict(float),"can":None
        })

        for r in vend_rows:
            cli  = r[cols['Nome_cliente']]
            fat  = float(r[cols['Faturamento_sem_icms_ipi']] or 0)
            ano  = int(r[cols['Ano']]  or 0)
            mes  = int(r[cols['Mes']]  or 0)
            qtd  = int(r[cols['Qtd']]  or 0)
            nf   = r[cols['Nf_saida']]
            prod = r[cols['Desc_produto']] or ''
            tab  = r[cols['Tab_vendas_nome']] or 'PROTHEUS'
            try:
                d = parse_date(r[cols['Data']])
            except:
                continue

            c = cli_agg[cli]
            c["fat"] += fat
            if ano == 2025:
                c["f25"] += fat
                if mes <= 4:
                    c["fja25"] += fat
            elif ano == 2026:
                c["f26"] += fat
            c["its"]   += qtd
            c["nfs"].add(nf)
            c["dates"].append(d)
            c["fmes"][mes_key(d)] += fat
            c["prods"][prod]      += fat
            if c["can"] is None:
                c["can"] = tab or "PROTHEUS"

        cli_final = {}
        for cli, c in cli_agg.items():
            if c["fat"] <= 0 or not c["dates"]:
                continue
            last_d = max(c["dates"])
            a26 = c["f26"] > 0
            a25 = c["f25"] > 0

            # YoY: Jan-Abr/2025 vs Jan-Abr/2026
            fja26 = sum(v for mk, v in c["fmes"].items()
                        if mk.startswith("2026-") and int(mk.split("-")[1]) <= 4)
            yoy = None
            if c["fja25"] > 0 and fja26 > 0:
                yoy = round((fja26 / c["fja25"] - 1) * 100, 1)

            tp5 = sorted(c["prods"].items(), key=lambda x: -x[1])[:5]

            cli_final[cli] = {
                "fat":   round(c["fat"],   2),
                "f25":   round(c["f25"],   2),
                "f26":   round(c["f26"],   2),
                "fja25": round(c["fja25"], 2),
                "ped":   len(c["nfs"]),
                "its":   c["its"],
                "ult":   fmt_br(last_d),
                "rec":   dias_desde(last_d, hoje),
                "a26":   a26,
                "a25":   a25,
                "fmes":  {k: round(v, 2) for k, v in sorted(c["fmes"].items())},
                "tp":    [[p, round(v, 2)] for p, v in tp5],
                "yoy":   yoy,
                "can":   c["can"]
            }

        if not cli_final:
            continue

        tf   = sum(c["fat"] for c in cli_final.values())
        f25  = sum(c["f25"] for c in cli_final.values())
        f26  = sum(c["f26"] for c in cli_final.values())
        nc   = len(cli_final)
        na26 = sum(1 for c in cli_final.values() if c["a26"])
        ni   = sum(1 for c in cli_final.values() if not c["a26"])
        nn   = sum(1 for c in cli_final.values() if c["a26"] and not c["a25"])

        carteira[vend] = {
            "res": {"tf":round(tf,2),"f25":round(f25,2),"f26":round(f26,2),
                    "nc":nc,"na26":na26,"ni":ni,"nn":nn},
            "cli": cli_final
        }
        log(f"  {vend}: {nc} clientes")

    return carteira

# ── DATA (mês atual) ───────────────────────────────────────────────────────────
def build_data(rows, cols, mes, ano, last_day_data):
    log(f"Construindo DATA ({MESES_NOME[mes]}/{ano})...")
    m_str = MESES_PT[mes]
    sem_keys = [
        f"01-07/{m_str}", f"08-14/{m_str}",
        f"15-21/{m_str}", f"22-{last_day_data}/{m_str}"
    ]
    result = {}

    for vend in VENDEDORES:
        vend_rows = [r for r in rows
                     if r[cols['Nome_vendedor']] == vend
                     and int(r[cols['Ano']] or 0) == ano
                     and int(r[cols['Mes']] or 0) == mes]

        if not vend_rows:
            result[vend] = {"total":0,"pedidos":0,
                            "semanas":{k:0 for k in sem_keys},"clientes":[]}
            continue

        total = sum(float(r[cols['Faturamento_sem_icms_ipi']] or 0) for r in vend_rows)
        nfs   = set(r[cols['Nf_saida']] for r in vend_rows)

        sem     = defaultdict(float)
        cli_fat = defaultdict(float)
        for r in vend_rows:
            try:
                d   = parse_date(r[cols['Data']])
                fat = float(r[cols['Faturamento_sem_icms_ipi']] or 0)
                sem[semana_label(d, mes, ano, last_day_data)] += fat
                cli_fat[r[cols['Nome_cliente']]]              += fat
            except:
                pass

        result[vend] = {
            "total":   round(total, 2),
            "pedidos": len(nfs),
            "semanas": {k: round(sem.get(k, 0), 2) for k in sem_keys},
            "clientes": pareto_list(cli_fat)
        }

    return result

# ── FT700_DATA ─────────────────────────────────────────────────────────────────
def build_ft700(rows, cols, mes, ano, last_day_data):
    log("Construindo FT700_DATA...")
    m_str    = MESES_PT[mes]
    sem_keys = [
        f"01-07/{m_str}", f"08-14/{m_str}",
        f"15-21/{m_str}", f"22-{last_day_data}/{m_str}"
    ]
    result = {}

    for vend in VENDEDORES:
        ft_rows = [r for r in rows
                   if r[cols['Nome_vendedor']] == vend
                   and int(r[cols['Ano']] or 0) == ano
                   and int(r[cols['Mes']] or 0) == mes
                   and r[cols['Desc_produto']] in ("FT700", "FT700PLUS")]

        if not ft_rows:
            continue

        total    = sum(float(r[cols['Faturamento_sem_icms_ipi']] or 0) for r in ft_rows)
        sem      = defaultdict(float)
        por_prod = defaultdict(float)

        for r in ft_rows:
            try:
                d   = parse_date(r[cols['Data']])
                fat = float(r[cols['Faturamento_sem_icms_ipi']] or 0)
                sem[semana_label(d, mes, ano, last_day_data)] += fat
                por_prod[r[cols['Desc_produto']]]             += fat
            except:
                pass

        if total > 0:
            result[vend] = {
                "total":      round(total, 2),
                "semanas":    {k: round(sem.get(k, 0), 2) for k in sem_keys},
                "por_produto":{k: round(v, 2) for k, v in por_prod.items()}
            }

    return result

# ── SITE_DATA ──────────────────────────────────────────────────────────────────
def build_site_data(rows, cols, mes, ano, last_day_data):
    log("Construindo SITE_DATA...")
    m_str    = MESES_PT[mes]
    sem_keys = [
        f"01-07/{m_str}", f"08-14/{m_str}",
        f"15-21/{m_str}", f"22-{last_day_data}/{m_str}"
    ]

    site_rows = [r for r in rows
                 if r[cols['Nome_vendedor']] in CANAIS_SITE
                 and int(r[cols['Ano']] or 0) == ano
                 and int(r[cols['Mes']] or 0) == mes]

    if not site_rows:
        log("  Sem dados de site/loja/eventos")
        return {}

    sem      = defaultdict(float)
    estados  = defaultdict(float)
    produtos = defaultdict(float)
    clientes = defaultdict(float)
    grupos   = defaultdict(float)

    for r in site_rows:
        fat = float(r[cols['Faturamento_sem_icms_ipi']] or 0)
        try:
            d = parse_date(r[cols['Data']])
            sem[semana_label(d, mes, ano, last_day_data)] += fat
        except:
            pass
        estados [r[cols['Estado']]      or "N/A"] += fat
        produtos[r[cols['Desc_produto']]or "N/A"] += fat
        clientes[r[cols['Nome_cliente']]or "N/A"] += fat
        grupos  [r[cols['Desc_grupo']]  or "N/A"] += fat

    total_grupos = sum(grupos.values()) or 1
    grupos_list  = [
        {"nome": k, "valor": round(v, 2), "pct": round(v/total_grupos*100, 1)}
        for k, v in sorted(grupos.items(), key=lambda x: -x[1])[:10]
    ]

    return {
        "semanas":  {k: round(sem.get(k, 0), 2) for k in sem_keys},
        "estados":  pareto_list(estados),
        "produtos": pareto_list(produtos),
        "clientes": pareto_list(clientes),
        "grupos":   grupos_list
    }

# ── Atualizar HTML ─────────────────────────────────────────────────────────────
def replace_js_var(content, var_name, value, has_const=True):
    """Substitui uma variável JS na linha exata onde ela aparece.
    Aceita tanto 'const X=' quanto 'const X =' (com ou sem espaço)."""
    json_val = json.dumps(value, ensure_ascii=False, separators=(',', ':'))
    # Prefixos possíveis (com ou sem espaço antes do =)
    if has_const:
        prefixes   = [f"const {var_name}=", f"const {var_name} ="]
        new_line   = f"const {var_name}={json_val};"
    else:
        prefixes   = [f"{var_name}=", f"{var_name} ="]
        new_line   = f"{var_name}={json_val};"

    lines = content.split('\n')
    found = False
    for i, line in enumerate(lines):
        stripped = line.strip()
        if any(stripped.startswith(p) for p in prefixes):
            lines[i] = new_line
            found = True
            break

    if not found:
        log(f"  AVISO: '{var_name}' não encontrada no HTML")
    else:
        log(f"  ✓ {var_name}")

    return '\n'.join(lines)

def replace_inline_val(content, var_name, new_val):
    """Substitui um valor numérico simples inline (ex: REAL_SITE=1846969)."""
    pattern = rf'({re.escape(var_name)}=)\d+(\.\d+)?'
    if re.search(pattern, content):
        result = re.sub(pattern, rf'\g<1>{new_val}', content)
        log(f"  ✓ {var_name} = {new_val}")
        return result
    else:
        log(f"  AVISO: '{var_name}' inline não encontrada")
        return content

def replace_text(content, old, new):
    result = content.replace(old, new)
    if result != content:
        log(f"  ✓ Texto '{old}' → '{new}'")
    return result

def build_cart_meses(mes_ref, ano_ref):
    """Gera CART_MESES e CART_ML de Jan/2025 até o mês de referência."""
    meses, labels = [], []
    ano, mes = 2025, 1
    abrev_ano = {2025:"25", 2026:"26", 2027:"27"}
    nomes_curtos = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]
    while (ano, mes) <= (ano_ref, mes_ref):
        meses.append(f"{ano}-{mes:02d}")
        labels.append(f"{nomes_curtos[mes-1]}/{abrev_ano.get(ano, str(ano)[2:])}")
        mes += 1
        if mes > 12:
            mes = 1
            ano += 1
    return meses, labels

def update_html(path, mes, ano, data_js, carteira, ft700, site_data, hoje, last_day_data):
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    mes_nome   = MESES_NOME[mes]
    mes_abrev  = MESES_PT[mes].capitalize()
    periodo    = f"{mes_nome[:3].capitalize()} {ano}"   # ex: "Abr 2026"
    periodo_lg = f"{mes_nome} {ano}"                    # ex: "Abril 2026"

    # --- Variáveis JS principais ---
    content = replace_js_var(content, "DATA",      data_js)
    content = replace_js_var(content, "CARTEIRA",  carteira)
    content = replace_js_var(content, "FT700_DATA",ft700)
    content = replace_js_var(content, "SITE_DATA", site_data)

    # --- REALIZADO_META ---
    realizado = {v: round(data_js[v]["total"], 0) for v in VENDEDORES if v in data_js}
    content   = replace_js_var(content, "REALIZADO_META", realizado, has_const=True)

    # --- TOTAL_FT700 ---
    total_ft700 = sum(v["total"] for v in ft700.values())
    content     = replace_inline_val(content, "TOTAL_FT700", round(total_ft700, 2))

    # --- REAL_SITE (total de loja + eventos) ---
    real_site = sum(v for v in site_data.get("semanas", {}).values())
    content   = replace_inline_val(content, "REAL_SITE", int(real_site))

    # --- CART_MESES e CART_ML ---
    cart_meses, cart_ml = build_cart_meses(mes, ano)

    lines = content.split('\n')
    for i, line in enumerate(lines):
        if line.strip().startswith("const CART_MESES"):
            lines[i] = f"const CART_MESES = {json.dumps(cart_meses)};"
            log("  ✓ CART_MESES")
        elif line.strip().startswith("const CART_ML"):
            lines[i] = f"const CART_ML = {json.dumps(cart_ml)};"
            log("  ✓ CART_ML")
    content = '\n'.join(lines)

    # --- Título e período no HTML ---
    # Substitui todas as ocorrências do período anterior por novo
    # Detecta dinamicamente o período antigo no título
    title_match = re.search(r'<title>FuelTech — Dashboard Comercial · (.+?)</title>', content)
    if title_match:
        old_period = title_match.group(1)
        content    = content.replace(
            f"<title>FuelTech — Dashboard Comercial · {old_period}</title>",
            f"<title>FuelTech — Dashboard Comercial · {mes_nome.capitalize()} {ano}</title>"
        )
        # Substituir todas ocorrências do período antigo no body
        for old in [old_period]:
            content = content.replace(old, f"{mes_nome.capitalize()} {ano}")
        log(f"  ✓ Título: '{old_period}' → '{mes_nome.capitalize()} {ano}'")

    # --- Backup do arquivo atual antes de sobrescrever ---
    backup = path + ".bak"
    try:
        import shutil
        shutil.copy2(path, backup)
        log(f"  ✓ Backup criado: {os.path.basename(backup)}")
    except Exception as e:
        log(f"  AVISO: não foi possível criar backup — {e}")

    # --- Salvar novo conteúdo ---
    try:
        with open(path, 'w', encoding='utf-8') as f:
            f.write(content)
        # Verificar integridade básica do arquivo salvo
        saved = open(path, encoding='utf-8').read()
        if not saved.strip().endswith('</html>'):
            log("  ERRO: arquivo salvo está truncado! Restaurando backup...")
            shutil.copy2(backup, path)
            raise RuntimeError("Arquivo salvo truncado — backup restaurado automaticamente.")
        log(f"  ✓ Dashboard salvo: {os.path.basename(path)} ({len(content):,} chars)")
    except RuntimeError:
        raise
    except Exception as e:
        log(f"  ERRO ao salvar: {e}. Tentando restaurar backup...")
        try:
            shutil.copy2(backup, path)
            log("  Backup restaurado.")
        except:
            pass
        raise

# ── Git Push via API ─────────────────────────────────────────────────────────
def git_push(mes_nome, ano):
    """Envia os arquivos para o GitHub via REST API (sem git clone).
    Muito mais rápido: apenas 2 chamadas HTTP por arquivo."""
    import urllib.request, base64, json as _json

    env_file = os.path.join(SCRIPT_DIR, ".env")
    if not os.path.exists(env_file):
        log("  AVISO: .env não encontrado — push GitHub ignorado")
        return

    cfg = {}
    for line in open(env_file).read().splitlines():
        if "=" in line and not line.startswith("#"):
            k, v = line.split("=", 1)
            cfg[k.strip()] = v.strip()

    token = cfg.get("GITHUB_TOKEN", "")
    org   = cfg.get("GITHUB_USER", "")
    repo  = "dashboard-comercial-fueltech"
    if not token or not org:
        log("  AVISO: GITHUB_TOKEN ou GITHUB_USER ausentes no .env")
        return

    headers = {
        "Authorization": f"token {token}",
        "Content-Type": "application/json",
        "User-Agent": "FuelTech-Dashboard/1.0"
    }
    commit_msg = f"auto: {mes_nome}/{ano} — {date.today().strftime('%d/%m/%Y')}"
    author     = {"name": cfg.get("GITHUB_USER","FuelTech"),
                  "email": cfg.get("GITHUB_EMAIL","noreply@fueltech.com.br")}

    files_to_push = [
        (DASHBOARD,                  os.path.basename(DASHBOARD)),
        (os.path.abspath(__file__),  os.path.basename(__file__)),
    ]

    any_updated = False
    for local_path, remote_name in files_to_push:
        api_url = f"https://api.github.com/repos/{org}/{repo}/contents/{remote_name}"

        # 1. Obter SHA atual do arquivo (necessário para update)
        sha = None
        try:
            req = urllib.request.Request(api_url, headers=headers)
            with urllib.request.urlopen(req, timeout=15) as resp:
                sha = _json.loads(resp.read()).get("sha")
        except urllib.error.HTTPError as e:
            if e.code != 404:
                log(f"  AVISO: erro ao buscar SHA de {remote_name}: {e}")
                continue

        # 2. Ler conteúdo local e encodar em base64
        raw = open(local_path, "rb").read()
        b64 = base64.b64encode(raw).decode()

        # 3. PUT para criar/atualizar o arquivo
        body = {"message": commit_msg, "content": b64, "author": author}
        if sha:
            body["sha"] = sha

        try:
            req = urllib.request.Request(
                api_url, method="PUT",
                data=_json.dumps(body).encode(),
                headers=headers
            )
            with urllib.request.urlopen(req, timeout=30) as resp:
                result = _json.loads(resp.read())
                action = "atualizado" if sha else "criado"
                log(f"  ✓ GitHub {action}: {remote_name}")
                any_updated = True
        except urllib.error.HTTPError as e:
            err_body = e.read().decode()[:200]
            log(f"  ERRO ao enviar {remote_name}: {e.code} — {err_body}")

    if not any_updated:
        log("  GitHub: nenhum arquivo atualizado")


# ── Vercel Status ──────────────────────────────────────────────────────────────
def vercel_wait_deploy():
    """Aguarda o Vercel processar o deploy disparado pelo git push e loga a URL."""
    import subprocess, time

    env_file = os.path.join(SCRIPT_DIR, ".env")
    if not os.path.exists(env_file):
        return

    cfg = {}
    for line in open(env_file).read().splitlines():
        if "=" in line and not line.startswith("#"):
            k, v = line.split("=", 1)
            cfg[k.strip()] = v.strip()

    vtoken  = cfg.get("VERCEL_TOKEN", "")
    proj_id = cfg.get("VERCEL_PROJECT_ID", "")
    if not vtoken or not proj_id:
        return

    import urllib.request, json as _json
    url = f"https://api.vercel.com/v6/deployments?projectId={proj_id}&limit=1&target=production"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {vtoken}"})

    for attempt in range(12):   # até 2 minutos
        time.sleep(10)
        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = _json.loads(resp.read())
            deployments = data.get("deployments", [])
            if not deployments:
                continue
            d = deployments[0]
            state = d.get("readyState", "?")
            deploy_url = d.get("url", "")
            if state == "READY":
                log(f"  ✓ Vercel online: https://{deploy_url}")
                return
            elif state == "ERROR":
                log(f"  ✗ Vercel deploy com erro: https://{deploy_url}")
                return
            else:
                log(f"  Vercel: {state}...")
        except Exception as e:
            log(f"  Vercel status: {e}")
            return

    log("  Vercel: timeout aguardando deploy (verifique manualmente)")


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    log("=" * 60)
    log("Dashboard FuelTech — Atualização automática")
    log("=" * 60)

    # 1. Mesclar TODOS os arquivos históricos + atual para a CARTEIRA
    all_rows, cols, latest_excel = load_all_excels()

    # 2. Detectar mês de referência a partir dos dados (filtra >= 2025)
    mes, ano, max_date = detect_ref_month(all_rows, cols)
    hoje       = date.today()
    last_day_d = max_date.day
    log(f"Referência: {MESES_NOME[mes]}/{ano} (último dado: dia {last_day_d})")

    # 3. Construir CARTEIRA com histórico completo (todos os anos)
    carteira = build_carteira(all_rows, cols, hoje)

    # 4. Métricas do mês atual: usar apenas o arquivo mais recente
    recent_rows, recent_cols = load_excel(latest_excel)
    data_js   = build_data(recent_rows, recent_cols, mes, ano, last_day_d)
    ft700     = build_ft700(recent_rows, recent_cols, mes, ano, last_day_d)
    site_data = build_site_data(recent_rows, recent_cols, mes, ano, last_day_d)

    # 5. Atualizar HTML
    log("Atualizando HTML...")
    update_html(DASHBOARD, mes, ano, data_js, carteira, ft700, site_data, hoje, last_day_d)

    # 6. Enviar para o GitHub
    log("Enviando para GitHub...")
    git_push(MESES_NOME[mes], ano)
    vercel_wait_deploy()

    log("=" * 60)
    log("✓ Atualização concluída com sucesso!")
    log("=" * 60)

if __name__ == "__main__":
    main()
