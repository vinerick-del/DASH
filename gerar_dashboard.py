"""
Gerador do Dashboard de Análise de Consumo - Ambar Energia
Fontes:
  - MB51.csv              : Saídas reais de estoque Jan-Mai/2026 (por Material+Depto+Mês)
  - DEMANDA DE MATERIAIS 28MAI26.csv : Demanda planejada Jan-Dez/2026 (por Prog+Depto+Código+Mês)
  - Saldo de estoque 29JUNHO2026.xlsx: Saldo de estoque por Código
"""
import pandas as pd
import numpy as np
import json, math, re, sys
from pathlib import Path

BASE = Path(__file__).parent

# ─────────────────────────────────────────────
# 1. LEITURA E LIMPEZA
# ─────────────────────────────────────────────

# MB51 – saídas reais
mb51 = pd.read_csv(BASE / 'MB51.csv', sep=';', encoding='latin1')
mb51.columns = ['MATERIAL','QTD','MONTANTE','DEPARTAMENTO','MES']
mb51['MATERIAL'] = pd.to_numeric(mb51['MATERIAL'], errors='coerce')
mb51['QTD']      = pd.to_numeric(mb51['QTD'], errors='coerce')
mb51 = mb51.dropna(subset=['MATERIAL','QTD'])
mb51['MATERIAL'] = mb51['MATERIAL'].astype(int)

# Demanda planejada
dem = pd.read_csv(BASE / 'DEMANDA DE MATERIAIS 28MAI26.csv', sep=';', encoding='latin1')
dem.columns = ['PROGRAMA_ORC','DEPOSITO','DEPARTAMENTO','CODIGO','DESCRICAO','MES','QTD']
dem['CODIGO'] = pd.to_numeric(dem['CODIGO'], errors='coerce')
dem['QTD']    = pd.to_numeric(dem['QTD'], errors='coerce').fillna(0)
dem = dem.dropna(subset=['CODIGO'])
dem['CODIGO'] = dem['CODIGO'].astype(int)

# Saldo de estoque
saldo = pd.read_excel(BASE / 'Saldo de estoque 29JUNHO2026.xlsx')
saldo.columns = ['CODIGO','SALDO_ATUAL']
saldo['CODIGO'] = pd.to_numeric(saldo['CODIGO'], errors='coerce')
saldo = saldo.dropna(subset=['CODIGO'])
saldo['CODIGO'] = saldo['CODIGO'].astype(int)
saldo = saldo.groupby('CODIGO', as_index=False)['SALDO_ATUAL'].sum()

# ─────────────────────────────────────────────
# 2. CONSUMO REAL (MB51) – soma algébrica por Material+Departamento
#    Saídas são negativas; devoluções positivas. Net = soma algébrica.
#    Consumo = max(0, -net)
# ─────────────────────────────────────────────
consumo = (
    mb51.groupby(['MATERIAL','DEPARTAMENTO'], as_index=False)['QTD'].sum()
    .rename(columns={'MATERIAL':'CODIGO','QTD':'NET_QTD'})
)
consumo['CONSUMO_REAL_JAN_MAI'] = (-consumo['NET_QTD']).clip(lower=0)
consumo = consumo[['CODIGO','DEPARTAMENTO','CONSUMO_REAL_JAN_MAI']]

# ─────────────────────────────────────────────
# 3. DEMANDA – separar Jan-Mai (histórico) e Jun-Dez (futuro)
# ─────────────────────────────────────────────
JAN_MAI = ['01/01/2026','01/02/2026','01/03/2026','01/04/2026','01/05/2026']
JUN_DEZ = ['01/06/2026','01/07/2026','01/08/2026','01/09/2026',
           '01/10/2026','01/11/2026','01/12/2026']

dem_hist = (
    dem[dem['MES'].isin(JAN_MAI)]
    .groupby(['CODIGO','DEPARTAMENTO'], as_index=False)['QTD'].sum()
    .rename(columns={'QTD':'DEMANDA_JAN_MAI'})
)
dem_fut = (
    dem[dem['MES'].isin(JUN_DEZ)]
    .groupby(['CODIGO','DEPARTAMENTO'], as_index=False)['QTD'].sum()
    .rename(columns={'QTD':'DEMANDA_JUN_DEZ'})
)

# Descrição por código
descricao = (
    dem.sort_values('CODIGO')
    .groupby('CODIGO', as_index=False)['DESCRICAO'].first()
)

# ─────────────────────────────────────────────
# 4. TABELA DETALHADA (por CODIGO+DEPARTAMENTO)
#    Apenas códigos com demanda (Jan-Mai ou Jun-Dez > 0)
# ─────────────────────────────────────────────
det = (
    dem_hist
    .merge(dem_fut, on=['CODIGO','DEPARTAMENTO'], how='outer')
    .fillna({'DEMANDA_JAN_MAI':0,'DEMANDA_JUN_DEZ':0})
)
det = det[(det['DEMANDA_JAN_MAI'] > 0) | (det['DEMANDA_JUN_DEZ'] > 0)].copy()

det = det.merge(descricao, on='CODIGO', how='left')
det = det.merge(consumo,   on=['CODIGO','DEPARTAMENTO'], how='left').fillna({'CONSUMO_REAL_JAN_MAI':0})
det = det.merge(saldo,     on='CODIGO', how='left').fillna({'SALDO_ATUAL':0})

def cobertura_meses(saldo_val, dem_jun_dez):
    if dem_jun_dez <= 0:
        return math.inf
    return saldo_val / (dem_jun_dez / 7)

def status_cobertura(meses):
    if math.isinf(meses) or meses >= 7:
        return 'ADEQUADO'
    elif meses >= 1:
        return 'ABAIXO'
    return 'CRÍTICO'

def status_consumo(real, demanda):
    if demanda <= 0:
        return 'Sem Demanda Planejada'
    var = (real - demanda) / demanda * 100
    if var > 10:
        return 'Acima do Plano'
    elif var < -10:
        return 'Abaixo do Plano'
    return 'Conforme Plano'

det['CONSUMO_MES_MEDIO_JUN_DEZ'] = det['DEMANDA_JUN_DEZ'] / 7
det['COBERTURA_MESES'] = det.apply(
    lambda r: cobertura_meses(r['SALDO_ATUAL'], r['DEMANDA_JUN_DEZ']), axis=1)
det['STATUS_COBERTURA'] = det['COBERTURA_MESES'].apply(status_cobertura)
det['STATUS_CONSUMO']   = det.apply(
    lambda r: status_consumo(r['CONSUMO_REAL_JAN_MAI'], r['DEMANDA_JAN_MAI']), axis=1)

# ─────────────────────────────────────────────
# 5. TABELA POR DEPARTAMENTO
# ─────────────────────────────────────────────
depto_agg = (
    det.groupby('DEPARTAMENTO', as_index=False).agg(
        NUM_CODIGOS=('CODIGO','nunique'),
        CONSUMO_REAL_JAN_MAI=('CONSUMO_REAL_JAN_MAI','sum'),
        DEMANDA_JAN_MAI=('DEMANDA_JAN_MAI','sum'),
        DEMANDA_JUN_DEZ=('DEMANDA_JUN_DEZ','sum'),
        SALDO_TOTAL=('SALDO_ATUAL','sum'),
    )
)
depto_agg['VARIANCIA_%'] = depto_agg.apply(
    lambda r: (r['CONSUMO_REAL_JAN_MAI']-r['DEMANDA_JAN_MAI'])/r['DEMANDA_JAN_MAI']*100
    if r['DEMANDA_JAN_MAI'] > 0 else 0, axis=1)
depto_agg['CONSUMO_MES_MEDIO'] = depto_agg['CONSUMO_REAL_JAN_MAI'] / 5
depto_agg['COBERTURA_MESES'] = depto_agg.apply(
    lambda r: cobertura_meses(r['SALDO_TOTAL'], r['DEMANDA_JUN_DEZ']), axis=1)
depto_agg['STATUS_COBERTURA'] = depto_agg['COBERTURA_MESES'].apply(status_cobertura)

# ─────────────────────────────────────────────
# 6. TABELA POR PROGRAMA ORÇAMENTÁRIO
#    Demanda por programa; consumo real alocado proporcionalmente
# ─────────────────────────────────────────────
prog_hist = (
    dem[dem['MES'].isin(JAN_MAI)]
    .groupby(['PROGRAMA_ORC','CODIGO','DEPARTAMENTO'], as_index=False)['QTD'].sum()
    .rename(columns={'QTD':'DEM_PROG_JAN_MAI'})
)
prog_fut = (
    dem[dem['MES'].isin(JUN_DEZ)]
    .groupby(['PROGRAMA_ORC','CODIGO','DEPARTAMENTO'], as_index=False)['QTD'].sum()
    .rename(columns={'QTD':'DEM_PROG_JUN_DEZ'})
)
prog_det = (
    prog_hist
    .merge(prog_fut, on=['PROGRAMA_ORC','CODIGO','DEPARTAMENTO'], how='outer')
    .fillna({'DEM_PROG_JAN_MAI':0,'DEM_PROG_JUN_DEZ':0})
)
# só codes com demanda
prog_det = prog_det[(prog_det['DEM_PROG_JAN_MAI']>0)|(prog_det['DEM_PROG_JUN_DEZ']>0)].copy()

# Calcular proporção de demanda do programa no total do CODIGO+DEPARTAMENTO
total_por_codigo_depto_hist = (
    prog_det.groupby(['CODIGO','DEPARTAMENTO'])['DEM_PROG_JAN_MAI'].sum()
    .reset_index().rename(columns={'DEM_PROG_JAN_MAI':'TOTAL_DEM_JAN_MAI'})
)
prog_det = prog_det.merge(total_por_codigo_depto_hist, on=['CODIGO','DEPARTAMENTO'])
prog_det['PROP'] = prog_det.apply(
    lambda r: r['DEM_PROG_JAN_MAI']/r['TOTAL_DEM_JAN_MAI'] if r['TOTAL_DEM_JAN_MAI']>0 else 0, axis=1)

# Juntar consumo real e alocar proporcionalmente
prog_det = prog_det.merge(consumo, on=['CODIGO','DEPARTAMENTO'], how='left').fillna({'CONSUMO_REAL_JAN_MAI':0})
prog_det['CONS_REAL_PROG'] = prog_det['CONSUMO_REAL_JAN_MAI'] * prog_det['PROP']

prog_agg = (
    prog_det.groupby('PROGRAMA_ORC', as_index=False).agg(
        NUM_CODIGOS=('CODIGO','nunique'),
        CONSUMO_REAL_JAN_MAI=('CONS_REAL_PROG','sum'),
        DEMANDA_JAN_MAI=('DEM_PROG_JAN_MAI','sum'),
        DEMANDA_JUN_DEZ=('DEM_PROG_JUN_DEZ','sum'),
    )
)
# Saldo total por programa (soma dos saldos dos códigos do programa)
saldo_prog = (
    prog_det[['PROGRAMA_ORC','CODIGO']].drop_duplicates()
    .merge(saldo, on='CODIGO', how='left').fillna({'SALDO_ATUAL':0})
    .groupby('PROGRAMA_ORC', as_index=False)['SALDO_ATUAL'].sum()
    .rename(columns={'SALDO_ATUAL':'SALDO_TOTAL'})
)
prog_agg = prog_agg.merge(saldo_prog, on='PROGRAMA_ORC', how='left').fillna({'SALDO_TOTAL':0})

prog_agg['VARIANCIA_%'] = prog_agg.apply(
    lambda r: (r['CONSUMO_REAL_JAN_MAI']-r['DEMANDA_JAN_MAI'])/r['DEMANDA_JAN_MAI']*100
    if r['DEMANDA_JAN_MAI'] > 0 else 0, axis=1)
prog_agg['CONSUMO_MES_MEDIO'] = prog_agg['CONSUMO_REAL_JAN_MAI'] / 5
prog_agg['COBERTURA_MESES'] = prog_agg.apply(
    lambda r: cobertura_meses(r['SALDO_TOTAL'], r['DEMANDA_JUN_DEZ']), axis=1)
prog_agg['STATUS_COBERTURA'] = prog_agg['COBERTURA_MESES'].apply(status_cobertura)

# ─────────────────────────────────────────────
# 7. SERIALIZAR PARA JSON (tratar Infinity)
# ─────────────────────────────────────────────
def to_records(df):
    rows = []
    for rec in df.to_dict('records'):
        clean = {}
        for k, v in rec.items():
            if isinstance(v, float) and math.isinf(v):
                clean[k] = 'Infinity'
            elif isinstance(v, float) and math.isnan(v):
                clean[k] = 0
            elif isinstance(v, (np.integer,)):
                clean[k] = int(v)
            elif isinstance(v, (np.floating,)):
                clean[k] = float(v)
            else:
                clean[k] = v
        rows.append(clean)
    return rows

dados = {
    'detalhado':    to_records(det[[
        'CODIGO','DEPARTAMENTO','DESCRICAO',
        'DEMANDA_JAN_MAI','CONSUMO_REAL_JAN_MAI','DEMANDA_JUN_DEZ',
        'SALDO_ATUAL','CONSUMO_MES_MEDIO_JUN_DEZ',
        'COBERTURA_MESES','STATUS_COBERTURA','STATUS_CONSUMO'
    ]].sort_values(['DEPARTAMENTO','CODIGO'])),
    'departamentos': to_records(depto_agg.sort_values('DEPARTAMENTO')),
    'programas':    to_records(prog_agg.sort_values('PROGRAMA_ORC')),
    'filtros': {
        'departamentos': sorted(det['DEPARTAMENTO'].dropna().unique().tolist()),
        'programas':     sorted(dem['PROGRAMA_ORC'].dropna().unique().tolist()),
    }
}

# Substituir string 'Infinity' por literal JS Infinity
json_str = json.dumps(dados, ensure_ascii=False)
json_str = json_str.replace('"Infinity"', 'Infinity')

# ─────────────────────────────────────────────
# 8. INJETAR NO HTML
# ─────────────────────────────────────────────
template_path = BASE / 'dashboard_analise_consumo_melhorado.html'
html = template_path.read_text(encoding='utf-8')

# Substituir o bloco dadosCompletos
html = re.sub(
    r'let dadosCompletos\s*=\s*\{.*?\};',
    f'let dadosCompletos = {json_str};',
    html,
    count=1,
    flags=re.DOTALL
)

out_path = BASE / 'dashboard_analise_consumo_melhorado.html'
out_path.write_text(html, encoding='utf-8')

print("✓ HTML gerado:", out_path)
print(f"  Detalhado : {len(dados['detalhado'])} linhas (CODIGO+DEPTO)")
print(f"  Deptos    : {len(dados['departamentos'])}")
print(f"  Programas : {len(dados['programas'])}")
print(f"\nVerificação 400072:")
rows_400072 = [r for r in dados['detalhado'] if r['CODIGO']==400072]
for r in rows_400072:
    print(f"  {r['DEPARTAMENTO']}: Real={r['CONSUMO_REAL_JAN_MAI']:.0f}, "
          f"Dem Jan-Mai={r['DEMANDA_JAN_MAI']:.0f}, "
          f"Dem Jun-Dez={r['DEMANDA_JUN_DEZ']:.0f}, "
          f"Saldo={r['SALDO_ATUAL']:.0f}, "
          f"Cobertura={r['COBERTURA_MESES']}, Status={r['STATUS_COBERTURA']}")
