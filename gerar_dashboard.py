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
# Montante financeiro (formato BR: ponto=milhar, vírgula=decimal)
mb51['MONTANTE_NUM'] = (mb51['MONTANTE'].astype(str)
    .str.replace('.','', regex=False).str.replace(',','.', regex=False))
mb51['MONTANTE_NUM'] = pd.to_numeric(mb51['MONTANTE_NUM'], errors='coerce').fillna(0)

# Valor financeiro retirado por departamento (soma algébrica → negativo = saída)
valor_depto_mb51 = (
    mb51.groupby('DEPARTAMENTO', as_index=False)['MONTANTE_NUM'].sum()
    .rename(columns={'MONTANTE_NUM':'NET_MONTANTE'})
)
valor_depto_mb51['VALOR_RETIRADO'] = (-valor_depto_mb51['NET_MONTANTE']).clip(lower=0)

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

# Quantidade necessária para pedido (elevar ao estoque adequado = 7 meses)
# Taxa base: consumo histórico médio/mês; se zero, usa demanda planejada/mês
det['TAXA_BASE'] = det.apply(
    lambda r: r['CONSUMO_REAL_JAN_MAI'] / 5 if r['CONSUMO_REAL_JAN_MAI'] > 0
              else r['DEMANDA_JAN_MAI'] / 5 if r['DEMANDA_JAN_MAI'] > 0
              else r['DEMANDA_JUN_DEZ'] / 7, axis=1)
det['ESTOQUE_SEGURO'] = (det['TAXA_BASE'] * 7).round(0)
# QTD_NECESSARIA = meta - (saldo atual + pedidos em aberto); apenas se positivo
# Preenchido depois do merge com ME2L (usa QTD_PEDIDO)

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

# Contagem de itens por status de cobertura (CODIGO+DEPARTAMENTO do detalhado)
status_depto = (
    det.groupby(['DEPARTAMENTO','STATUS_COBERTURA']).size()
    .unstack(fill_value=0).reset_index()
)
for col in ['CRÍTICO','ABAIXO','ADEQUADO']:
    if col not in status_depto.columns:
        status_depto[col] = 0
depto_agg = depto_agg.merge(
    status_depto[['DEPARTAMENTO','CRÍTICO','ABAIXO','ADEQUADO']],
    on='DEPARTAMENTO', how='left'
).fillna({'CRÍTICO':0,'ABAIXO':0,'ADEQUADO':0})
depto_agg[['CRÍTICO','ABAIXO','ADEQUADO']] = depto_agg[['CRÍTICO','ABAIXO','ADEQUADO']].astype(int)

# Valor financeiro retirado por departamento (do MB51)
depto_agg = depto_agg.merge(
    valor_depto_mb51[['DEPARTAMENTO','VALOR_RETIRADO']], on='DEPARTAMENTO', how='left'
).fillna({'VALOR_RETIRADO': 0})

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

# Contagem de itens por status de cobertura por programa
# Juntar status do detalhado (CODIGO+DEPTO) com os itens do programa
prog_status = (
    prog_det[['PROGRAMA_ORC','CODIGO','DEPARTAMENTO']].drop_duplicates()
    .merge(det[['CODIGO','DEPARTAMENTO','STATUS_COBERTURA']], on=['CODIGO','DEPARTAMENTO'], how='left')
)
status_prog = (
    prog_status.groupby(['PROGRAMA_ORC','STATUS_COBERTURA']).size()
    .unstack(fill_value=0).reset_index()
)
for col in ['CRÍTICO','ABAIXO','ADEQUADO']:
    if col not in status_prog.columns:
        status_prog[col] = 0
prog_agg = prog_agg.merge(
    status_prog[['PROGRAMA_ORC','CRÍTICO','ABAIXO','ADEQUADO']],
    on='PROGRAMA_ORC', how='left'
).fillna({'CRÍTICO':0,'ABAIXO':0,'ADEQUADO':0})
prog_agg[['CRÍTICO','ABAIXO','ADEQUADO']] = prog_agg[['CRÍTICO','ABAIXO','ADEQUADO']].astype(int)

# ─────────────────────────────────────────────
# 7. ME2L – Pedidos em aberto
# ─────────────────────────────────────────────
me2l = pd.read_excel(BASE / 'ME2L.xlsx')
me2l['Material'] = pd.to_numeric(me2l['Material'], errors='coerce')
me2l = me2l.dropna(subset=['Material'])
me2l['Material'] = me2l['Material'].astype(int)

# Apenas linhas com quantidade pendente
me2l_aberto = me2l[me2l['a ser fornecida (quantidade)'] > 0].copy()

# Formatar datas como string para JSON
for col in ['Data do documento', 'Data de remessa']:
    me2l_aberto[col] = pd.to_datetime(me2l_aberto[col], errors='coerce').dt.strftime('%d/%m/%Y')

me2l_aberto['Contrato básico'] = me2l_aberto['Contrato básico'].fillna(0).astype(int)
me2l_aberto['Requisição de compra'] = me2l_aberto['Requisição de compra'].fillna(0).astype(int)

# Totalizar por código para injetar na tabela detalhada
pedido_por_codigo = (
    me2l_aberto.groupby('Material', as_index=False)
    .agg(QTD_PEDIDO=('a ser fornecida (quantidade)', 'sum'),
         VALOR_PEDIDO=('a ser fornecido (valor', 'sum'),
         PROXIMA_ENTREGA=('Data de remessa', 'min'))
    .rename(columns={'Material': 'CODIGO'})
)

# Enriquecer detalhado com qtd em pedido e cobertura recalculada
det = det.merge(pedido_por_codigo[['CODIGO','QTD_PEDIDO','VALOR_PEDIDO','PROXIMA_ENTREGA']],
                on='CODIGO', how='left').fillna({'QTD_PEDIDO': 0, 'VALOR_PEDIDO': 0, 'PROXIMA_ENTREGA': ''})

det['COBERTURA_COM_PEDIDO'] = det.apply(
    lambda r: cobertura_meses(r['SALDO_ATUAL'] + r['QTD_PEDIDO'], r['DEMANDA_JUN_DEZ']), axis=1)
det['STATUS_COB_COM_PEDIDO'] = det['COBERTURA_COM_PEDIDO'].apply(status_cobertura)

# QTD necessária para pedido: estoque seguro (7 meses) menos o que já existe+pedido
det['QTD_NECESSARIA_PEDIDO'] = (
    det['ESTOQUE_SEGURO'] - det['SALDO_ATUAL'] - det['QTD_PEDIDO']
).clip(lower=0).round(0)

# Lista completa de pedidos para a nova aba
pedidos_records = me2l_aberto[[
    'Material', 'Texto breve', 'Documento de compras', 'Requisição de compra',
    'Contrato básico', 'Fornecedor/centro fornecedor',
    'Data do documento', 'Data de remessa',
    'a ser fornecida (quantidade)', 'a ser fornecido (valor'
]].rename(columns={
    'Material':                       'CODIGO',
    'Texto breve':                    'DESCRICAO',
    'Documento de compras':           'PEDIDO_COMPRA',
    'Requisição de compra':           'REQUISICAO',
    'Contrato básico':                'CONTRATO',
    'Fornecedor/centro fornecedor':   'FORNECEDOR',
    'Data do documento':              'DATA_DOCUMENTO',
    'Data de remessa':                'DATA_REMESSA',
    'a ser fornecida (quantidade)':   'QTD_A_RECEBER',
    'a ser fornecido (valor':         'VALOR',
}).sort_values(['DATA_REMESSA','CODIGO'])

# ─────────────────────────────────────────────
# 8. SERIALIZAR PARA JSON (tratar Infinity)
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
    'detalhado': to_records(det[[
        'CODIGO','DEPARTAMENTO','DESCRICAO',
        'DEMANDA_JAN_MAI','CONSUMO_REAL_JAN_MAI','DEMANDA_JUN_DEZ',
        'SALDO_ATUAL','CONSUMO_MES_MEDIO_JUN_DEZ',
        'COBERTURA_MESES','STATUS_COBERTURA','STATUS_CONSUMO',
        'QTD_PEDIDO','VALOR_PEDIDO','PROXIMA_ENTREGA',
        'COBERTURA_COM_PEDIDO','STATUS_COB_COM_PEDIDO',
        'QTD_NECESSARIA_PEDIDO','TAXA_BASE','ESTOQUE_SEGURO'
    ]].sort_values(['DEPARTAMENTO','CODIGO'])),
    'departamentos': to_records(depto_agg.sort_values('DEPARTAMENTO')),
    'programas':     to_records(prog_agg.sort_values('PROGRAMA_ORC')),
    'pedidos':       to_records(pedidos_records),
    'filtros': {
        'departamentos': sorted(det['DEPARTAMENTO'].dropna().unique().tolist()),
        'programas':     sorted(dem['PROGRAMA_ORC'].dropna().unique().tolist()),
        # lookup: programa → lista de "CODIGO|DEPARTAMENTO" para filtrar detalhado
        'prog_to_codigos': {
            prog: (
                dem[dem['PROGRAMA_ORC']==prog][['CODIGO','DEPARTAMENTO']]
                .drop_duplicates()
                .apply(lambda r: f"{r['CODIGO']}|{r['DEPARTAMENTO']}", axis=1)
                .tolist()
            )
            for prog in dem['PROGRAMA_ORC'].dropna().unique()
        },
        # KPI: itens distintos demandados (unique CODIGOs)
        'total_codigos_distintos': int(det['CODIGO'].nunique()),
        # ABAIXO sub-período
        'abaixo_1_2':  int(((det['COBERTURA_MESES'] >= 1) & (det['COBERTURA_MESES'] < 2)).sum()),
        'abaixo_2_4':  int(((det['COBERTURA_MESES'] >= 2) & (det['COBERTURA_MESES'] < 4)).sum()),
        'abaixo_4_7':  int(((det['COBERTURA_MESES'] >= 4) & (det['COBERTURA_MESES'] < 7)).sum()),
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
print(f"  Pedidos   : {len(dados['pedidos'])} linhas em aberto")
print(f"\nVerificação 400072:")
rows_400072 = [r for r in dados['detalhado'] if r['CODIGO']==400072]
for r in rows_400072:
    print(f"  {r['DEPARTAMENTO']}: Real={r['CONSUMO_REAL_JAN_MAI']:.0f}, "
          f"Saldo={r['SALDO_ATUAL']:.0f}, Pedido={r['QTD_PEDIDO']:.0f}, "
          f"Cob.={r['COBERTURA_MESES']:.2f} → c/pedido={r['COBERTURA_COM_PEDIDO']:.2f} ({r['STATUS_COB_COM_PEDIDO']})")
