import collections
import csv
import json
import os
import statistics as st

import openpyxl

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.abspath(os.path.join(AQUI, '..'))
SAIDA = os.path.join(RAIZ, 'resultado', 'analise')
SAIDA_CSV = os.path.join(SAIDA, 'dados')
SAIDA_FIG = os.path.join(SAIDA, 'figuras')
os.makedirs(SAIDA_CSV, exist_ok=True)
os.makedirs(SAIDA_FIG, exist_ok=True)

REGIAO = {'AC': 'Norte', 'AP': 'Norte', 'AM': 'Norte', 'PA': 'Norte', 'RO': 'Norte', 'RR': 'Norte', 'TO': 'Norte',
          'AL': 'Nordeste', 'BA': 'Nordeste', 'CE': 'Nordeste', 'MA': 'Nordeste', 'PB': 'Nordeste', 'PE': 'Nordeste', 'PI': 'Nordeste', 'RN': 'Nordeste', 'SE': 'Nordeste',
          'DF': 'Centro-Oeste', 'GO': 'Centro-Oeste', 'MT': 'Centro-Oeste', 'MS': 'Centro-Oeste',
          'ES': 'Sudeste', 'MG': 'Sudeste', 'RJ': 'Sudeste', 'SP': 'Sudeste',
          'PR': 'Sul', 'RS': 'Sul', 'SC': 'Sul'}
NIVEIS = ['Macro', 'Meso', 'Micro']
CORTE_ANO = 2019


def tab(wb, nome):
    linhas = list(wb[nome].iter_rows(values_only=True))
    return [dict(zip(linhas[0], l)) for l in linhas[1:]]


def resumo(valores):
    v = sorted(valores)
    q = st.quantiles(v, n=4, method='inclusive') if len(v) > 1 else v * 3  # um PPC só: os quartis são o próprio valor
    return {'n': len(v), 'total': round(sum(v), 2), 'media': round(st.mean(v), 2), 'mediana': round(st.median(v), 2),
            'q1': round(q[0], 2), 'q3': round(q[2], 2), 'min': round(v[0], 2), 'max': round(v[-1], 2), 'dp': round(st.pstdev(v), 2)}


def grava_csv(nome, cab, linhas):
    with open(os.path.join(SAIDA_CSV, nome), 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(cab)
        w.writerows(linhas)


wb = openpyxl.load_workbook(os.path.join(RAIZ, 'resultado', 'Dados_Unificados.xlsx'), read_only=True)
CURSOS = {d['id_curso']: d for d in tab(wb, 'Dados_Unificados')}
PPC = tab(wb, 'PPC')
CAT = {c['id_categoria']: c for c in tab(wb, 'CATEGORIA')}
OCO = {o['id_ocorrencia']: o for o in tab(wb, 'OCORRENCIA')}
CLA = tab(wb, 'CLASSIFICACAO')
FATO = tab(wb, 'FATO_ANALISE')
NIVEL = tab(wb, 'NIVEL')
REL = {r['id_ppc']: r for r in tab(wb, 'RELACIONAL')}

com_pdf = [p for p in PPC if p['status_coleta'] == 1]
info = {}
for p in com_pdf:
    c = CURSOS[p['id_curso']]
    info[p['id_ppc']] = {'sigla': c['Sigla da IES'], 'campus': c['Nome do Campus'], 'uf': c['UF'], 'regiao': REGIAO[c['UF']],
                         'categoria': c['Categoria Administrativa'], 'organizacao': c['Organização Acadêmica'], 'ano': int(p['ano_ppc']),
                         'pag_validas': p['n_paginas_validas'], 'codigo': str(c['Código do Curso'])}
R = {}
CURSOS_MUN = {p['id_ppc']: CURSOS[p['id_curso']]['Município'] for p in com_pdf}

estados = collections.Counter(c['estado'] for c in CLA)
mencoes = [c for c in CLA if c['estado'] in ('conta', 'sem_topico')]
desc = [c for c in CLA if c['estado'] == 'descartada']
R['ocorrencias'] = len(CLA)
R['mencoes'] = len(mencoes)
R['conta'] = estados['conta']
R['sem_topico'] = estados['sem_topico']
R['descartadas'] = estados['descartada']
R['ppcs'] = len(com_pdf)
termo_m = collections.Counter(OCO[c['id_ocorrencia']]['termo'] for c in mencoes)
termo_o = collections.Counter(OCO[c['id_ocorrencia']]['termo'] for c in CLA)
R['termo_mencoes'] = dict(termo_m)
R['termo_ocorrencias'] = dict(termo_o)
R['termo_descartes_bibliografia'] = dict(collections.Counter(OCO[c['id_ocorrencia']]['termo'] for c in CLA if c['estado'] == 'descartada' and c['motivo'] == 'bibliografia'))

por_ppc = {i: 0 for i in info}
paginas_ppc = {i: set() for i in info}
for c in mencoes:
    o = OCO[c['id_ocorrencia']]
    por_ppc[o['id_ppc']] += 1
    paginas_ppc[o['id_ppc']].add(o['pagina'])
dens = {i: 100 * por_ppc[i] / info[i]['pag_validas'] for i in info}
R['presenca_ppcs'] = sum(1 for v in por_ppc.values() if v > 0)
R['sem_mencao'] = sorted(info[i]['sigla'] for i, v in por_ppc.items() if v == 0)
R['resumo_mencoes'] = resumo(list(por_ppc.values()))
R['resumo_paginas'] = resumo([len(s) for s in paginas_ppc.values()])
R['resumo_densidade'] = resumo(list(dens.values()))
R['densidade_agregada'] = round(100 * R['mencoes'] / sum(i['pag_validas'] for i in info.values()), 2)
R['paginas_com_mencao'] = sum(len(s) for s in paginas_ppc.values())
R['mencoes_por_pagina'] = round(R['mencoes'] / (R['paginas_com_mencao'] or 1), 2)
ordem = sorted(info, key=lambda i: -por_ppc[i])
corte5 = por_ppc[ordem[min(4, len(ordem) - 1)]]  # com menos de 5 PPCs, entram todos
top5 = [i for i in ordem if por_ppc[i] >= corte5]
R['top5_n'] = len(top5)
R['top5_mencoes'] = [{'sigla': info[i]['sigla'], 'municipio': CURSOS_MUN[i], 'campus': info[i]['campus'], 'ano': info[i]['ano'], 'mencoes': por_ppc[i], 'paginas': len(paginas_ppc[i]), 'pag_validas': info[i]['pag_validas'], 'densidade': round(dens[i], 2)} for i in top5]
R['top5_share'] = round(100 * sum(por_ppc[i] for i in top5) / (R['mencoes'] or 1), 1)
ordem_d = sorted(info, key=lambda i: -dens[i])
R['top5_densidade'] = [{'sigla': info[i]['sigla'], 'municipio': CURSOS_MUN[i], 'campus': info[i]['campus'], 'ano': info[i]['ano'], 'mencoes': por_ppc[i], 'paginas': len(paginas_ppc[i]), 'densidade': round(dens[i], 2), 'pag_validas': info[i]['pag_validas']} for i in ordem_d[:5]]
R['mediana_pag_validas'] = st.median([info[i]['pag_validas'] for i in info])
grava_csv('por_ppc.csv', ['id_ppc', 'sigla', 'campus', 'uf', 'regiao', 'categoria', 'organizacao', 'ano_ppc', 'paginas_validas', 'mencoes', 'paginas_com_mencao', 'mencoes_por_100_paginas', 'tem_macro', 'tem_meso', 'tem_micro', 'n_niveis'],
          [[i, info[i]['sigla'], info[i]['campus'], info[i]['uf'], info[i]['regiao'], info[i]['categoria'], info[i]['organizacao'], info[i]['ano'], info[i]['pag_validas'],
            por_ppc[i], len(paginas_ppc[i]), round(dens[i], 2), REL[i]['tem_macro'], REL[i]['tem_meso'], REL[i]['tem_micro'], REL[i]['n_niveis']] for i in ordem])

nivel_m = collections.Counter(c['nivel'] for c in mencoes)
R['nivel_mencoes'] = {n: nivel_m[n] for n in NIVEIS}
R['nivel_sem_topico'] = {n: sum(1 for c in mencoes if c['nivel'] == n and c['estado'] == 'sem_topico') for n in NIVEIS}
R['nivel_presenca'] = {n: sum(1 for r in REL.values() if r['tem_' + n.lower()] == 1) for n in NIVEIS}
R['nivel_paginas'] = {n: sum(r['n_paginas'] for r in NIVEL if r['nivel'] == n) for n in NIVEIS}
comb = collections.Counter()
for r in REL.values():
    chave = ' + '.join(n for n in NIVEIS if r['tem_' + n.lower()] == 1) or 'Nenhum nível'
    comb[chave] += 1
R['combinacoes'] = dict(comb.most_common())
R['n_niveis'] = dict(sorted(collections.Counter(r['n_niveis'] for r in REL.values()).items()))
ementa = collections.Counter(c['motivo'] for c in CLA if c['estado'] == 'conta' and str(c['id_categoria']) == '12')
R['ementa_tipo'] = dict(ementa)
topicos_ppc = collections.defaultdict(set)
for c in CLA:
    if c['estado'] == 'conta':
        topicos_ppc[OCO[c['id_ocorrencia']]['id_ppc']].add(int(c['id_categoria']))
R['macro_projeto_ppcs'] = sum(1 for i in info if topicos_ppc[i] & {1, 2, 3, 4, 5})
R['macro_e_micro'] = sum(1 for r in REL.values() if r['tem_macro'] == 1 and r['tem_micro'] == 1)
R['macro_com_meso_ou_micro'] = sum(1 for r in REL.values() if r['tem_macro'] == 1 and (r['tem_meso'] == 1 or r['tem_micro'] == 1))
R['macro_com_componente'] = {str(t): [sum(1 for i in info if t in topicos_ppc[i] and 7 in topicos_ppc[i]), sum(1 for i in info if t in topicos_ppc[i])] for t in range(1, 7)}
R['ementa_total'] = sum(ementa.values())

top_m = collections.Counter()
top_p = collections.Counter()
for f in FATO:
    top_m[f['id_categoria']] += f['n_mencoes']
    if f['n_mencoes'] > 0:
        top_p[f['id_categoria']] += 1
topicos = []
for k in range(1, 18):
    topicos.append({'id': k, 'nivel': CAT[k]['nivel'], 'nome': CAT[k]['nome_categoria'], 'mencoes': top_m[k], 'ppcs': top_p[k]})
R['topicos'] = topicos
R['topicos_zero'] = [t['nome'] for t in topicos if t['mencoes'] == 0]
grava_csv('por_topico.csv', ['id_categoria', 'nivel', 'topico', 'mencoes', 'ppcs_com_mencao'], [[t['id'], t['nivel'], t['nome'], t['mencoes'], t['ppcs']] for t in topicos])

pres = collections.defaultdict(set)
for f in FATO:
    if f['n_mencoes'] > 0:
        pres[f['id_ppc']].add(f['id_categoria'])
M = [[0] * 17 for _ in range(17)]
for s in pres.values():
    for a in s:
        for b in s:
            M[a - 1][b - 1] += 1
pares = []
for a in range(17):
    for b in range(a + 1, 17):
        if M[a][b] > 0:
            uniao = M[a][a] + M[b][b] - M[a][b]
            esp = M[a][a] * M[b][b] / len(info)
            pares.append({'a': topicos[a]['nome'], 'na': topicos[a]['nivel'], 'b': topicos[b]['nome'], 'nb': topicos[b]['nivel'], 'ppcs': M[a][b], 'jaccard': round(M[a][b] / uniao, 2),
                          'esperado': round(esp, 1), 'prop_a': round(M[a][b] / M[a][a], 2), 'prop_b': round(M[a][b] / M[b][b], 2), 'pa': M[a][a], 'pb': M[b][b]})
R['pares_com_coocorrencia'] = len(pares)
R['pares_possiveis'] = 17 * 16 // 2
vivos_n = sum(1 for k in range(17) if M[k][k] > 0)
R['topicos_com_mencao'] = vivos_n
R['pares_possiveis_vivos'] = vivos_n * (vivos_n - 1) // 2
R['pares_top_contagem'] = sorted(pares, key=lambda x: (-x['ppcs'], -x['jaccard']))[:6]
R['jaccard_corte'] = 5
R['pares_top_jaccard'] = sorted([p for p in pares if p['ppcs'] >= 5], key=lambda x: (-x['jaccard'], -x['ppcs']))[:6]
grava_csv('coocorrencia_topicos.csv', [''] + [t['nome'] for t in topicos], [[topicos[a]['nome']] + M[a] for a in range(17)])

mot = collections.Counter(c['motivo'] for c in desc)
R['descarte_motivos'] = dict(mot.most_common())
tri = collections.Counter()
for c in CLA:
    t = OCO[c['id_ocorrencia']]['triagem_auto']
    final = 'descartada' if c['estado'] == 'descartada' else 'mencao'
    tri[(t, final)] += 1
R['triagem_x_final'] = {'%s -> %s' % k: v for k, v in sorted(tri.items())}
R['mantidas_descartadas'] = tri[('mantida', 'descartada')]
R['mantidas_contam'] = tri[('mantida', 'mencao')]
R['triagem_descarte_contam'] = sum(v for (t, f), v in tri.items() if t != 'mantida' and f == 'mencao')
R['triagem_descarte_contam_bib'] = tri[('bibliografia', 'mencao')]
bib_conta = [c for c in CLA if OCO[c['id_ocorrencia']]['triagem_auto'] == 'bibliografia' and c['estado'] != 'descartada']
R['bibliografia_que_conta'] = dict(collections.Counter('%s %s %s' % (c['nivel'], c['id_categoria'], c['motivo'] or '') for c in bib_conta))
mot_mantidas = collections.Counter(c['motivo'] for c in desc if OCO[c['id_ocorrencia']]['triagem_auto'] == 'mantida')
R['motivos_das_mantidas_descartadas'] = dict(mot_mantidas.most_common())


def grupo(chave_fn, nome):
    g = collections.defaultdict(list)
    for i in info:
        g[chave_fn(i)].append(i)
    cursos = collections.Counter()
    for c in CURSOS.values():
        cursos[chave_fn_curso(c, nome)] += 1
    linhas = []
    for k in sorted(g, key=lambda x: -len(g[x])):
        ids = g[k]
        pequeno = len(ids) < 5
        linhas.append({'grupo': k, 'cursos': cursos[k], 'ppcs': len(ids), 'com_mencao': sum(1 for i in ids if por_ppc[i] > 0),
                       'mediana_mencoes': None if pequeno else round(st.median([por_ppc[i] for i in ids]), 1),
                       'mediana_densidade': None if pequeno else round(st.median([dens[i] for i in ids]), 2),
                       'mediana_paginas': None if pequeno else round(st.median([len(paginas_ppc[i]) for i in ids]), 1),
                       'razao_agregada': round(100 * sum(por_ppc[i] for i in ids) / sum(info[i]['pag_validas'] for i in ids), 2),
                       'casos': [{'nome': '%s (%s)' % (info[i]['sigla'], CURSOS_MUN[i]), 'mencoes': por_ppc[i], 'densidade': round(dens[i], 2),
                                  'niveis': ' + '.join(n for n in NIVEIS if REL[i]['tem_' + n.lower()] == 1)} for i in sorted(ids, key=lambda x: info[x]['sigla'])] if pequeno else [],
                       'macro': sum(REL[i]['tem_macro'] for i in ids), 'meso': sum(REL[i]['tem_meso'] for i in ids), 'micro': sum(REL[i]['tem_micro'] for i in ids),
                       'tres_niveis': sum(1 for i in ids if REL[i]['n_niveis'] == 3), 'siglas': sorted('%s (%s)' % (info[i]['sigla'], CURSOS_MUN[i]) for i in ids) if len(ids) < 5 else []})
    return linhas


def chave_fn_curso(c, nome):
    if nome == 'regiao':
        return REGIAO[c['UF']]
    if nome == 'categoria':
        return c['Categoria Administrativa']
    if nome == 'organizacao':
        return c['Organização Acadêmica']
    return None


R['recorte_regiao'] = grupo(lambda i: info[i]['regiao'], 'regiao')
R['recorte_categoria'] = grupo(lambda i: info[i]['categoria'], 'categoria')
R['recorte_organizacao'] = grupo(lambda i: info[i]['organizacao'], 'organizacao')
anos = collections.defaultdict(list)
for i in info:
    anos['até %d' % (CORTE_ANO - 1) if info[i]['ano'] < CORTE_ANO else '%d em diante' % CORTE_ANO].append(i)
R['recorte_ano'] = [{'grupo': k, 'ppcs': len(v), 'com_mencao': sum(1 for i in v if por_ppc[i] > 0), 'mediana_mencoes': round(st.median([por_ppc[i] for i in v]), 1),
                     'mediana_densidade': round(st.median([dens[i] for i in v]), 2), 'mediana_paginas': round(st.median([len(paginas_ppc[i]) for i in v]), 1),
                     'razao_agregada': round(100 * sum(por_ppc[i] for i in v) / sum(info[i]['pag_validas'] for i in v), 2), 'casos': [], 'macro': sum(REL[i]['tem_macro'] for i in v), 'meso': sum(REL[i]['tem_meso'] for i in v),
                     'micro': sum(REL[i]['tem_micro'] for i in v), 'tres_niveis': sum(1 for i in v if REL[i]['n_niveis'] == 3)} for k, v in sorted(anos.items(), key=lambda kv: kv[0] != 'até %d' % (CORTE_ANO - 1))]
R['corte_ano'] = CORTE_ANO
tres = collections.defaultdict(list)
for i in info:
    a_ = info[i]['ano']
    tres['até 2018' if a_ < 2019 else ('2019 a 2022' if a_ < 2023 else '2023 em diante')].append(i)
R['ano_tres_faixas'] = [{'grupo': k, 'ppcs': len(tres[k]), 'mediana_mencoes': round(st.median([por_ppc[i] for i in tres[k]]), 1),
                         'mediana_densidade': round(st.median([dens[i] for i in tres[k]]), 2), 'mediana_pag_validas': st.median([info[i]['pag_validas'] for i in tres[k]])}
                        for k in ('até 2018', '2019 a 2022', '2023 em diante') if tres[k]]
grava_csv('recortes.csv', ['recorte', 'grupo', 'cursos', 'ppcs', 'com_mencao', 'mediana_mencoes', 'mediana_paginas', 'mediana_densidade', 'razao_agregada', 'macro', 'meso', 'micro', 'tres_niveis'],
          [[nome, g['grupo'], g.get('cursos', ''), g['ppcs'], g['com_mencao'], g['mediana_mencoes'], g['mediana_paginas'], g['mediana_densidade'], g['razao_agregada'], g['macro'], g['meso'], g['micro'], g['tres_niveis']]
           for nome, lista in (('regiao', R['recorte_regiao']), ('categoria', R['recorte_categoria']), ('organizacao', R['recorte_organizacao']), ('ano', R['recorte_ano'])) for g in lista])
json.dump(R, open(os.path.join(SAIDA, 'resultados.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

plt.rcParams.update({'font.family': ['Arial', 'DejaVu Sans'], 'font.size': 8.5, 'axes.edgecolor': '#c9c8c3', 'axes.linewidth': 0.8,
                     'xtick.color': '#52514e', 'ytick.color': '#52514e', 'axes.labelcolor': '#52514e', 'text.color': '#0b0b0b'})
COR_NIVEL = {'Macro': '#2a78d6', 'Meso': '#eb6834', 'Micro': '#1baf7a'}
AZUL = '#2a78d6'
TINTA2 = '#52514e'
GRADE = '#e6e5e0'


def limpa(ax, eixo='x'):
    for lado in ('top', 'right'):
        ax.spines[lado].set_visible(False)
    ax.grid(axis=eixo, color=GRADE, linewidth=0.8)
    ax.set_axisbelow(True)


def salva(fig, nome):
    fig.savefig(os.path.join(SAIDA_FIG, nome), dpi=300, facecolor='white', bbox_inches='tight')
    plt.close(fig)


from matplotlib.ticker import MaxNLocator
fig, ax = plt.subplots(figsize=(6.3, 2.9))
vals = list(por_ppc.values())
bins = list(range(0, max(vals) + 6, 5))
ax.hist(vals, bins=bins, color=AZUL, rwidth=0.9)
ax.set_xticks(bins)
ax.yaxis.set_major_locator(MaxNLocator(integer=True))
ymax = ax.get_ylim()[1]
ax.set_ylim(0, ymax * 1.18)
med = R['resumo_mencoes']['mediana']
ax.axvline(med, color='#0b0b0b', linewidth=1)
ax.text(med + 0.5, ymax * 1.02, 'mediana = %g' % med, fontsize=8, color='#0b0b0b', va='center')
zeros = sum(1 for v in vals if v == 0)
ax.text(bins[-1], ymax * 1.1, 'a primeira classe inclui %d PPCs com zero menção' % zeros, fontsize=7.5, color=TINTA2, va='center', ha='right')
ax.set_xlabel('Menções por PPC (classes de 5: 0 a 4, 5 a 9, ...)')
ax.set_ylabel('Número de PPCs')
limpa(ax, 'y')
salva(fig, 'grafico_4_1_mencoes_por_ppc.png')

fig, axs = plt.subplots(1, 2, figsize=(6.3, 1.9))
for ax, chave, titulo in ((axs[0], 'nivel_mencoes', 'Menções'), (axs[1], 'nivel_presenca', 'PPCs com menção (de %d)' % R['ppcs'])):
    y = list(range(3))[::-1]
    v = [R[chave][n] for n in NIVEIS]
    ax.barh(y, v, height=0.55, color=[COR_NIVEL[n] for n in NIVEIS])
    ax.set_yticks(y)
    ax.set_yticklabels(NIVEIS)
    for yy, vv in zip(y, v):
        rot = str(vv) if chave == 'nivel_mencoes' else '%d (%s%%)' % (vv, str(round(100 * vv / R['ppcs'])))
        ax.text(vv + max(v) * 0.02, yy, rot, va='center', fontsize=8, color='#0b0b0b')
    ax.set_title(titulo, fontsize=9, color='#0b0b0b', loc='left')
    ax.set_xlim(0, max(v) * 1.32)
    limpa(ax)
fig.tight_layout()
salva(fig, 'grafico_4_2_niveis.png')

fig, ax = plt.subplots(figsize=(6.3, 2.4))
itens = list(R['combinacoes'].items())
y = list(range(len(itens)))[::-1]
ax.barh(y, [v for _, v in itens], height=0.6, color=AZUL)
ax.set_yticks(y)
ax.set_yticklabels([k for k, _ in itens])
for yy, (_, v) in zip(y, itens):
    ax.text(v + 0.5, yy, '%d (%d%%)' % (v, round(100 * v / R['ppcs'])), va='center', fontsize=8, color='#0b0b0b')
ax.set_xlabel('Número de PPCs (de %d)' % R['ppcs'])
ax.set_xlim(0, max(v for _, v in itens) * 1.15)
limpa(ax)
salva(fig, 'grafico_4_3_combinacoes_niveis.png')

fig, ax = plt.subplots(figsize=(6.3, 4.4))
ts = topicos[::-1]
y = list(range(len(ts)))
TONS = {'ementa': '#0f7a52', 'titulo_ficha': '#6fcca6', 'cabecalho_ficha': '#b3e5d0'}
for yy, t in zip(y, ts):
    if t['nome'] == 'Ementa':
        ini = 0
        for tipo in ('ementa', 'titulo_ficha', 'cabecalho_ficha'):
            w = R['ementa_tipo'].get(tipo, 0)
            ax.barh(yy, w, left=ini, height=0.62, color=TONS[tipo], edgecolor='white', linewidth=1)
            ini += w
    else:
        ax.barh(yy, t['mencoes'], height=0.62, color=COR_NIVEL[t['nivel']])
ax.set_yticks(y)
ax.set_yticklabels(['%d %s' % (t['id'], t['nome']) for t in ts])
mx = max(t['mencoes'] for t in ts)
for yy, t in zip(y, ts):
    ax.text(t['mencoes'] + mx * 0.01, yy, '%d (%d PPCs)' % (t['mencoes'], t['ppcs']), va='center', fontsize=7.5, color=TINTA2)
for corte_ in (5.5, 10.5):
    ax.axhline(corte_, color=GRADE, linewidth=0.8)
ax.set_xlim(0, mx * 1.28)
ax.set_xlabel('Menções com tópico (entre parênteses, PPCs com menção no tópico)')
from matplotlib.patches import Patch
ax.legend(handles=[Patch(color=COR_NIVEL['Macro'], label='Macro'), Patch(color=COR_NIVEL['Meso'], label='Meso'), Patch(color=COR_NIVEL['Micro'], label='Micro'),
          Patch(color=TONS['ementa'], label='Ementa: texto da ementa'), Patch(color=TONS['titulo_ficha'], label='Ementa: nome da disciplina'), Patch(color=TONS['cabecalho_ficha'], label='Ementa: cabeçalho da ficha')],
          loc='upper center', bbox_to_anchor=(0.42, -0.13), ncol=3, frameon=False, fontsize=7.5)
limpa(ax)
salva(fig, 'grafico_4_4_topicos.png')

from matplotlib.colors import LinearSegmentedColormap
import numpy as np
vivos = [k for k in range(17) if M[k][k] > 0]
nv = len(vivos)
fora = [M[a][b] for a in vivos for b in vivos if a != b]
cmap = LinearSegmentedColormap.from_list('azul', ['#e3eefc', '#9ec5f4', '#3987e5', '#1c5cab', '#0d366b'])
cmap.set_bad('white')
MM = np.full((nv, nv), np.nan)
for x, a in enumerate(vivos):
    for z, b in enumerate(vivos):
        if z < x and M[a][b] > 0:
            MM[x][z] = M[a][b]
fig, ax = plt.subplots(figsize=(6.3, 5.4))
ax.imshow(MM, cmap=cmap, vmin=1, vmax=max(fora))
for x, a in enumerate(vivos):
    ax.add_patch(plt.Rectangle((x - 0.5, x - 0.5), 1, 1, color='#e6e5e0'))
    ax.text(x, x, str(M[a][a]), ha='center', va='center', fontsize=7, color='#0b0b0b', fontweight='bold')
    for z, b in enumerate(vivos):
        if z < x:
            v = M[a][b]
            ax.text(z, x, str(v) if v else '–', ha='center', va='center', fontsize=7, color='white' if v > 30 else ('#0b0b0b' if v else '#9a998f'))
ax.set_xticks(range(nv))
ax.set_xticklabels([str(topicos[k]['id']) for k in vivos], fontsize=7.5)
ax.set_yticks(range(nv))
ax.set_yticklabels(['%d %s' % (topicos[k]['id'], topicos[k]['nome']) for k in vivos], fontsize=7.5)
for corte_ in (sum(1 for k in vivos if k < 6) - 0.5, sum(1 for k in vivos if k < 11) - 0.5):
    ax.axhline(corte_, color='white', linewidth=2)
    ax.axvline(corte_, color='white', linewidth=2)
ax.set_xlabel('Tópico (número, como no eixo vertical). Diagonal, em cinza: PPCs com o tópico.\nSem os tópicos 10, 16 e 17, que não têm menção.')
for lado in ('top', 'right', 'left', 'bottom'):
    ax.spines[lado].set_visible(False)
ax.tick_params(length=0)
salva(fig, 'grafico_4_5_coocorrencia.png')

NOME_MOT = {'bibliografia': 'Bibliografia', 'equivalencia': 'Equivalência entre currículos', 'fora_do_ppc': 'Fora do PPC', 'docente': 'Docente',
            'duplicata_extenso': 'Duplicidade do termo por extenso', 'termo_citado_norma': 'Termo citado de norma', 'infraestrutura': 'Infraestrutura',
            'sumario': 'Sumário ou índice', 'linha_pesquisa': 'Linha ou grupo de pesquisa', 'sigla_outro_sentido': 'Sigla com outro sentido',
            'outro_nao_curricular': 'Outro não curricular', 'ata': 'Ata de reunião'}
itens = list(R['descarte_motivos'].items())
fig, ax = plt.subplots(figsize=(6.3, 2.9))
y = list(range(len(itens)))[::-1]
ax.barh(y, [v for _, v in itens], height=0.62, color=AZUL)
ax.set_yticks(y)
ax.set_yticklabels([NOME_MOT.get(k, k) for k, _ in itens])
for yy, (_, v) in zip(y, itens):
    ax.text(v + 4, yy, '%d (%s%%)' % (v, str(round(100 * v / R['descartadas'], 1)).replace('.', ',')), va='center', fontsize=8, color='#0b0b0b')
ax.set_xlim(0, max(v for _, v in itens) * 1.2)
ax.set_xlabel('Ocorrências descartadas (de %d)' % R['descartadas'])
limpa(ax)
salva(fig, 'grafico_4_6_descartes.png')
print('ok: resultados.json, %d CSV e %d gráficos' % (len(os.listdir(SAIDA_CSV)), len(os.listdir(SAIDA_FIG))))
