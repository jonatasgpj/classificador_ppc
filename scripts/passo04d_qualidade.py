import collections
import csv
import json
import os
import re
import sys

AQUI = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import revisao_comum as rc

REV = os.path.join(AQUI, 'resultado', 'revisao')
XLSX = os.path.join(AQUI, 'resultado', 'Dados_Unificados.xlsx')
MOTIVOS_SEM_ASSINATURA = ('fora_do_ppc', 'duplicata_extenso')
RX_REFERENCIA = re.compile(r'\bisbn\b|\bissn\b|\beditora\b|\bed\.|\bet al\.|\b\d{1,2}\s?(a|ª)\.? ed(i[cç][aã]o)?\b|\bp\. ?\d{1,4}\b', re.I)
RX_AUTOR = re.compile(r'\b[A-ZÀ-Ú][A-ZÀ-Ú\'-]{2,}, [A-ZÀ-Ú]\.|\b[A-Za-zÀ-ú-]{3,}, [A-Za-zÀ-ú]\.\s?[A-Za-zÀ-ú]?\.?')


def classe_de(ev):
    d = ev['decisao']
    if d == 'descarta':
        return ('descarte', None)
    if d in ('conta', 'sem_topico'):
        return (ev['nivel'], int(ev['id_categoria']) if d == 'conta' else None)
    return ('duvida', None)


def checagens(itens, real):
    por_chave = {i['chave']: i for i in itens}
    flags = collections.defaultdict(list)
    grupos = collections.defaultdict(list)
    for chave, ev in real.items():
        if ev['decisao'] == 'duvida' or (ev['decisao'] == 'descarta' and ev.get('motivo') in MOTIVOS_SEM_ASSINATURA):
            continue
        grupos[por_chave[chave]['assinatura']].append(chave)
    for chaves in grupos.values():
        if len(chaves) < 2:
            continue
        niveis = {classe_de(real[c])[0] for c in chaves}
        if len(niveis) > 1:
            for c in chaves:
                flags[c].append(('assinatura_igual_niveis_diferentes', 'mesmo trecho decidido como: %s' % ', '.join(sorted(niveis))))
        else:
            topicos = {classe_de(real[c])[1] for c in chaves if real[c]['decisao'] == 'conta'}
            if len(topicos) > 1:
                for c in chaves:
                    flags[c].append(('assinatura_igual_topicos_diferentes', 'mesmo trecho com tópicos: %s' % ', '.join(str(t) for t in sorted(topicos))))
    paginas = collections.defaultdict(list)
    for chave, ev in real.items():
        if ev['decisao'] in ('conta', 'sem_topico'):
            i = por_chave[chave]
            paginas[(i['codigo_curso'], i['pagina'])].append(chave)
    for (cc, pg), chaves in paginas.items():
        niveis = {real[c]['nivel'] for c in chaves}
        if len(niveis) > 1:
            for c in chaves:
                flags[c].append(('pagina_com_niveis_misturados', '%s p.%s: %s' % (cc, pg, ', '.join(sorted(niveis)))))
    for chave, ev in real.items():
        trecho = por_chave[chave]['trecho']
        if ev['decisao'] == 'conta' and RX_REFERENCIA.search(trecho):
            flags[chave].append(('conta_com_marca_de_referencia', 'o trecho tem ISBN, editora, edição ou "et al."'))
        if ev['decisao'] == 'descarta' and ev.get('motivo') == 'bibliografia' and not (RX_REFERENCIA.search(trecho) or RX_AUTOR.search(trecho)):
            flags[chave].append(('descarte_bibliografia_sem_marcador', 'o trecho não tem marca de referência'))
    return dict(flags)


def executar():
    itens = json.load(open(os.path.join(REV, 'fila.json'), encoding='utf-8'))
    try:
        real, religadas, orfas = rc.religar(rc.ler_tabela(), itens)
    except rc.ErroTabela as e:
        raise SystemExit('Tabela de classificação com problema: %s' % e)
    flags = checagens(itens, real)
    saida = [{'chave': c, 'motivos': [{'motivo': m, 'detalhe': d} for m, d in f]} for c, f in sorted(flags.items())]
    with open(os.path.join(REV, 'sinalizadas.json'), 'w', encoding='utf-8', newline='\n') as f:
        json.dump(saida, f, ensure_ascii=False, indent=1)
    with open(os.path.join(REV, 'sinalizadas.csv'), 'w', newline='', encoding='utf-8-sig') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(['chave', 'motivo', 'detalhe'])
        for c, fl in sorted(flags.items()):
            for m, d in fl:
                w.writerow([c, m, d])
    cont = collections.Counter(m for f in flags.values() for m, _ in f)
    sem_pdf = rc.codigos_sem_pdf(XLSX)
    de_lado = sum(1 for c in orfas if c.split('|')[0] in sem_pdf)
    print('  classificações: %d (religadas pela assinatura: %d; sem par na lista: %d; de cursos sem PDF: %d) | ocorrências sinalizadas: %d | por motivo: %s' % (
        len(real), len(religadas), len(orfas) - de_lado, de_lado, len(flags), dict(cont)))
    print('  gravado em resultado/revisao/sinalizadas.json e sinalizadas.csv')


if __name__ == '__main__':
    executar()
