import importlib.metadata as md
import os
import subprocess
import sys

KIT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REQUISITOS = os.path.join(KIT, 'requisitos.txt')
TESSERACT = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
VERSAO_TESSERACT = '5.4.0.20240606'

falta, diferente = [], []


def ok(texto):
    print('  OK         ' + texto)


versao = '%d.%d.%d' % sys.version_info[:3]
if sys.version_info[:2] == (3, 12):
    ok('Python %s (%s)' % (versao, sys.executable))
else:
    diferente.append('Python %s; o kit usa o 3.12' % versao)

with open(REQUISITOS, encoding='utf-8') as f:
    for linha in f:
        linha = linha.strip()
        if not linha or linha.startswith('#'):
            continue
        nome, esperada = linha.split('==')
        try:
            instalada = md.version(nome)
        except md.PackageNotFoundError:
            falta.append('pacote %s (versão %s)' % (nome, esperada))
            continue
        if instalada == esperada:
            ok('%s %s' % (nome, instalada))
        else:
            diferente.append('pacote %s na versão %s; o kit usa a %s' % (nome, instalada, esperada))

if not os.path.exists(TESSERACT):
    falta.append('Tesseract OCR em %s' % TESSERACT)
else:
    r = subprocess.run([TESSERACT, '--version'], capture_output=True, text=True)
    primeira = ((r.stdout or '') + (r.stderr or '')).strip().splitlines()[:1]
    lida = primeira[0].replace('tesseract', '').strip().lstrip('v') if primeira else '?'
    if lida == VERSAO_TESSERACT:
        ok('Tesseract %s' % lida)
    else:
        diferente.append('Tesseract %s; o kit usa o %s (o OCR pode sair diferente)' % (lida, VERSAO_TESSERACT))
for caminho in (os.path.join('entradas', 'ocr', 'por.traineddata'), os.path.join('entradas', 'ocr', 'configs', 'tsv')):
    if os.path.exists(os.path.join(KIT, caminho)):
        ok(caminho)
    else:
        falta.append(caminho)

pdfs = [n for n in os.listdir(os.path.join(KIT, 'entradas', 'ppc')) if n.lower().endswith('.pdf')] if os.path.isdir(os.path.join(KIT, 'entradas', 'ppc')) else []
if pdfs:
    ok('%d PDFs em entradas/ppc' % len(pdfs))
else:
    falta.append('os PDFs dos PPCs em entradas/ppc')

for t in diferente:
    print('  DIFERENTE  ' + t)
for t in falta:
    print('  FALTA      ' + t)
print()
if falta:
    print('Falta o que está marcado como FALTA: o kit não roda assim.')
elif diferente:
    print('O kit roda, mas o que está marcado como DIFERENTE pode mudar os resultados.')
else:
    print('Ambiente completo: tudo nas versões usadas para gerar os resultados.')
sys.exit(1 if falta else 0)
