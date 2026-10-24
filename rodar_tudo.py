import os
import shutil
import subprocess
import sys

KIT = os.path.dirname(os.path.abspath(__file__))
RES = os.path.join(KIT, 'resultado')
PY = [sys.executable, '-I', '-X', 'utf8']


def etapa(titulo, script, *args):
    print('\n' + '=' * 80 + '\n' + titulo + '\n' + '=' * 80, flush=True)
    r = subprocess.run(PY + [os.path.join(KIT, 'scripts', script), *args], cwd=KIT)
    if r.returncode != 0:
        sys.exit('PAROU em: %s (código %d)' % (titulo, r.returncode))


if os.path.exists(RES):
    shutil.rmtree(RES)
os.makedirs(RES)

etapa('Passo 1 - unificar o e-MEC', 'passo01_unificar_emec.py', '--recriar')
etapa('Passo 2 - PDF para texto', 'passo02_pdf_para_texto.py')
etapa('Passo 3 - buscar os termos', 'passo03_buscar_termos.py')
etapa('Passo 4A - classificação automática por regras', 'passo04a_preparar_fila.py')
etapa('Passo 4C - consolidar a classificação', 'passo04c_consolidar.py')
etapa('Passo 4D - checagens automáticas', 'passo04d_qualidade.py')
etapa('Análise do capítulo 4', 'analise_resultados.py')
print('\nPronto. Os resultados estão em resultado\\.')
