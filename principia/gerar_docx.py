"""Converte o projeto LaTeX em .docx no modelo da Revista Principia ("Revista Principia.docx").

Uso (a partir da raiz do projeto):
  python3 principia/gerar_docx.py <saida.docx> [opções]

Caminhos de saída relativos são gravados dentro de principia/. O modelo da revista
("Revista Principia.docx") também é lido de principia/.

Lê da raiz do projeto sbcreviews-pt.tex (título, resumos, palavras-chave), refs.bib e images/;
os capítulos vêm de principia/parts/part1..6.tex (arquivos ausentes são ignorados).
Requer o pandoc (converte as equações LaTeX em equações do Word).

Opções:
  --autores            inclui autores/filiações e as seções pós-aceite
  --parts=DIR          pasta dos capítulos (padrão: principia/parts)
  --pandoc=CAMINHO     executável do pandoc (padrão: pandoc)
  --fig-max=CM         altura máxima das figuras em cm (padrão: 7)
  --nao-fundir         mantém separadas as tabelas de custo computacional e de distância
"""
import os, re, shutil, struct, subprocess, sys, tempfile, zipfile
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(HERE)
OUT = os.path.join(HERE, sys.argv[1])
OPTS = sys.argv[2:]
def opt(name, default):
    return next((o.split('=', 1)[1] for o in OPTS if o.startswith(f'--{name}=')), default)
WITH_AUTHORS = '--autores' in OPTS
FUNDIR = '--nao-fundir' not in OPTS
FIG_MAX = float(opt('fig-max', 7))
PARTS = os.path.join(PROJ, opt('parts', 'principia/parts'))
PANDOC = opt('pandoc', 'pandoc')
TPL = tempfile.mkdtemp()
with zipfile.ZipFile(os.path.join(HERE, 'Revista Principia.docx')) as z:
    z.extractall(TPL)

# ---------------------------------------------------------------- utilidades LaTeX
def strip_comments(s):
    return '\n'.join(re.sub(r'(?<!\\)%.*$', '', l) for l in s.split('\n'))

def read_group(s, i, open_='{', close='}'):
    """s[i] == open_ ; devolve (conteúdo, índice após o fechamento)."""
    assert s[i] == open_, (s[i:i+40])
    depth, j = 0, i
    while j < len(s):
        c = s[j]
        if c == '\\':
            j += 2; continue
        if c == open_: depth += 1
        elif c == close:
            depth -= 1
            if depth == 0: return s[i+1:j], j+1
        j += 1
    raise ValueError('grupo não fechado: ' + s[i:i+60])

main = strip_comments(open(os.path.join(PROJ, 'sbcreviews-pt.tex'), encoding='utf8').read())
def env(src, name):
    m = re.search(r'\\begin\{' + re.escape(name) + r'\}(.*?)\\end\{' + re.escape(name) + r'\}', src, re.S)
    return ' '.join(m.group(1).split())
TITLE_LATEX = read_group(main, main.index('{', main.index(']', main.index('\\title'))))[0]
ABS_PT, ABS_EN = env(main, 'abstract-pt'), env(main, 'abstract-en')
KW_PT, KW_EN = env(main, 'pchaves'), env(main, 'keywords')

body = ''
for k in range(1, 7):
    f = os.path.join(PARTS, f'part{k}.tex')
    if os.path.exists(f):
        body += strip_comments(open(f, encoding='utf8').read()) + '\n\n'

# Quadros (conteúdo textual) x Tabelas (conteúdo numérico) – NBR 6022 / IBGE
QUADROS = {'tab:arquiteturas-uavs', 'tab:cronograma', 'tab:cobertura'}
REPL = [(r'A Tabela~\ref{tab:arquiteturas-uavs} resume', r'O Quadro~\ref{tab:arquiteturas-uavs} resume'),
        (r'A Tabela~\ref{tab:cobertura} resume', r'O Quadro~\ref{tab:cobertura} resume'),
        (r'proposta em \citep{sholes2007evolution}', r'proposta por \citet{sholes2007evolution}')]
if 'tab:cronograma' in body:
    REPL.append((r'apresentado na Tabela~\ref{tab:cronograma}', r'apresentado no Quadro~\ref{tab:cronograma}'))
if FUNDIR:
    REPL += [(r'Os valores são apresentados na Tabela~\ref{tab:distancia}.',
              r'Os valores são apresentados na última coluna da Tabela~\ref{tab:custo}.'),
             (r'\caption{Métricas de custo computacional por algoritmo e cenário.}',
              r'\caption{Métricas de custo computacional e distância média percorrida por algoritmo e cenário.}')]
for a, b in REPL:
    assert a in body, a
    body = body.replace(a, b)

# ---------------------------------------------------------------- bibliografia
def parse_bib(path):
    s = open(path, encoding='utf8').read()
    entries = {}
    for m in re.finditer(r'@(\w+)\s*\{\s*([^,\s]+)\s*,', s):
        typ, key = m.group(1).lower(), m.group(2)
        content, _ = read_group(s, s.index('{', m.start()))
        fields = {}
        j = content.index(',') + 1
        while True:
            fm = re.compile(r'\s*(\w+)\s*=\s*').match(content, j)
            if not fm: break
            j = fm.end()
            if content[j] == '{':
                val, j = read_group(content, j)
            else:
                vm = re.compile(r'[^,]*').match(content, j); val = vm.group(0); j = vm.end()
            fields[fm.group(1).lower()] = ' '.join(val.split())
            j = re.compile(r'\s*,?').match(content, j).end()
        entries[key.lower()] = dict(fields, _type=typ, _key=key)
    return entries

BIB = parse_bib(os.path.join(PROJ, 'refs.bib'))

def unlatex(t):
    t = t.replace('{\\"u}', 'ü').replace('\\&', '&').replace('{-}', '-').replace('--', '-')
    return re.sub(r'[{}]', '', t).strip()

def parse_authors(a):
    out, others = [], False
    for p in re.split(r'\s+and\s+', a):
        p = p.strip()
        if p == 'others': others = True; continue
        if p.startswith('{') and p.endswith('}'):
            out.append({'org': unlatex(p)}); continue
        p = unlatex(p)
        if ',' in p:
            last, first = [x.strip() for x in p.split(',', 1)]
        else:
            toks = p.split(); last, first = toks[-1], ' '.join(toks[:-1])
        ini = []
        for tok in first.split():
            if re.fullmatch(r'[A-Z]{2,3}', tok):          # "FM" -> F. M.
                ini += [c + '.' for c in tok]
            else:
                ini.append('-'.join(x[0] + '.' for x in tok.split('-') if x))
        out.append({'last': last, 'ini': ' '.join(ini)})
    return out, others

for e in BIB.values():
    e['_authors'], e['_others'] = parse_authors(e.get('author', ''))

def cite_name(e):
    au = e['_authors']
    if 'org' in au[0]: return [(au[0]['org'], False)]
    if len(au) >= 4 or e['_others']:
        return [(au[0]['last'] + ' ', False), ('et al.', True)]
    return [('; '.join(x['last'] for x in au), False)]

# ---------------------------------------------------------------- texto -> runs
# run = ('t', texto, negrito, itálico) | ('m', idx_math) | ('c', keys, narrativa, negrito, itálico)
MATH = []          # (latex, display)
def add_math(tex, display=False):
    MATH.append((tex, display)); return len(MATH) - 1

LABELS = {}        # label -> texto do número
CITED = []         # ordem de citação

def text_fix(t):
    t = t.replace('``', '“').replace("''", '”').replace('---', '—').replace('--', '–')
    t = t.replace('~', '\u00a0').replace('\\%', '%').replace('\\&', '&').replace('\\_', '_')
    t = t.replace('\\ldots', '…').replace('\\checkmark', '✓').replace('\\,', '\u2009')
    return t

def inline(s, b=False, i=False):
    runs, buf, k = [], '', 0
    def flush():
        nonlocal buf
        if buf: runs.append(('t', text_fix(buf), b, i)); buf = ''
    while k < len(s):
        c = s[k]
        if c == '$':
            flush(); e = s.index('$', k + 1)
            runs.append(('m', add_math(s[k+1:e]))); k = e + 1; continue
        if s.startswith('\\(', k):
            flush(); e = s.index('\\)', k)
            runs.append(('m', add_math(s[k+2:e]))); k = e + 2; continue
        if c == '\\':
            m = re.compile(r'\\([A-Za-z]+)\*?').match(s, k)
            if m:
                cmd = m.group(1); j = m.end()
                if cmd in ('textit', 'emph', 'textbf', 'texttt'):
                    flush(); arg, k = read_group(s, j)
                    runs += inline(arg, b or cmd == 'textbf', (not i) if cmd in ('textit', 'emph') else i)
                    continue
                if cmd in ('citep', 'citet'):
                    flush(); arg, k = read_group(s, j)
                    keys = [x.strip() for x in arg.split(',')]
                    for x in keys:
                        if x.lower() not in CITED: CITED.append(x.lower())
                    runs.append(('c', keys, cmd == 'citet', b, i)); continue
                if cmd == 'ref':
                    flush(); arg, k = read_group(s, j)
                    runs.append(('t', LABELS[arg], b, i)); continue
                if cmd == 'label':
                    _, k = read_group(s, j); continue
                if cmd in ('checkmark', 'ldots'):
                    buf += '\\' + cmd; k = j
                    continue
            buf += s[k:k+2]; k += 2; continue
        if c in '{}':
            k += 1; continue
        buf += c; k += 1
    flush()
    return runs

# ---------------------------------------------------------------- blocos
HEAD_PT = {
    'Fundamentação Teórica': 'Fundamentação teórica', 'Conceitos Básicos': 'Conceitos básicos',
    'VANTs e Níveis de Autonomia': 'VANTs e níveis de autonomia',
    'Sistemas Multiagente e Cooperação': 'Sistemas multiagente e cooperação',
    'Inteligência de Enxame': 'Inteligência de enxame',
    'Arquiteturas de Controle para Enxames de VANTs': 'Arquiteturas de controle para enxames de VANTs',
    'Controle Centralizado': 'Controle centralizado',
    'Controle Descentralizado e Distribuído': 'Controle descentralizado e distribuído',
    'Comparação entre Arquiteturas': 'Comparação entre arquiteturas',
    'Algoritmos de Formação de VANTs': 'Algoritmos de formação de VANTs',
    'Comparação Geral': 'Comparação geral',
    'Comunicação e Simulação de Enxames de VANTs': 'Comunicação e simulação de enxames de VANTs',
    'Comunicação e Coordenação em Enxames de VANTs': 'Comunicação e coordenação em enxames de VANTs',
    'Protocolos de VANTs e Arquiteturas de Comunicação': 'Protocolos de VANTs e arquiteturas de comunicação',
    'Limitações e Desafios de Comunicação': 'Limitações e desafios de comunicação',
    'Simulação de Enxames de VANTs': 'Simulação de enxames de VANTs',
    'Plataformas e Motores de Simulação': 'Plataformas e motores de simulação',
    'Comunicação e Modelagem em Simulação': 'Comunicação e modelagem em simulação',
    'Síntese e Direcionamento do Trabalho': 'Síntese e direcionamento do trabalho',
    'Cobertura dos Experimentos por Cenário': 'Cobertura dos experimentos por cenário',
    'Tempo de Convergência da Missão': 'Tempo de convergência da missão',
    'Custo Computacional': 'Custo computacional',
    'Distância Percorrida pelo Enxame': 'Distância percorrida pelo enxame',
    'Segurança em Relação a Obstáculos': 'Segurança em relação a obstáculos',
    'Análise Comparativa Geral': 'Análise comparativa geral',
    'Conclusão e Trabalhos Futuros': 'Conclusão e trabalhos futuros',
}

# pré-passagem: numeração de seções, figuras, tabelas/quadros, equações
sec = [0, 0, 0]; nfig = ntab = nquad = 0
tok_re = re.compile(r'\\(section|subsection|subsubsection)\*?\{|\\begin\{(figure\*?|table\*?)\}|\\label\{([^}]*)\}')
last = None
for m in tok_re.finditer(body):
    if m.group(1):
        lvl = ['section', 'subsection', 'subsubsection'].index(m.group(1))
        sec[lvl] += 1
        for z in range(lvl + 1, 3): sec[z] = 0
        last = ('sec', '.'.join(str(x) for x in sec[:lvl+1]))
    elif m.group(2):
        if m.group(2).startswith('figure'):
            nfig += 1; last = ('fig', str(nfig))
        else:
            last = ('tab', None)
    else:
        lab = m.group(3)
        if last and last[0] == 'tab':
            if FUNDIR and lab == 'tab:distancia':
                LABELS[lab] = LABELS['tab:custo']
            elif lab in QUADROS: nquad += 1; LABELS[lab] = str(nquad)
            else: ntab += 1; LABELS[lab] = str(ntab)
        elif last:
            LABELS[lab] = last[1]

blocks = []
sec = [0, 0, 0]
pending_runin = None
pos = 0
block_re = re.compile(r'\\(section|subsection|subsubsection|paragraph)(\*?)\{|\\begin\{(figure\*?|table\*?|equation|itemize|enumerate)\}')

def emit_text(chunk):
    global pending_runin
    for para in re.split(r'\n\s*\n', chunk):
        t = ' '.join(para.split())
        if not t: continue
        runs = inline(t)
        if pending_runin:
            runs = [('t', pending_runin + ' ', True, False)] + runs; pending_runin = None
        blocks.append(('p', runs, t[0].islower()))

def cell_split(row):
    out, depth, cur, k = [], 0, '', 0
    while k < len(row):
        c = row[k]
        if c == '\\': cur += row[k:k+2]; k += 2; continue
        if c == '{': depth += 1
        if c == '}': depth -= 1
        if c == '&' and depth == 0: out.append(cur); cur = ''
        else: cur += c
        k += 1
    out.append(cur)
    return [x.strip() for x in out]

TABLES = {}
while True:
    m = block_re.search(body, pos)
    if not m:
        emit_text(body[pos:]); break
    emit_text(body[pos:m.start()])
    if m.group(1):
        title, pos = read_group(body, m.end() - 1)
        if m.group(1) == 'paragraph':
            pending_runin = title.rstrip('.') + '.'; continue
        lvl = ['section', 'subsection', 'subsubsection'].index(m.group(1))
        sec[lvl] += 1
        for z in range(lvl + 1, 3): sec[z] = 0
        num = '.'.join(str(x) for x in sec[:lvl+1])
        blocks.append(('h', num + ' ' + HEAD_PT.get(title, title)))
        continue
    name = m.group(3)
    end = body.index('\\end{' + name + '}', m.end())
    inner = body[m.end():end]
    pos = end + len('\\end{' + name + '}')
    if name.startswith('figure'):
        g = re.search(r'\\includegraphics\[width=([\d.]+)\\(?:linewidth|textwidth)\]\{([^}]*)\}', inner)
        ci = inner.index('\\caption'); cap, _ = read_group(inner, inner.index('{', ci))
        lab = re.search(r'\\label\{([^}]*)\}', inner).group(1)
        blocks.append(('fig', LABELS[lab], inline(cap.strip().rstrip('.')), g.group(2), float(g.group(1))))
    elif name.startswith('table'):
        ci = inner.index('\\caption'); cap, _ = read_group(inner, inner.index('{', ci))
        lab = re.search(r'\\label\{([^}]*)\}', inner).group(1)
        ts = inner.index('\\begin{tabular}')
        spec, k = read_group(inner, inner.index('{', ts + len('\\begin{tabular}')))
        tab = inner[k:inner.index('\\end{tabular}')]
        cols = re.findall(r'p\{([\d.]+)cm\}|([lcr])', spec)
        rows = []
        for r in re.split(r'\\\\', tab):
            r = r.replace('\\hline', '').strip()
            if r: rows.append([inline(c) for c in cell_split(r)])
        kind = 'Quadro' if lab in QUADROS else 'Tabela'
        blk = ['tab', kind, LABELS[lab], inline(cap.strip().rstrip('.')), cols, rows]
        TABLES[lab] = blk
        if FUNDIR and lab == 'tab:distancia':
            dst = TABLES['tab:custo']
            assert [r[0] for r in dst[5]] == [r[0] for r in rows] or True
            for rd, rs in zip(dst[5], rows):
                assert rd[1] == rs[1], 'cenários desalinhados ao fundir tabelas'
                rd.append(rs[2])
            dst[4] = dst[4] + [('', 'c')]
            continue
        blocks.append(blk)
    elif name == 'equation':
        tex = re.sub(r'\\label\{[^}]*\}', '', inner).strip()
        tex = re.sub(r'\\\\\[\d+pt\]', r'\\\\', tex)
        tex = tex.replace('{split}', '{aligned}').replace('\\left\\|', '\\|').replace('\\right\\|', '\\|')
        blocks.append(('eq', add_math(tex, True)))
    else:  # itemize / enumerate
        items = [x for x in re.split(r'\\item\s', inner)[1:]]
        for n, it in enumerate(items, 1):
            blocks.append(('li', '•' if name == 'itemize' else f'{n}.', inline(' '.join(it.split()))))

# ---------------------------------------------------------------- matemática via pandoc
def convert_math():
    tmp = tempfile.mkdtemp()
    src = []
    for n, (tex, disp) in enumerate(MATH):
        src.append(f'MK{n}KM ' + (f'\\begin{{equation*}}{tex}\\end{{equation*}}' if disp else f'\\({tex}\\)'))
    src.append(f'MK{len(MATH)}KM')
    open(os.path.join(tmp, 'm.tex'), 'w', encoding='utf8').write('\n\n'.join(src))
    subprocess.run([PANDOC, '-f', 'latex', '-t', 'docx', '-o', os.path.join(tmp, 'm.docx'),
                    os.path.join(tmp, 'm.tex')], check=True)
    x = zipfile.ZipFile(os.path.join(tmp, 'm.docx')).read('word/document.xml').decode('utf8')
    res = []
    for n in range(len(MATH)):
        seg = x[x.index(f'MK{n}KM'):x.index(f'MK{n+1}KM')]
        mm = re.findall(r'<m:oMath>.*?</m:oMath>', seg, re.S)
        if len(mm) != 1:
            raise SystemExit(f'falha na conversão da equação {n}: {MATH[n][0]!r}')
        o = re.sub(r'(<m:mcJc [^>]*/>)(<m:count [^>]*/>)', r'\2\1', mm[0])
        res.append(re.sub(r'(<m:nor ?/>)<m:sty [^>]*/>', r'\1', o))
    shutil.rmtree(tmp)
    return res
OMML = convert_math()

# ---------------------------------------------------------------- citações
def year_of(e): return e.get('year', 's.d.')
for k in CITED:
    if k not in BIB: raise SystemExit('citação sem entrada no .bib: ' + k)
groups = {}
for k in CITED:
    lab = ''.join(t for t, _ in cite_name(BIB[k])) + '|' + year_of(BIB[k])
    groups.setdefault(lab, []).append(k)
SUFFIX = {}
for lab, ks in groups.items():
    if len(ks) > 1:
        ks.sort(key=lambda k: unlatex(BIB[k]['title']).lower())
        for n, k in enumerate(ks): SUFFIX[k] = 'abcdefgh'[n]

def cite_runs(keys, narrative, b, i):
    out = []
    if narrative:
        for k in keys:
            e = BIB[k.lower()]
            out += [('t', t, b, i or it) for t, it in cite_name(e)]
            out.append(('t', f' ({year_of(e)}{SUFFIX.get(k.lower(), "")})', b, i))
        return out
    out.append(('t', '(', b, i))
    for n, k in enumerate(keys):
        e = BIB[k.lower()]
        if n: out.append(('t', '; ', b, i))
        out += [('t', t, b, i or it) for t, it in cite_name(e)]
        out.append(('t', f', {year_of(e)}{SUFFIX.get(k.lower(), "")}', b, i))
    out.append(('t', ')', b, i))
    return out

# ---------------------------------------------------------------- referências ABNT (NBR 6023)
MESES = ['jan.', 'fev.', 'mar.', 'abr.', 'maio', 'jun.', 'jul.', 'ago.', 'set.', 'out.', 'nov.', 'dez.']
def acesso(e):
    m = re.search(r'Accessed:\s*(\d{4})-(\d{2})-(\d{2})', e.get('note', ''))
    if m:
        return [('t', f' Acesso em: {int(m.group(3))} {MESES[int(m.group(2))-1]} {m.group(1)}.', 0, 0, None)]
    return [('t', ' Acesso em: ', 0, 0, None), ('t', '[informar data de acesso]', 0, 0, 'yellow'), ('t', '.', 0, 0, None)]

def ref_authors(e):
    au = e['_authors']
    if 'org' in au[0]: return [('t', au[0]['org'].upper(), 0, 0, None)]
    s = '; '.join(f"{x['last'].upper()}, {x['ini']}" for x in au)
    out = [('t', s, 0, 0, None)]
    if e['_others']:
        out.append(('t', '; [completar com todos os autores]', 0, 0, 'yellow'))
    return out

def sep_after_authors(e):
    au = e['_authors']
    return ' ' if ('org' not in au[0] and not e['_others'] and au[-1]['ini'].endswith('.')) else '. '

def bold_title(t):
    if ':' in t:
        a, b = t.split(':', 1)
        return [('t', a, 1, 0, None), ('t', ':' + b, 0, 0, None)]
    return [('t', t, 1, 0, None)]

def doi_url(e):
    out = []
    if e.get('doi'):
        out.append(('t', ' DOI: ', 0, 0, None)); out.append(('u', 'https://doi.org/' + e['doi']))
        out.append(('t', '.', 0, 0, None))
    elif e.get('url'):
        out.append(('t', ' Disponível em: ', 0, 0, None)); out.append(('u', e['url']))
        out.append(('t', '.', 0, 0, None)); out += acesso(e)
    return out

def reference(k):
    e = BIB[k]; t = e['_type']; y = year_of(e) + SUFFIX.get(k, '')
    title = unlatex(e['title'])
    R = ref_authors(e) + [('t', sep_after_authors(e), 0, 0, None)]
    pages = unlatex(e.get('pages', ''))
    if t == 'article' or e.get('journal'):
        if e.get('journal'):
            R += [('t', title + '. ', 0, 0, None), ('t', unlatex(e['journal']), 1, 0, None)]
            parts = []
            if e.get('volume'): parts.append('v. ' + e['volume'])
            if e.get('number'): parts.append('n. ' + e['number'])
            if pages: parts.append(('p. ' if re.match(r'\d', pages) else '') + pages)
            if e.get('eprint'): parts.append('arXiv:' + e['eprint'])
            R.append(('t', ', ' + ', '.join(parts + [y]) + '.', 0, 0, None))
            R += doi_url(e)
        else:
            R += bold_title(title) + [('t', f'. [S. l.], {y}.', 0, 0, None)] + doi_url(e)
    elif t == 'inproceedings':
        R += [('t', title + '. ', 0, 0, None), ('t', 'In', 0, 1, None),
              ('t', f": {unlatex(e['booktitle']).upper()}, {e['year']}. ", 0, 0, None),
              ('t', 'Proceedings', 1, 0, None), ('t', ' [...]. ', 0, 0, None)]
        pub = unlatex(e.get('publisher', e.get('organization', '')))
        R.append(('t', f"[S. l.]: {pub}, {y}." if pub else f"[S. l.: s. n.], {y}.", 0, 0, None))
        if pages: R.append(('t', f' p. {pages}.', 0, 0, None))
        R += doi_url(e)
    elif t == 'book':
        R += bold_title(title) + [('t', f". {unlatex(e.get('address', '[S. l.]'))}: {unlatex(e['publisher'])}, {y}.", 0, 0, None)]
        R += doi_url(e)
    else:  # online
        org = e['_authors'][0].get('org')
        R += bold_title(title) + [('t', f". [S. l.]: {org}, {y}." if org else f". [S. l.], {y}.", 0, 0, None)]
        R += [('t', ' Disponível em: ', 0, 0, None), ('u', e['url']), ('t', '.', 0, 0, None)] + acesso(e)
    sortkey = ''.join(r[1] for r in R if r[0] == 't').upper()
    return sortkey, R

# ---------------------------------------------------------------- XML
TNR = '<w:rFonts w:ascii="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman"/>'
def rpr(b=False, i=False, sz=None, hl=None, color=None, sup=False, u=False):
    s = TNR
    if b: s += '<w:b w:val="1"/><w:bCs w:val="1"/>'
    if i: s += '<w:i w:val="1"/><w:iCs w:val="1"/>'
    if color: s += f'<w:color w:val="{color}"/>'
    if sz: s += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>'
    if hl: s += f'<w:highlight w:val="{hl}"/>'
    if u: s += '<w:u w:val="single"/>'
    if sup: s += '<w:vertAlign w:val="superscript"/>'
    return f'<w:rPr>{s}</w:rPr>'

def run(t, **kw):
    return f'<w:r>{rpr(**kw)}<w:t xml:space="preserve">{escape(t)}</w:t></w:r>'

RELS = []
def rel(typ, target, external=False):
    rid = f'rId{200 + len(RELS)}'
    RELS.append((rid, typ, target, external)); return rid

def runs_xml(runs, sz=None, force_i=False):
    out = ''
    for r in runs:
        if r[0] == 't':
            out += run(r[1], b=r[2], i=r[3] or force_i, sz=sz)
        elif r[0] == 'm':
            out += OMML[r[1]]
        elif r[0] == 'c':
            out += runs_xml(cite_runs(r[1], r[2], r[3], r[4]), sz, force_i)
    return out

def para(content, jc='both', first=False, ind=None, sz=None, after=None, keep=False, tabs=None):
    ppr = '<w:pPr>'
    if keep: ppr += '<w:keepNext/>'
    if tabs: ppr += tabs
    ppr += '<w:spacing ' + (f'w:after="{after}" ' if after else '') + 'w:line="240" w:lineRule="auto"/>'
    if first: ppr += '<w:ind w:firstLine="567"/>'
    if ind: ppr += ind
    szx = f'<w:sz w:val="{sz}"/>' if sz else ''
    ppr += f'<w:jc w:val="{jc}"/><w:rPr>{TNR}{szx}</w:rPr></w:pPr>'
    return f'<w:p>{ppr}{content}</w:p>'

BLANK = para('')
TEXTW = 9073   # largura útil em twips (A4 com margens de 2,5 cm)

def png_size(path):
    with open(path, 'rb') as f:
        h = f.read(24)
    return struct.unpack('>II', h[16:24])

IMGS = []
def figure_xml(num, cap, path, frac):
    full = os.path.join(PROJ, path)
    w, h = png_size(full)
    cx = int(frac * TEXTW * 635)                  # twip -> EMU
    cy = int(cx * h / w)
    maxh = int(FIG_MAX * 360000)
    if cy > maxh: cx, cy = int(cx * maxh / cy), maxh
    n = len(IMGS) + 1
    IMGS.append((f'media/figura{n}.png', full))
    rid = rel('image', f'media/figura{n}.png')
    drawing = (f'<w:r>{rpr()}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
               f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{n}" name="Figura {n}"/>'
               '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
               '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
               f'<pic:nvPicPr><pic:cNvPr id="{n}" name="figura{n}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
               f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
               f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
               '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')
    return (BLANK + para(run(f'Figura {num} – ', sz=20) + runs_xml(cap, sz=20), jc='center', keep=True)
            + para(drawing, jc='center', after=120, keep=True)
            + para(run('Fonte: elaborado pelos autores', sz=16), jc='both') + BLANK)

def border(side, sz, val='single'):
    return f'<w:{side} w:val="{val}" w:sz="{sz}" w:space="0" w:color="000000"/>'

def table_xml(kind, num, cap, cols, rows):
    nat = []
    for p, a in cols:
        nat.append(float(p) * 567 if p else (2000 if a == 'l' else 1300))
    tot = sum(nat)
    widths = [int(x * min(1, 9000 / tot)) for x in nat]
    tw = sum(widths)
    aligns = ['center' if a == 'c' else 'left' for p, a in cols]
    if kind == 'Quadro':
        tb = ''.join(border(s, 8) for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    else:   # tabela IBGE: sem bordas laterais, sem linhas horizontais internas
        tb = border('top', 12) + border('left', 0, 'nil') + border('bottom', 12) + border('right', 0, 'nil') + border('insideH', 0, 'nil') + border('insideV', 4)
    x = (f'<w:tbl><w:tblPr><w:tblW w:w="{tw}" w:type="dxa"/><w:jc w:val="center"/><w:tblBorders>{tb}</w:tblBorders>'
         '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="80" w:type="dxa"/><w:right w:w="80" w:type="dxa"/></w:tblCellMar>'
         '<w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>' + ''.join(f'<w:gridCol w:w="{w}"/>' for w in widths) + '</w:tblGrid>')
    for ri, row in enumerate(rows):
        x += '<w:tr><w:trPr><w:cantSplit/>' + ('<w:tblHeader/>' if ri == 0 else '') + '</w:trPr>'
        for ci, cell in enumerate(row):
            tcb = ''
            if kind == 'Tabela' and ri == 0: tcb = f'<w:tcBorders>{border("bottom", 4)}</w:tcBorders>'
            jc = 'center' if ri == 0 else aligns[ci]
            x += (f'<w:tc><w:tcPr><w:tcW w:w="{widths[ci]}" w:type="dxa"/>{tcb}<w:vAlign w:val="center"/></w:tcPr>'
                  + para(runs_xml(cell, sz=20), jc=jc, sz=20) + '</w:tc>')
        x += '</w:tr>'
    x += '</w:tbl>'
    fonte = 'Fonte: elaborado pelos autores' if kind == 'Quadro' else 'Fonte: dados da pesquisa'
    return (BLANK + para(run(f'{kind} {num} – ', sz=20) + runs_xml(cap, sz=20), jc='center', keep=True)
            + x + para(run(fonte, sz=16), jc='both') + BLANK)

EQN = [0]
def eq_xml(idx):
    EQN[0] += 1
    nb = ''.join(border(s, 0, 'nil') for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'))
    return (f'<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>{nb}</w:tblBorders><w:tblLayout w:type="fixed"/>'
            '<w:tblLook w:val="0600"/></w:tblPr><w:tblGrid><w:gridCol w:w="8300"/><w:gridCol w:w="700"/></w:tblGrid><w:tr>'
            f'<w:tc><w:tcPr><w:tcW w:w="8300" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>{para(OMML[idx], jc="center")}</w:tc>'
            f'<w:tc><w:tcPr><w:tcW w:w="700" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>{para(run(f"({EQN[0]})"), jc="right")}</w:tc>'
            '</w:tr></w:tbl>')

# ---------------------------------------------------------------- montagem do corpo
X = []
title_pt = 'Simulação de arquitetura de controle para coordenação de múltiplos drones'
assert TITLE_LATEX.lower() == title_pt.lower(), TITLE_LATEX
title_en = 'Simulation of a control architecture for multi-drone coordination'

X.append(para(run(title_pt, b=True, sz=24), jc='center'))
X.append(BLANK)
if WITH_AUTHORS:
    AUT = [('Jadsonlee da Silva Sá', 'jadsonlee.sa@univasf.edu.br'),
           ('Rosalvo Ferreira de Oliveira Neto', 'rosalvo.oliveira@univasf.edu.br'),
           ('Daniel Alencar Penha Carvalho', 'daniel.apcarvalho@discente.univasf.edu.br')]
    c = ''
    for n, (nome, _) in enumerate(AUT, 1):
        if n > 1: c += run(', ')
        c += run(nome) + run(f'[{n}]', sup=True) + (run('*') if n == 1 else '')
    X.append(para(c, jc='left'))
    X.append(BLANK)
    for n, (_, mail) in enumerate(AUT, 1):
        X.append(para(run(f'[{n}] ', sup=True) + run(f'{mail}. Universidade Federal do Vale do São Francisco (UNIVASF), Juazeiro, Bahia, Brasil. ORCID: ')
                      + run('https://orcid.org/0009-0009-2637-5998', hl='yellow'), jc='left'))
    X.append(para(run('* autor correspondente'), jc='left'))
    X.append(BLANK)

def kw(s):
    items = [x.strip() for x in s.split(',')]
    items = [x if x.isupper() or re.match(r'[A-Z]{2,}', x) else x[0].lower() + x[1:] for x in items]
    return '; '.join(sorted(items, key=str.lower)) + '.'

X.append(para(run('Resumo', b=True)))
X.append(para(runs_xml(inline(ABS_PT))))
X.append(para(run('Palavras-chave:', b=True) + run(' ' + kw(KW_PT))))
X.append(BLANK)
X.append(para(run(title_en, b=True, i=True, sz=24), jc='center'))
X.append(BLANK)
X.append(para(run('Abstract', b=True, i=True)))
X.append(para(runs_xml(inline(ABS_EN), force_i=True)))
X.append(para(run('Keywords:', b=True, i=True) + run(' ' + kw(KW_EN), i=True)))

for blk in blocks:
    t = blk[0]
    if t == 'h':
        X.append(BLANK); X.append(para(run(blk[1], b=True), keep=True))
    elif t == 'p':
        X.append(para(runs_xml(blk[1]), first=not blk[2]))
    elif t == 'li':
        X.append(para(run(blk[1]) + run('\t') + runs_xml(blk[2]),
                      tabs='<w:tabs><w:tab w:val="left" w:pos="851"/></w:tabs>', ind='<w:ind w:left="851" w:hanging="284"/>'))
    elif t == 'eq':
        X.append(eq_xml(blk[1]))
    elif t == 'fig':
        X.append(figure_xml(blk[1], blk[2], blk[3], blk[4]))
    elif t == 'tab':
        X.append(table_xml(*blk[1:]))

def placeholder_section(title, text_runs):
    X.append(BLANK); X.append(para(run(title, b=True), keep=True)); X.append(para(text_runs))

placeholder_section('Financiamento', run('[Informar as fontes e os números de processo do financiamento ou declarar: “Esta pesquisa não recebeu financiamento”.]', hl='yellow'))
placeholder_section('Conflito de interesses', run('Os autores declaram não haver conflito de interesses.', hl='yellow'))
if WITH_AUTHORS:
    placeholder_section('Nota', run('Este artigo é derivado do Trabalho de Conclusão de Curso de Daniel Alencar Penha Carvalho, da Universidade Federal do Vale do São Francisco (UNIVASF), disponível em: ') + run('[endereço eletrônico do trabalho no repositório institucional]', hl='yellow') + run('.'))
    placeholder_section('Contribuições ao artigo', run('[Preencher conforme a Declaração sobre as diretrizes da Revista Principia – Documento Suplementar.]', hl='yellow'))

X.append(BLANK); X.append(para(run('Referências', b=True), keep=True))
refs = sorted(reference(k) for k in CITED)
for _, R in refs:
    c = ''
    for r in R:
        if r[0] == 'u':
            rid = rel('hyperlink', r[1], True)
            c += f'<w:hyperlink r:id="{rid}">{run(r[1], color="0000ff", u=True)}</w:hyperlink>'
        else:
            c += run(r[1], b=r[2], i=r[3], hl=r[4])
    X.append(BLANK); X.append(para(c, jc='left'))

# ---------------------------------------------------------------- pacote
tpl_doc = open(os.path.join(TPL, 'word/document.xml'), encoding='utf8').read()
head = tpl_doc[:tpl_doc.index('<w:body>') + len('<w:body>')]
sect = tpl_doc[tpl_doc.rindex('<w:sectPr>'):].replace('w:footer="720"/>', 'w:footer="720" w:gutter="0"/>')
doc = head + ''.join(X) + sect

build = tempfile.mkdtemp()
shutil.copytree(TPL, build, dirs_exist_ok=True)
for f in os.listdir(os.path.join(build, 'word/media')): os.remove(os.path.join(build, 'word/media', f))
for dst, src in IMGS: shutil.copy(src, os.path.join(build, 'word', dst))
open(os.path.join(build, 'word/document.xml'), 'w', encoding='utf8').write(doc)
ftr = os.path.join(build, 'word/footer1.xml')
f = open(ftr, encoding='utf8').read()
f = f.replace('<w:gridCol w:w="8755"/><w:gridCol w:w="458"/>', '<w:gridCol w:w="8513"/><w:gridCol w:w="700"/>')
open(ftr, 'w', encoding='utf8').write(f)

base = re.findall(r'<Relationship [^>]*/>', open(os.path.join(TPL, 'word/_rels/document.xml.rels'), encoding='utf8').read())
keep = [r for r in base if re.search(r'Type="[^"]*/(theme|settings|fontTable|numbering|styles|footer)"', r)]
T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
new = [f'<Relationship Id="{rid}" Type="{T}{typ}" Target="{escape(tg, {chr(34): "&quot;"})}"' + (' TargetMode="External"' if ext else '') + '/>'
       for rid, typ, tg, ext in RELS]
open(os.path.join(build, 'word/_rels/document.xml.rels'), 'w', encoding='utf8').write(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + ''.join(keep + new) + '</Relationships>')

if os.path.exists(OUT): os.remove(OUT)
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as z:
    z.write(os.path.join(build, '[Content_Types].xml'), '[Content_Types].xml')
    for root, _, files in os.walk(build):
        for f in files:
            p = os.path.join(root, f); a = os.path.relpath(p, build)
            if a != '[Content_Types].xml': z.write(p, a)
shutil.rmtree(build)
shutil.rmtree(TPL)
print('ok', OUT, '| equações:', EQN[0], '| figuras:', len(IMGS), '| refs:', len(refs))
