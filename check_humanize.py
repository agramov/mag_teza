# -*- coding: utf-8 -*-
"""Сравнява оригинален и хуманизиран текст и отчита само това, което заплашва
достоверността на цитираното: цитати, числа, модалност, обхватни думи.

Употреба:
    python3 check_humanize.py original.docx humanized.docx [--section 1.5]
Приема .docx или .txt/.md.
"""
import sys, re, unicodedata

def load(path):
    if path.lower().endswith('.docx'):
        import docx
        return '\n'.join(p.text for p in docx.Document(path).paragraphs)
    return open(path, encoding='utf-8').read()

def section(txt, sec):
    if not sec: return txt
    m = re.search(rf'^\s*{re.escape(sec)}\..*$', txt, re.M)
    if not m: sys.exit(f'Разделът {sec} не е намерен.')
    nxt = re.search(r'^\s*\d\.\d\.\s', txt[m.end():], re.M)
    return txt[m.start(): m.end()+nxt.start()] if nxt else txt[m.start():]

CITE   = re.compile(r'\[(\d{1,2})(?:[^\]\[]{0,40})?\]')
NUMTOK = re.compile(r'\d[\d\s  ,.]*\d|\d')
WORDNUM = ('една шеста|една трета|половин|двойно|тридесет и три|четиридесет и четири|'
           'осем|девет|пет|три|четири|шест|седем|десет|сто|над|около|приблизително|едва')
HEDGE = ('може|могат|възможно|вероятно|по правило|обикновено|често|склонн|тенденция|'
         'отчасти|донякъде|изглежда|при определени условия|в изследвани условия|'
         'най-общо|до голяма степен|преобладава|частично|в редица случаи|нерядко')
HARD  = ('доказва|доказано|всички|винаги|никога|безспорно|категорично|ясно показва|'
         'установява|потвърждава|неизменно|напълно|изцяло|единствено|само')

def counts(t, pat):
    return len(re.findall(pat, t, re.I))

def sentences(t):
    return [s.strip() for s in re.split(r'(?<=[.!?])\s+(?=[А-ЯA-Z„])', t) if s.strip()]

def norm_num(s):
    return re.sub(r'[\s ]', '', s)

def report(a, b):
    ca = [m.group(1) for m in CITE.finditer(a)]
    cb = [m.group(1) for m in CITE.finditer(b)]
    print('=== ЦИТАТИ')
    print(f'  оригинал: {len(ca)} маркера, източници {sorted(set(map(int,ca)))}')
    print(f'  нов:      {len(cb)} маркера, източници {sorted(set(map(int,cb)))}')
    if ca != cb:
        print('  !! РАЗЛИКА в поредността/броя на цитатите — проверявай ръчно')
        lost = set(ca) - set(cb); new = set(cb) - set(ca)
        if lost: print('  !! изчезнали източници:', sorted(map(int, lost)))
        if new:  print('  !! появили се източници:', sorted(map(int, new)))
    else:
        print('  ОК: същите маркери в същия ред')

    na = sorted(norm_num(x) for x in NUMTOK.findall(re.sub(CITE, ' ', a)))
    nb = sorted(norm_num(x) for x in NUMTOK.findall(re.sub(CITE, ' ', b)))
    print('=== ЧИСЛА (без номерата на цитатите)')
    if na != nb:
        print('  !! РАЗЛИКА'); print('   само в оригинала:', [x for x in na if x not in nb])
        print('   само в новия:  ', [x for x in nb if x not in na])
    else:
        print(f'  ОК: {len(na)} числа, идентични')
    wa, wb_ = counts(a, WORDNUM), counts(b, WORDNUM)
    print(f'=== ЧИСЛА С ДУМИ / ОБХВАТНИ: {wa} -> {wb_}' + ('   !! ПРОВЕРИ' if wa != wb_ else '   ОК'))

    ha, hb = counts(a, HEDGE), counts(b, HEDGE)
    sa, sb = counts(a, HARD), counts(b, HARD)
    print('=== МОДАЛНОСТ (най-важното)')
    print(f'  смекчаващи: {ha} -> {hb}   ({hb-ha:+d})')
    print(f'  усилващи:   {sa} -> {sb}   ({sb-sa:+d})')
    if hb < ha or sb > sa:
        print('  !! ТЕКСТЪТ Е СТАНАЛ ПО-КАТЕГОРИЧЕН. Всяко такова изречение с цитат')
        print('     трябва да се провери наново срещу източника.')
    else:
        print('  ОК: категоричността не е нараснала')

    la, lb = len(a.split()), len(b.split())
    print(f'=== ОБЕМ: {la} -> {lb} думи ({(lb-la)/la*100:+.1f}%)')
    if lb < la*0.75: print('  !! свиване над 25 % — вероятно е изпуснато съдържание')

    print('=== ИЗРЕЧЕНИЯ С ЦИТАТ, КОИТО СА ПРОМЕНЕНИ')
    sb_set = set(sentences(b))
    n = 0
    for s in sentences(a):
        if CITE.search(s) and s not in sb_set:
            n += 1
            print(f'  [{n}] {s[:150]}...')
    if n == 0: print('  няма — всички цитирани изречения са непокътнати')
    else: print(f'  общо {n} цитирани изречения са променени -> подлежат на повторна проверка')

if __name__ == '__main__':
    if len(sys.argv) < 3: sys.exit(__doc__)
    sec = None
    if '--section' in sys.argv:
        sec = sys.argv[sys.argv.index('--section')+1]
    A = section(load(sys.argv[1]), sec)
    B = section(load(sys.argv[2]), sec)
    report(A, B)
