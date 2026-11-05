"""Delivery checks plus a reference model; this does NOT execute VBA or Word."""
import json
import random
import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'WordLaTexConverter.bas'
code = SOURCE.read_text(encoding='ascii')

def escaped(s, p):
    n = 0
    p -= 1
    while p >= 0 and s[p] == '\\':
        n += 1
        p -= 1
    return n % 2 == 1

def dollar_run(s, p):
    q = p
    while q < len(s) and s[q] == '$':
        q += 1
    return q - p

def token(s, p):
    if escaped(s, p):
        return ''
    if s[p:p+2] in (r'\(', r'\[', r'\)', r'\]'):
        return s[p:p+2]
    if s[p:p+1] == '$' and (p == 0 or s[p-1] != '$'):
        n = dollar_run(s, p)
        if n in (1, 2):
            return '$' * n
    return ''

def scan(s):
    p = 0
    spans = []
    spaces = ' \t\r\n\v'
    while p < len(s):
        op = token(s, p)
        cl = {r'\(': r'\)', r'\[': r'\]', '$': '$', '$$': '$$'}.get(op, '')
        if not cl:
            p += 1
            continue
        width = len(op)
        display = op in (r'\[', '$$')
        if op == '$' and s[p+1:p+2] and s[p+1:p+2] in spaces:
            p += 1
            continue
        q = p + width
        found = False
        while q < len(s):
            ch = s[q]
            if ch in '\0\x07\x0c':
                break
            if not display and ch in '\r\n\v':
                break
            if display and s[q:q+2] == '\r\r':
                break
            if q - p - width > 16000:
                break
            if ch == '$' and not escaped(s, q) and dollar_run(s, q) > 2:
                break
            other = token(s, q)
            if other == cl:
                if cl != '$' or s[q-1] not in spaces:
                    found = True
                    break
            elif other:
                break
            q += 1
        if found:
            spans.append((p, q + width, display))
            p = q + width
        else:
            p = q if q > p + width else p + width
    return spans

def vba_literal(expr):
    terms = re.findall(r'"(?:[^"]|"")*"|ChrW\(92\)|\bvbCr\b|\bBARRIER\b', expr)
    return ''.join(chr(92) if t == 'ChrW(92)' else '\r' if t == 'vbCr' else '\0' if t == 'BARRIER'
                   else t[1:-1].replace('""', '"') for t in terms)

cases = []
for name, expr, expected in re.findall(
    r'^\s*CheckScan "([^"]+)", (.*), (\d+), report, passed, failed$', code, re.M
):
    cases.append((name, vba_literal(expr), int(expected)))
assert len(cases) == 16
for name, value, expected in cases:
    actual = len(scan(value))
    assert actual == expected, (name, actual, expected)

# Independent golden offsets and delimiter widths (zero-based, half-open).
golden = [
    ('aa ' + r'\(x\)' + ' bb', [(3, 8, False)]),
    ('$x$$$y$', []),
    (r'\$100 ' + r'\(y\)', [(6, 11, False)]),
    ('$$x$$ $y$', [(0, 5, True), (6, 9, False)]),
    ('\\[' + 'x' * 16001 + '\\]', []),
    ('\\[' + 'x' * 16000 + '\\]', [(0, 16004, True)]),
    ('\\[bad\r\r\\(x\\)', [(7, 12, False)]),
]
for value, expected in golden:
    assert scan(value) == expected, (value[:40], scan(value), expected)

rng = random.Random(20261105)
for _ in range(1000):
    value = ''.join(rng.choice(['x', '$', '$$', '$$$', '\\', '(', ')', '[', ']', '\r', '\0', ' '])
                    for _ in range(rng.randrange(50)))
    last = 0
    for a, b, display in scan(value):
        assert last <= a < b <= len(value)
        assert '\0' not in value[a:b]
        if not display:
            assert not any(c in value[a:b] for c in '\r\n\v')
        last = b

# Structural source checks, explicitly not a VBA compiler.
starts = re.findall(r'^(?:Public|Private) (Sub|Function) (\w+)', code, re.M)
ends = re.findall(r'^End (Sub|Function)', code, re.M)
assert [x[0] for x in starts] == ends
assert len({x[1].lower() for x in starts}) == len(starts)
assert max(map(len, code.splitlines())) < 1024
assert code.index('eq.BuildUp') < code.index('target.FormattedText = eq.Range.FormattedText')
assert 'doc.Undo' not in code and 'CreateObject' not in code
assert 'If target Is Nothing Then' in code
assert 'changed.FormattedText = saved.FormattedText' in code
assert 'ROLLBACK FAILED; STOPPED' in code
result = {
    'vba_execution': 'NOT RUN by this checker; see mac-word-results.json',
    'word_macro_compilation': 'NOT RUN by this checker; see mac-word-results.json',
    'reference_model_cases_from_vba_test_vectors': len(cases),
    'independent_golden_cases': len(golden),
    'random_reference_model_invariant_cases': 1000,
    'structural_source_checks': 'PASS (not syntax/type compilation)',
    'reference_model_checks': 'PASS',
}
print(json.dumps(result, indent=2))
