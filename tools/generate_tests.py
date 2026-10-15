"""Generate the VBA test modules in tests/ from Python's textdistance.

Expected values come from textdistance's own pure-Python implementations
(external=False), so the VBA port is checked against the reference code.

Usage:  pip install textdistance
        python3 tools/generate_tests.py
"""
import os
import random

import textdistance as td

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

COMMON_CASES = [
    ("", ""), ("", "abc"), ("abc", ""), ("a", "a"), ("a", "b"), ("same", "same"),
    ("test", "text"), ("test", "testing"), ("testing", "test"),
    ("kitten", "sitting"), ("sitting", "kitten"), ("flaw", "lawn"),
    ("karolin", "kathrin"), ("1011101", "1001001"), ("abcd", "abxyz"),
    ("ca", "abc"), ("abcd", "acbd"), ("abcdef", "badcfe"),
    ("martha", "marhta"), ("dwayne", "duane"), ("dixon", "dicksonx"),
    ("jellyfish", "smellyfish"), ("Hello", "hello"), ("ABC", "xyz"),
    ("crate", "trace"), ("abcdefghijklmnop", "abcdefghijklmnoq"),
    ("Jonathan Smith", "Jonathon Smyth"), ("123 Main Street", "123 Main St."),
    ('say "hi"', "say 'hi'"), ("café", "cafe"), ("naïve", "naive"),
]
CASE_INSENSITIVE_CASES = [("Hello", "hello"), ("TeSt", "tEXT"), ("ABC", "abcd"),
                          ("MARTHA", "marhta"), ("Kitten", "SITTING")]


def long_cases(seed):
    """Strings of 200+ characters, where textdistance.LCSStr switches algorithm."""
    rng = random.Random(seed)
    word = lambda n, alphabet: "".join(rng.choice(alphabet) for _ in range(n))
    return [(word(n1, al), word(n2, al)) for n1, n2, al in (
        (210, 230, "ab"), (230, 210, "ab"), (205, 199, "abc"), (150, 190, "abcd"),
        (240, 220, "abcd"), (220, 240, "abcdefgh"), (260, 201, "ab"), (300, 250, "abc"),
        (250, 230, "abcdef"), (330, 290, "abcd"))]


def random_cases(seed, count=25):
    rng = random.Random(seed)
    word = lambda: "".join(rng.choice("abcd") for _ in range(rng.randint(1, 9)))
    return [(word(), word()) for _ in range(count)]


# name: (VBA prefix, [(VBA extra args, textdistance instance)], integer results?)
ALGORITHMS = {
    "Hamming": ("Hamming", [
        ("False", td.Hamming(truncate=False, external=False)),
        ("True", td.Hamming(truncate=True, external=False)),
    ], True),
    "Levenshtein": ("Levenshtein", [
        ("", td.Levenshtein(external=False)),
    ], True),
    "DamerauLevenshtein": ("DamerauLevenshtein", [
        ("True", td.DamerauLevenshtein(restricted=True, external=False)),
        ("False", td.DamerauLevenshtein(restricted=False, external=False)),
    ], True),
    "Jaro": ("Jaro", [
        ("False", td.Jaro(long_tolerance=False, external=False)),
        ("True", td.Jaro(long_tolerance=True, external=False)),
    ], False),
    "JaroWinkler": ("JaroWinkler", [
        ("False", td.JaroWinkler(long_tolerance=False, external=False)),
        ("True", td.JaroWinkler(long_tolerance=True, external=False)),
    ], False),
    "LCSSeq": ("LCSSeq", [("", td.LCSSeq(external=False))], True),
    "LCSStr": ("LCSStr", [("", td.LCSStr(external=False))], True),
    "RatcliffObershelp": ("RatcliffObershelp", [("", td.RatcliffObershelp(external=False))], False),
    "Prefix": ("Prefix", [("", td.Prefix())], True),
    "Postfix": ("Postfix", [("", td.Postfix())], True),
    "Length": ("Length", [("", td.Length())], True),
    "Identity": ("Identity", [("", td.Identity())], True),
    "MRA": ("MRA", [("", td.MRA())], True),
    "Editex": ("Editex", [
        ("False", td.Editex(local=False, external=False)),
        ("True", td.Editex(local=True, external=False)),
        ("False, 1, 2, 4", td.Editex(match_cost=1, group_cost=2, mismatch_cost=4, external=False)),
        ("False, 3, 1, 2", td.Editex(match_cost=3, group_cost=1, mismatch_cost=2, external=False)),
    ], True),
}

# Algorithms that always ignore case (no CaseSensitive argument)
NO_CASE_ARG = {"MRA", "Editex"}
# Extra cases for particular algorithms
NAME_CASES = [("Byrne", "Boern"), ("Smith", "Smyth"), ("Catherine", "Kathryn"),
              ("Robert", "Rupert"), ("Ashcraft", "Ashcroft"), ("Thompson", "Thomson"),
              ("Aaron", "Erin"), ("Lee", "Leigh"), ("Jon", "John"), ("Schmidt", "Smith"),
              ("Williams", "Wiliams"), ("Hawthorne", "Hathorn"), ("Shackleford", "Shackelford"),
              ("Bartholomew", "Bart"), ("O'Brien", "OBrien"), ("Mc Donald", "MacDonald")]
EXTRA_CASES = {"MRA": NAME_CASES, "Editex": NAME_CASES}

# Matrix takes a lookup table instead of the usual arguments; see matrix_checks
DNA_TABLE = [("A", "A", 10), ("G", "G", 7), ("C", "C", 9), ("T", "T", 8),
             ("A", "G", -1), ("A", "C", -3), ("A", "T", -4), ("G", "C", -5),
             ("G", "T", -3), ("C", "T", 0)]
DUP_TABLE = DNA_TABLE + [("A", "G", 99), ("a", "c", 42)]
HELPERS = {"Matrix": [("DnaTable", DNA_TABLE), ("DupTable", DUP_TABLE)]}

# Extra string-returning functions: name -> VBA suffix (Python: algo(a, b))
STRING_FUNCS = {"LCSSeq": "Sequence", "LCSStr": "Substring",
                "Prefix": "CommonPrefix", "Postfix": "CommonPostfix"}
# Algorithms also tested on 200+ character strings
LONG_CASE_ALGOS = {"LCSSeq", "LCSStr", "RatcliffObershelp"}

# Which test module each algorithm's tests go into
TEST_FILES = {
    "HammingTests": ["Hamming"],
    "LevenshteinTests": ["Levenshtein"],
    "DamerauLevenshteinTests": ["DamerauLevenshtein"],
    "JaroWinklerTests": ["Jaro", "JaroWinkler"],
    "LCSSeqTests": ["LCSSeq"],
    "LCSStrTests": ["LCSStr"],
    "RatcliffObershelpTests": ["RatcliffObershelp"],
    "PrefixTests": ["Prefix"],
    "PostfixTests": ["Postfix"],
    "LengthTests": ["Length"],
    "IdentityTests": ["Identity"],
    "MatrixTests": ["Matrix"],
    "MRATests": ["MRA"],
    "EditexTests": ["Editex"],
}


def vba_str(s):
    """VBA expression for s, using ChrW() for non-ASCII characters."""
    parts, buf = [], ""
    for ch in s:
        if ord(ch) < 128:
            buf += ch
        else:
            if buf:
                parts.append('"' + buf.replace('"', '""') + '"')
                buf = ""
            parts.append(f"ChrW({ord(ch)})")
    if buf or not parts:
        parts.append('"' + buf.replace('"', '""') + '"')
    return " & ".join(parts)


def vba_num(x):
    return repr(int(x)) if float(x).is_integer() else repr(float(x))


def short_repr(s, limit=24):
    return repr(s) if len(s) <= limit else repr(s[:limit]) + f"...({len(s)} chars)"


def checks_for(name, seed):
    prefix, variants, integer = ALGORITHMS[name]
    check = "CheckL" if integer else "CheckD"
    cases = COMMON_CASES + random_cases(seed)
    if name in LONG_CASE_ALGOS:
        cases += long_cases(seed)
    cases += EXTRA_CASES.get(name, [])
    cs_arg = "" if name in NO_CASE_ARG else ", True"
    lines = []
    for index, (extra, algo) in enumerate(variants):
        # Variants beyond the first two get a smaller case set to keep modules small
        for a, b in (cases if index < 2 else COMMON_CASES[:12] + EXTRA_CASES.get(name, [])):
            args = f"{vba_str(a)}, {vba_str(b)}{cs_arg}" + (f", {extra}" if extra else "")
            label = f"{prefix}({short_repr(a)}, {short_repr(b)}{', ' + extra if extra else ''})"
            for func, value, chk in (
                ("Distance", algo.distance(a, b), check),
                ("Similarity", algo.similarity(a, b), check),
                ("NormalizedDistance", algo.normalized_distance(a, b), "CheckD"),
                ("NormalizedSimilarity", algo.normalized_similarity(a, b), "CheckD"),
            ):
                lines.append(f'    {chk} {vba_str(label + " " + func)}, '
                             f"{prefix}_{func}({args}), {vba_num(value)}")
            if name in STRING_FUNCS:
                func = STRING_FUNCS[name]
                lines.append(f'    CheckS {vba_str(label + " " + func)}, '
                             f"{prefix}_{func}({args}), {vba_str(algo(a, b))}")
    if name == "MRA":
        names = sorted({x for pair in cases for x in pair})
        for word in names:
            lines.append(f"    CheckS {vba_str('MRA_Codex(' + repr(word) + ')')}, "
                         f"MRA_Codex({vba_str(word)}), {vba_str(td.MRA()._calc_mra(word))}")
    # CaseSensitive:=False is a VBA addition: compare with Python on lowercased input
    extra, algo = variants[0]
    for a, b in ([] if name in NO_CASE_ARG else CASE_INSENSITIVE_CASES):
        args = f"{vba_str(a)}, {vba_str(b)}, False" + (f", {extra}" if extra else "")
        label = f"{prefix}({a!r}, {b!r}, CaseSensitive=False)"
        lines.append(f"    {check} {vba_str(label + ' Distance')}, {prefix}_Distance({args}), "
                     f"{vba_num(algo.distance(a.lower(), b.lower()))}")
        if name in STRING_FUNCS:
            func = STRING_FUNCS[name]
            lines.append(f"    CheckS {vba_str(label + ' ' + func)}, {prefix}_{func}({args}), "
                         f"{vba_str(algo(a.lower(), b.lower()))}")
    lines.append(f'    CheckD "{prefix}_Maximum", {prefix}_Maximum("test", "testing"), '
                 f'{vba_num(variants[0][1].maximum("test", "testing"))}')
    return lines


def matrix_checks():
    pairs = [(a, b) for a in "ACGT" for b in "ACGT"] + [
        ("A", "X"), ("X", "A"), ("X", "X"), ("AG", "AG"), ("AG", "GA"), ("", ""), ("", "A")]
    # (VBA table expression, Python table, MismatchCost, MatchCost, Symmetric, CaseSensitive)
    variants = [
        ("Empty", None, 0, 1, True, True),
        ("Empty", None, -2, 5, True, True),
        ("DnaTable()", DNA_TABLE, 0, 1, True, True),
        ("DnaTable()", DNA_TABLE, 0, 1, False, True),
        ("DnaTable()", DNA_TABLE, -7, 0, True, True),
        ("DupTable()", DUP_TABLE, 0, 1, True, True),
        ("DupTable()", DUP_TABLE, 0, 1, True, False),
    ]
    lines = []
    for table_expr, table, mismatch, match, symmetric, cs in variants:
        mat = None
        if table:
            mat = {}
            for a, b, v in table:
                mat[(a, b) if cs else (a.lower(), b.lower())] = v
        algo = td.Matrix(mat=mat, mismatch_cost=mismatch, match_cost=match, symmetric=symmetric)
        extra = [] if cs else [(p.lower(), q) for p, q in pairs[:6]]
        for a, b in pairs + extra:
            pa, pb = (a, b) if cs else (a.lower(), b.lower())
            args = f"{vba_str(a)}, {vba_str(b)}, {table_expr}, {mismatch}, {match}, {symmetric}, {cs}"
            label = f"Matrix({a!r}, {b!r}, {table_expr}, {mismatch}, {match}, {symmetric}, {cs})"
            for func, value in (("Similarity", algo.similarity(pa, pb)),
                                ("Distance", algo.distance(pa, pb)),
                                ("NormalizedDistance", algo.normalized_distance(pa, pb)),
                                ("NormalizedSimilarity", algo.normalized_similarity(pa, pb)),
                                ("Maximum", algo.maximum(pa, pb))):
                lines.append(f"    CheckD {vba_str(label + ' ' + func)}, "
                             f"Matrix_{func}({args}), {vba_num(value)}")
    # Omitted optional arguments
    lines.append('    CheckD "Matrix(""A"", ""A"") defaults", Matrix_Similarity("A", "A"), 1')
    lines.append('    CheckD "Matrix(""A"", ""G"") defaults", Matrix_Similarity("A", "G"), 0')
    lines.append('    CheckD "Matrix(""A"", ""G"", DnaTable()) defaults", '
                 'Matrix_Similarity("A", "G", DnaTable()), -1')
    return lines


def helper_lines(names):
    """Private functions building the 2-D tables the tests pass to Matrix."""
    lines = []
    for name in names:
        for func, table in HELPERS.get(name, []):
            lines += ["", f"Private Function {func}() As Variant", "    Dim t() As Variant",
                      f"    ReDim t(1 To {len(table)}, 1 To 3)"]
            for r, (a, b, v) in enumerate(table, 1):
                lines.append(f'    t({r}, 1) = "{a}": t({r}, 2) = "{b}": t({r}, 3) = {v}')
            lines += [f"    {func} = t", "End Function"]
    return lines


def module(test_name, names):
    sub = "Test_" + test_name[:-len("Tests")]
    body = []
    for name in names:
        body += matrix_checks() if name == "Matrix" else checks_for(name, seed=sum(map(ord, name)))
    # Keep each procedure well under VBA's 64KB "Procedure too large" limit
    chunks = [body[i:i + 100] for i in range(0, len(body), 100)]
    parts = []
    for n, chunk in enumerate(chunks, 1):
        parts += ["", f"Private Sub Part{n}()", *chunk, "End Sub"]
    lines = [
        f'Attribute VB_Name = "mod{test_name}"',
        f"' Generated by tools/generate_tests.py from Python textdistance {td.__version__}.",
        "' Do not edit by hand. Run " + sub + " and check the Immediate window (Ctrl+G).",
        "Option Explicit",
        "Option Compare Binary",
        "",
        "Private mPass As Long, mFail As Long",
        "",
        f"Public Sub {sub}()",
        "    mPass = 0: mFail = 0",
        *[f"    Part{n}" for n in range(1, len(chunks) + 1)],
        f'    Debug.Print "{sub[5:]}: " & mPass & " passed, " & mFail & " failed"',
        "End Sub",
        *parts,
        *helper_lines(names),
        "",
        "Private Sub CheckL(ByVal Name As String, ByVal Actual As Long, ByVal Expected As Long)",
        "    If Actual = Expected Then",
        "        mPass = mPass + 1",
        "    Else",
        "        mFail = mFail + 1",
        '        Debug.Print "FAIL " & Name & ": got " & Actual & ", expected " & Expected',
        "    End If",
        "End Sub",
        "",
        "Private Sub CheckS(ByVal Name As String, ByVal Actual As String, ByVal Expected As String)",
        "    If StrComp(Actual, Expected, vbBinaryCompare) = 0 Then",
        "        mPass = mPass + 1",
        "    Else",
        "        mFail = mFail + 1",
        '        Debug.Print "FAIL " & Name & ": got """ & Actual & """, expected """ & Expected & """"',
        "    End If",
        "End Sub",
        "",
        "Private Sub CheckD(ByVal Name As String, ByVal Actual As Double, ByVal Expected As Double)",
        "    If Abs(Actual - Expected) < 0.000000001 Then",
        "        mPass = mPass + 1",
        "    Else",
        "        mFail = mFail + 1",
        '        Debug.Print "FAIL " & Name & ": got " & Actual & ", expected " & Expected',
        "    End If",
        "End Sub",
    ]
    size = sum(len(l) + 2 for l in lines)
    assert size < 150_000, f"{test_name} is {size} bytes; split it (large modules fail to load)"
    too_long = [l for l in lines if len(l) > 1000]
    assert not too_long, f"VBA lines are limited to 1023 characters: {too_long[0][:80]}"
    return "\r\n".join(lines) + "\r\n"


def main():
    os.makedirs(os.path.join(ROOT, "tests"), exist_ok=True)
    for test_name, names in TEST_FILES.items():
        path = os.path.join(ROOT, "tests", test_name + ".bas")
        with open(path, "w", encoding="ascii", newline="") as f:
            f.write(module(test_name, names))
        print("wrote", os.path.relpath(path, ROOT))


if __name__ == "__main__":
    main()
