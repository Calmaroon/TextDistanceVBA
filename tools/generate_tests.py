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
}

# Which test module each algorithm's tests go into
TEST_FILES = {
    "HammingTests": ["Hamming"],
    "LevenshteinTests": ["Levenshtein"],
    "DamerauLevenshteinTests": ["DamerauLevenshtein"],
    "JaroWinklerTests": ["Jaro", "JaroWinkler"],
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


def checks_for(name, seed):
    prefix, variants, integer = ALGORITHMS[name]
    check = "CheckL" if integer else "CheckD"
    lines = []
    for extra, algo in variants:
        for a, b in COMMON_CASES + random_cases(seed):
            args = f"{vba_str(a)}, {vba_str(b)}, True" + (f", {extra}" if extra else "")
            label = f"{prefix}({a!r}, {b!r}{', ' + extra if extra else ''})"
            for func, value, chk in (
                ("Distance", algo.distance(a, b), check),
                ("Similarity", algo.similarity(a, b), check),
                ("NormalizedDistance", algo.normalized_distance(a, b), "CheckD"),
                ("NormalizedSimilarity", algo.normalized_similarity(a, b), "CheckD"),
            ):
                lines.append(f'    {chk} {vba_str(label + " " + func)}, '
                             f"{prefix}_{func}({args}), {vba_num(value)}")
    # CaseSensitive:=False is a VBA addition: compare with Python on lowercased input
    extra, algo = variants[0]
    for a, b in CASE_INSENSITIVE_CASES:
        args = f"{vba_str(a)}, {vba_str(b)}, False" + (f", {extra}" if extra else "")
        label = f"{prefix}({a!r}, {b!r}, CaseSensitive=False) Distance"
        lines.append(f"    {check} {vba_str(label)}, {prefix}_Distance({args}), "
                     f"{vba_num(algo.distance(a.lower(), b.lower()))}")
    lines.append(f'    CheckD "{prefix}_Maximum", {prefix}_Maximum("test", "testing"), '
                 f'{vba_num(variants[0][1].maximum("test", "testing"))}')
    return lines


def module(test_name, names):
    sub = "Test_" + test_name[:-len("Tests")]
    body = []
    for name in names:
        body += checks_for(name, seed=sum(map(ord, name)))
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
        "Private Sub CheckD(ByVal Name As String, ByVal Actual As Double, ByVal Expected As Double)",
        "    If Abs(Actual - Expected) < 0.000000001 Then",
        "        mPass = mPass + 1",
        "    Else",
        "        mFail = mFail + 1",
        '        Debug.Print "FAIL " & Name & ": got " & Actual & ", expected " & Expected',
        "    End If",
        "End Sub",
    ]
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
