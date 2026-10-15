"""Run the VBA test modules in headless LibreOffice (VBA compatibility mode).

LibreOffice Basic is not Office VBA, but with `Option VBASupport 1` it is
close enough to catch compile errors and logic bugs without Windows.

Usage:  python3 tools/run_tests.py [TestName ...]   e.g. Hamming Levenshtein
Requires LibreOffice and its Python UNO bridge (python3-uno).
"""
import glob
import os
import re
import subprocess
import sys
import time

import uno

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PORT = 2083

LIBRARY = sorted(glob.glob(os.path.join(ROOT, "*.bas")))
TESTS = sorted(f for f in glob.glob(os.path.join(ROOT, "tests", "*.bas"))
               if not sys.argv[1:] or
               os.path.basename(f)[:-len("Tests.bas")] in sys.argv[1:])

HARNESS = """Option VBASupport 1
Option Explicit
Public gLog As String
Public Sub LogLine(ByVal Text As String)
    gLog = gLog & Text & Chr(10)
End Sub
Public Function RunTest(ByVal Name As String) As String
    gLog = ""
    Select Case Name
{cases}
    End Select
    RunTest = gLog
End Function
"""


# Worksheet formulas for Excel_UDF.bas (Calc uses ; as argument separator)
UDF_CELLS = {"A1": "martha", "B1": "marhta", "A2": 12345, "B2": 12354,
             # Matrix lookup table in F1:H3
             "F1": "A", "G1": "A", "H1": 10, "F2": "A", "G2": "G", "H2": -1,
             "F3": "G", "G3": "G", "H3": 7}
UDF_FORMULAS = [
    ('=HAMMING("test";"testing")', 3), ('=HAMMING_SIM("test";"text")', 3),
    ('=HAMMING_NSIM("Test";"text";0)', 0.75),
    ('=LEVENSHTEIN("kitten";"sitting")', 3), ('=LEVENSHTEIN(A2;B2)', 2),
    ('=LEVENSHTEIN(A1;C9)', 6), ('=LEVENSHTEIN_NDIST("kitten";"sitting")', 3 / 7),
    ('=DAMERAU_LEVENSHTEIN("ca";"abc")', 3), ('=DAMERAU_LEVENSHTEIN("ca";"abc";1;0)', 2),
    ('=JARO(A1;B1)', 0.9444444444444445), ('=JARO_DIST(A1;B1)', 0.0555555555555555),
    ('=JARO_WINKLER(A1;B1)', 0.9611111111111111), ('=JARO_WINKLER_DIST(A1;B1)', 0.0388888888888889),
    ('=LCSSEQ("test";"text")', "tet"), ('=LCSSEQ_SIM("test";"text")', 3),
    ('=LCSSEQ_NDIST("test";"text")', 0.25), ('=LCSSTR("test";"text")', "te"),
    ('=LCSSTR_DIST(A1;B1)', 3), ('=LCSSTR("ABC";"zbc";0)', "bc"),
    ('=RATCLIFF_OBERSHELP("WIKIMEDIA";"WIKIMANIA")', 7 / 9),
    ('=RATCLIFF_OBERSHELP_DIST("WIKIMEDIA";"WIKIMANIA")', 2 / 9),
    ('=PREFIX("test";"text")', "te"), ('=PREFIX_NSIM("test";"text")', 0.5),
    ('=POSTFIX("ring";"sing")', "ing"), ('=POSTFIX_DIST("ring";"sing")', 1),
    ('=LENGTH("test";"testing")', 3), ('=LENGTH_SIM("test";"testing")', 4),
    ('=IDENTITY("abc";"ABC";0)', 1), ('=IDENTITY_DIST("abc";"ABC")', 1),
    ('=MATRIX("A";"A")', 1), ('=MATRIX("G";"A";F1:H3)', -1),
    ('=MATRIX("G";"A";F1:H3;0;1;0)', 0), ('=MATRIX_DIST("A";"A";F1:H3;0;10)', 0),
    ('=MRACODEX("Catherine")', "CTHRN"), ('=MRA("Byrne";"Boern")', 1),
    ('=MRA_DIST("Smith";"Smyth")', 3), ('=EDITEX("Byrne";"Boern")', 4),
    ('=EDITEX_NSIM("Smith";"smith")', 1),
    ('=COSINE("nelson";"neilsen")', 0.7715167498104595), ('=COSINE("Smith John";"John Smith";1;0)', 1),
    ('=COSINE_DIST("Smith John";"john smith";0;0)', 0),
    ('=JACCARD("night";"nacht")', 0.42857142857142855),
    ('=SORENSEN("night";"nacht";1;2)', 0.25),
    ('=OVERLAP("John Smith";"Smith";1;0)', 1.0),
    ('=TANIMOTO_DIST("ab";"ac")', 2.584962500721156),
    ('=TVERSKY("test";"text";1;1;0;0.5;2)', 0.5454545454545454),
    ('=TVERSKY("test";"text";1;1;0;1;1;1)', 0.5),
    ('=MONGE_ELKAN("hello world";"hello word";1;0)', 2.25),
    ('=BAG("abc";"abd")', 1),
]


def check_udfs(doc):
    sheet = doc.Sheets.getByIndex(0)
    for ref, value in UDF_CELLS.items():
        cell = sheet.getCellRangeByName(ref)
        if isinstance(value, str):
            cell.String = value
        else:
            cell.Value = value
    for row, (formula, _) in enumerate(UDF_FORMULAS):
        sheet.getCellByPosition(3, row).Formula = formula
    doc.calculateAll()
    failures = 0
    for row, (formula, expected) in enumerate(UDF_FORMULAS):
        cell = sheet.getCellByPosition(3, row)
        if isinstance(expected, str):
            ok = not cell.getError() and cell.String == expected
        else:
            ok = not cell.getError() and abs(cell.Value - expected) <= 1e-9
        if not ok:
            failures += 1
            print(f"FAIL {formula}: got {cell.String}, expected {expected}")
    print(f"Excel UDFs: {len(UDF_FORMULAS) - failures} passed, {failures} failed")
    return failures == 0


def enable_macros(ctx):
    """Allow document macros so Calc formulas can call the Basic functions."""
    from com.sun.star.beans import PropertyValue
    provider = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.configuration.ConfigurationProvider", ctx)
    node = PropertyValue(Name="nodepath",
                         Value="/org.openoffice.Office.Common/Security/Scripting")
    access = provider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationUpdateAccess", (node,))
    access.setPropertyValue("MacroSecurityLevel", 0)
    access.commitChanges()
    return (PropertyValue(Name="MacroExecutionMode", Value=4),)  # ALWAYS_EXECUTE_NO_WARN


def check_names():
    """VBA names ignore case: duplicate public names across modules are ambiguous."""
    seen = {}
    ok = True
    for path in LIBRARY:
        with open(path, encoding="cp1252") as f:
            src = f.read()
        found = [re.search(r'VB_Name = "(\w+)"', src).group(1)]
        found += re.findall(r"^Public (?:Function|Sub) (\w+)", src, flags=re.M)
        for name in found:
            other = seen.setdefault(name.lower(), os.path.basename(path))
            if other != os.path.basename(path):
                print(f"FAIL duplicate name {name}: {other} and {os.path.basename(path)}")
                ok = False
    return ok


def load_source(path, is_test):
    with open(path, encoding="cp1252") as f:
        src = f.read().replace("\r\n", "\n")
    name = re.search(r'Attribute VB_Name = "(\w+)"', src).group(1)
    src = re.sub(r"^Attribute .*\n", "", src, flags=re.M)
    if is_test:
        src = src.replace("Debug.Print ", "LogLine ")
    return name, "Option VBASupport 1\n" + src


def connect():
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local)
    url = f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext"
    for _ in range(60):
        try:
            return resolver.resolve(url)
        except Exception:
            time.sleep(0.5)
    sys.exit("Could not connect to LibreOffice")


def main():
    if not check_names():
        return 1
    office = subprocess.Popen([
        "soffice", "--headless", "--invisible", "--norestore", "--nologo",
        f"--accept=socket,host=localhost,port={PORT};urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        ctx = connect()
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0,
                                           enable_macros(ctx))
        libs = doc.BasicLibraries
        if not libs.hasByName("Standard"):
            libs.createLibrary("Standard")
        lib = libs.getByName("Standard")

        test_subs = []
        for path in LIBRARY + TESTS:
            is_test = path in TESTS
            name, src = load_source(path, is_test)
            lib.insertByName(name, src)
            if is_test:
                test_subs += re.findall(r"^Public Sub (Test_\w+)\(\)", src, flags=re.M)

        cases = "\n".join(f'        Case "{s}": {s}' for s in test_subs)
        lib.insertByName("TestHarness", HARNESS.format(cases=cases))

        provider = doc.getScriptProvider()
        script = provider.getScript(
            "vnd.sun.star.script:Standard.TestHarness.RunTest"
            "?language=Basic&location=document")

        failed = False
        for sub in test_subs:
            output = script.invoke((sub,), (), ())[0]
            print(output if output.strip() else f"{sub}: no output (Basic error?)\n", end="")
            summary = output.strip().splitlines()[-1] if output.strip() else ""
            if not re.search(r" 0 failed$", summary):
                failed = True
        if not sys.argv[1:] and not check_udfs(doc):
            failed = True
        doc.close(True)
        return 1 if failed or not test_subs else 0
    finally:
        office.terminate()
        office.wait()


if __name__ == "__main__":
    sys.exit(main())
