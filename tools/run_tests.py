"""Run the VBA test modules in headless LibreOffice (VBA compatibility mode).

LibreOffice Basic is not Office VBA, but with `Option VBASupport 1` it is
close enough to catch compile errors and logic bugs without Windows.

Usage:  python3 tools/run_tests.py
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

# Library modules to load (Excel_UDF needs Excel's object model, so skip it)
LIBRARY = sorted(f for f in glob.glob(os.path.join(ROOT, "*.bas"))
                 if os.path.basename(f) != "Excel_UDF.bas")
TESTS = sorted(glob.glob(os.path.join(ROOT, "tests", "*.bas")))

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
    office = subprocess.Popen([
        "soffice", "--headless", "--invisible", "--norestore", "--nologo",
        f"--accept=socket,host=localhost,port={PORT};urp;"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        ctx = connect()
        desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
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
            print(output, end="")
            summary = output.strip().splitlines()[-1] if output.strip() else ""
            if not re.search(r" 0 failed$", summary):
                failed = True
        doc.close(True)
        return 1 if failed or not test_subs else 0
    finally:
        office.terminate()
        office.wait()


if __name__ == "__main__":
    sys.exit(main())
