# TextDistanceVBA
Implementation of Python's [textdistance](https://github.com/life4/textdistance) library within VBA.

The algorithm modules are pure VBA with no host references, so they work in Excel, Access, Word and other Office applications. Results match Python's textdistance.

## Installation
In the VBA editor (Alt+F11) choose **File > Import File...** and import:

- the algorithm modules you need (e.g. `Levenshtein.bas`); each one works on its own
- `Excel_UDF.bas` if you want worksheet functions (needs all algorithm modules)
- optionally `tests/*.bas` to check the results

## Algorithms

| Algorithm | Module | Status |
|---|---|---|
| Hamming | `Hamming.bas` | Done |
| Levenshtein | `Levenshtein.bas` | Done |
| Damerau-Levenshtein | `DamerauLevenshtein.bas` | Done |
| Jaro, Jaro-Winkler | `JaroWinkler.bas` | Done |

Every algorithm exposes the same five functions, prefixed with its name:

| Function | Meaning |
|---|---|
| `<Name>_Distance(s1, s2, ...)` | Raw distance (0 = identical) |
| `<Name>_Similarity(s1, s2, ...)` | `Maximum - Distance` |
| `<Name>_NormalizedDistance(s1, s2, ...)` | Distance scaled to 0..1 |
| `<Name>_NormalizedSimilarity(s1, s2, ...)` | Similarity scaled to 0..1 |
| `<Name>_Maximum(s1, s2)` | Largest possible distance |

### Hamming
Counts the positions where the strings differ. Extra characters in the longer string count as differences.

Optional arguments: `CaseSensitive` (default `True`) and `Truncate` (default `False`; when `True`, only the first `Min(Len(s1), Len(s2))` characters are compared).

```vba
Debug.Print Hamming_Distance("test", "text")                 ' 1
Debug.Print Hamming_Distance("test", "testing")              ' 3
Debug.Print Hamming_Distance("Hello", "hello", False)        ' 0
Debug.Print Hamming_NormalizedSimilarity("test", "text")     ' 0.75
```

### Levenshtein
Minimum number of single-character insertions, deletions and substitutions. Optional argument: `CaseSensitive`.

```vba
Debug.Print Levenshtein_Distance("kitten", "sitting")            ' 3
Debug.Print Levenshtein_NormalizedSimilarity("kitten", "sitting") ' 0.571...
```

### Damerau-Levenshtein
Like Levenshtein, but swapping two adjacent characters counts as one edit. Optional arguments: `CaseSensitive` and `Restricted` (default `True`, the "optimal string alignment" variant; `False` gives the unrestricted distance).

```vba
Debug.Print DamerauLevenshtein_Distance("abcd", "acbd")           ' 1
Debug.Print DamerauLevenshtein_Distance("ca", "abc")              ' 3
Debug.Print DamerauLevenshtein_Distance("ca", "abc", True, False) ' 2
```

### Jaro and Jaro-Winkler
Similarity from 0 to 1 based on nearby matching characters and transpositions. Jaro-Winkler boosts strings that share a prefix, which suits names. Here similarity is the primary value, the maximum is always 1, and `Distance = 1 - Similarity`.

Optional arguments: `CaseSensitive`, `LongTolerance` (default `False`) and, for Jaro-Winkler, `PrefixWeight` (default `0.1`).

```vba
Debug.Print Jaro_Similarity("martha", "marhta")          ' 0.944...
Debug.Print JaroWinkler_Similarity("martha", "marhta")   ' 0.961...
```

In Access you can call these directly in queries, e.g. `SELECT Hamming_Distance([Name1], [Name2]) FROM ...`.

## Excel worksheet functions
With `Excel_UDF.bas` imported:

| Worksheet function | Library function |
|---|---|
| `=HAMMING(A1, B1, [CaseSensitive], [Truncate])` | `Hamming_Distance` |
| `=HAMMING_SIM(...)` | `Hamming_Similarity` |
| `=HAMMING_NDIST(...)` | `Hamming_NormalizedDistance` |
| `=HAMMING_NSIM(...)` | `Hamming_NormalizedSimilarity` |
| `=LEVENSHTEIN(A1, B1, [CaseSensitive])`, `_SIM`, `_NDIST`, `_NSIM` | `Levenshtein_*` |
| `=DAMERAU_LEVENSHTEIN(A1, B1, [CaseSensitive], [Restricted])`, `_SIM`, `_NDIST`, `_NSIM` | `DamerauLevenshtein_*` |
| `=JARO(A1, B1, [CaseSensitive], [LongTolerance])` | `Jaro_Similarity` |
| `=JARO_DIST(...)` | `Jaro_Distance` |
| `=JARO_WINKLER(A1, B1, [CaseSensitive], [LongTolerance], [PrefixWeight])` | `JaroWinkler_Similarity` |
| `=JARO_WINKLER_DIST(...)` | `JaroWinkler_Distance` |

The bare name returns the same value as Python's `textdistance.<name>(a, b)`. Cells, numbers and blanks are converted to text; multi-cell ranges and error values give `#VALUE!`.

## Tests
The modules in `tests/` contain expected values generated from Python textdistance (`tools/generate_tests.py`).

- **In Office:** import a test module, run its `Test_*` sub (e.g. `Test_Levenshtein`), and check the Immediate window (Ctrl+G) for the pass/fail summary.
- **Without Office:** `python3 tools/run_tests.py` runs every test module, plus worksheet formulas for `Excel_UDF.bas`, in headless LibreOffice (VBA compatibility mode).

After changing an algorithm or adding one, regenerate the tests with `pip install textdistance` and `python3 tools/generate_tests.py`.
