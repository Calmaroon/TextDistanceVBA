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
| Longest common subsequence (LCSSeq) | `LCSSeq.bas` | Done |
| Longest common substring (LCSStr) | `LCSStr.bas` | Done |
| Ratcliff-Obershelp | `RatcliffObershelp.bas` | Done |
| Prefix, Postfix, Length, Identity, Matrix | `Simple.bas` | Done |
| MRA, Editex | `Phonetic.bas` | Done |

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

### LCSSeq and LCSStr
Longest common subsequence (characters in the same order, gaps allowed) and longest common substring (consecutive characters). Similarity is the length of the match; `Distance = Maximum - Similarity`. `LCSSeq_Sequence` and `LCSStr_Substring` return the matching text itself, picking the same one as Python when there are several. Optional argument: `CaseSensitive`.

```vba
Debug.Print LCSSeq_Sequence("test", "text")      ' tet
Debug.Print LCSSeq_Similarity("test", "text")    ' 3
Debug.Print LCSStr_Substring("test", "text")     ' te
Debug.Print LCSStr_Similarity("test", "text")    ' 2
```

### Ratcliff-Obershelp
Gestalt pattern matching, the idea behind Python's `difflib.SequenceMatcher.ratio()`. Similarity from 0 to 1, `Distance = 1 - Similarity`. Optional argument: `CaseSensitive`.

```vba
Debug.Print RatcliffObershelp_Similarity("WIKIMEDIA", "WIKIMANIA")   ' 0.777...
```

### Simple measures (`Simple.bas`)
| Measure | Similarity | Example |
|---|---|---|
| Prefix | Length of the common start | `Prefix_CommonPrefix("test", "text")` = `"te"` |
| Postfix | Length of the common end | `Postfix_CommonPostfix("ring", "sing")` = `"ing"` |
| Length | Length of the shorter string (distance = difference in length) | `Length_Distance("test", "testing")` = 3 |
| Identity | 1 if identical, else 0 | `Identity_Similarity("a", "a")` = 1 |
| Matrix | Score looked up in a table | see below |

`Matrix_Similarity(s1, s2, [Table], [MismatchCost], [MatchCost], [Symmetric], [CaseSensitive])` looks the pair up in `Table`, a 2-D array or Excel range with 3 columns (string 1, string 2, score). With `Symmetric` (default `True`) it also tries `(s2, s1)`. If the pair isn't found it returns `MatchCost` (default 1) when the strings are equal, else `MismatchCost` (default 0). If a pair appears more than once, the last row wins. The maximum is `MatchCost`.

```vba
' Table in E1:G3:  A | A | 10,  A | G | -1,  G | G | 7
Debug.Print Matrix_Similarity("G", "A", Range("E1:G3"))   ' -1 (symmetric lookup)
```

### Phonetic measures (`Phonetic.bas`)
Both compare upper-cased text, so they are always case-insensitive and have no `CaseSensitive` argument.

- **MRA** (Match Rating Approach, the Western Airlines surname-matching method). Each name becomes a short codex: vowels after the first letter are removed, repeated letters collapsed, and anything over 6 characters is cut to the first 3 + last 3. `MRA_Similarity` compares the codexes (0 to `MRA_Maximum`, the longer codex length), and is 0 when the codex lengths differ by more than 2.
- **Editex** is an edit distance where replacing letters from the same sound group (vowels, B/P, C/K/Q, D/T, L/R, M/N, G/J, F/P/V, S/X/Z, C/S/Z) costs less than unrelated letters. Optional arguments: `LocalMode` (default `False`), `MatchCost` (0), `GroupCost` (1), `MismatchCost` (2). The maximum is the longer length times `MismatchCost`.

```vba
Debug.Print MRA_Codex("Catherine")              ' CTHRN
Debug.Print MRA_Similarity("Byrne", "Boern")    ' 1
Debug.Print Editex_Distance("Byrne", "Boern")   ' 4
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
| `=LCSSEQ(A1, B1, [CaseSensitive])` | `LCSSeq_Sequence` (returns text) |
| `=LCSSEQ_SIM(...)`, `_DIST`, `_NDIST`, `_NSIM` | `LCSSeq_*` |
| `=LCSSTR(A1, B1, [CaseSensitive])` | `LCSStr_Substring` (returns text) |
| `=LCSSTR_SIM(...)`, `_DIST`, `_NDIST`, `_NSIM` | `LCSStr_*` |
| `=RATCLIFF_OBERSHELP(A1, B1, [CaseSensitive])` | `RatcliffObershelp_Similarity` |
| `=RATCLIFF_OBERSHELP_DIST(...)` | `RatcliffObershelp_Distance` |
| `=PREFIX(A1, B1, [CaseSensitive])`, `=POSTFIX(...)` | Common prefix / postfix (returns text) |
| `=PREFIX_SIM(...)`, `_DIST`, `_NDIST`, `_NSIM` (same for `POSTFIX`) | `Prefix_*`, `Postfix_*` |
| `=LENGTH(A1, B1)`, `_SIM`, `_NDIST`, `_NSIM` | `Length_*` |
| `=IDENTITY(A1, B1, [CaseSensitive])`, `=IDENTITY_DIST(...)` | `Identity_Similarity`, `Identity_Distance` |
| `=MATRIX(A1, B1, [Table], [MismatchCost], [MatchCost], [Symmetric], [CaseSensitive])`, `=MATRIX_DIST(...)` | `Matrix_Similarity`, `Matrix_Distance` |
| `=MRACODEX(A1)` | `MRA_Codex` (returns text) |
| `=MRA(A1, B1)`, `_DIST`, `_NDIST`, `_NSIM` | `MRA_*` |
| `=EDITEX(A1, B1, [LocalMode], [MatchCost], [GroupCost], [MismatchCost])`, `_SIM`, `_NDIST`, `_NSIM` | `Editex_*` |

The bare name returns the same value as Python's `textdistance.<name>(a, b)`. Cells, numbers and blanks are converted to text; multi-cell ranges and error values give `#VALUE!`.

## Tests
The modules in `tests/` contain expected values generated from Python textdistance (`tools/generate_tests.py`).

- **In Office:** import a test module, run its `Test_*` sub (e.g. `Test_Levenshtein`), and check the Immediate window (Ctrl+G) for the pass/fail summary.
- **Without Office:** `python3 tools/run_tests.py` runs every test module, plus worksheet formulas for `Excel_UDF.bas`, in headless LibreOffice (VBA compatibility mode).

After changing an algorithm or adding one, regenerate the tests with `pip install textdistance` and `python3 tools/generate_tests.py`.
