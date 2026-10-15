# TextDistanceVBA
Implementation of Python's [textdistance](https://github.com/life4/textdistance) library within VBA.

The algorithm modules are pure VBA with no host references, so they work in Excel, Access, Word and other Office applications. Results match Python's textdistance.

## Installation
In the VBA editor (Alt+F11) choose **File > Import File...** and import:

- the algorithm modules you need (e.g. `Hamming.bas`)
- `Excel_UDF.bas` if you want worksheet functions (Excel only)
- optionally `tests/*.bas` to check the results

## Algorithms

| Algorithm | Module | Status |
|---|---|---|
| Hamming | `Hamming.bas` | Done |

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

In Access you can call these directly in queries, e.g. `SELECT Hamming_Distance([Name1], [Name2]) FROM ...`.

## Excel worksheet functions
With `Excel_UDF.bas` imported:

| Worksheet function | Library function |
|---|---|
| `=HAMMING(A1, B1, [CaseSensitive], [Truncate])` | `Hamming_Distance` |
| `=HAMMING_SIM(...)` | `Hamming_Similarity` |
| `=HAMMING_NDIST(...)` | `Hamming_NormalizedDistance` |
| `=HAMMING_NSIM(...)` | `Hamming_NormalizedSimilarity` |

They accept cell references, numbers and blank cells, and return `#VALUE!` for errors or multi-cell ranges.

## Tests
`tests/HammingTests.bas` contains expected values generated from Python textdistance. Import it, run `Test_Hamming`, and check the Immediate window (Ctrl+G) for the pass/fail summary.
