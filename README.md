# Excel Quality and Six Sigma Functions

An Excel add-in containing custom worksheet functions for quality management, process yield, Taguchi loss, and Six Sigma calculations.

## Included functions

| Function | Purpose |
|---|---|
| `QMS_STANDARD_ERROR` | Calculates the standard error of a sample mean |
| `QMS_TAGUCHI_LOSS` | Calculates Taguchi loss for one measurement |
| `QMS_TAGUCHI_PROCESS_LOSS` | Estimates average Taguchi loss for a process sample |
| `QMS_THROUGHPUT_YIELD` | Calculates yield for one process step |
| `QMS_ROLLED_YIELD` | Multiplies the yields from multiple process stages |
| `QMS_DPMO` | Calculates defects per million opportunities |
| `QMS_SIGMA_LEVEL` | Converts DPMO to an estimated sigma level |

## Installation

1. Download `Quality_Six_Sigma_Functions.xlam` from the latest GitHub Release.
2. Open Excel.
3. Select **File → Options → Add-ins**.
4. Next to **Manage Excel Add-ins**, select **Go**.
5. Select **Browse** and locate the downloaded `.xlam` file.
6. Ensure the add-in is checked, then select **OK**.
7. Restart Excel if the functions do not appear immediately.

Excel might display a security warning because the VBA project is not digitally signed. Only enable macros after confirming that the file came from this repository.

## Examples

Standard error:

```excel
=QMS_STANDARD_ERROR(A2:A26)
```

Taguchi loss for one measurement:

```excel
=QMS_TAGUCHI_LOSS(A2,6,15,1.25)
```

Estimated Taguchi process loss:

```excel
=QMS_TAGUCHI_PROCESS_LOSS(A2:A26,6,15,1.25)
```

Throughput yield:

```excel
=QMS_THROUGHPUT_YIELD(970,1000)
```

Rolled throughput yield:

```excel
=QMS_ROLLED_YIELD(B2:B5)
```

DPMO:

```excel
=QMS_DPMO(12,500,4)
```

Estimated sigma level:

```excel
=QMS_SIGMA_LEVEL(6000)
```

## Input conventions

- Enter yields as decimals, such as `0.98` or `98%`.
- Taguchi tolerance must be greater than zero.
- `QMS_TAGUCHI_PROCESS_LOSS` requires at least two numeric observations.
- The default sigma shift is `1.5`. Enter `0` as the second argument to calculate without a shift.

## Source code

The VBA source is available in [`src/QualityFunctions.bas`](src/QualityFunctions.bas).

## License

Released under the MIT License.
