# Benchmarks

Performance comparison of opensheet_core (Rust-backed) versus openpyxl (pure Python) across multiple dataset sizes.

For details on the benchmarking methodology (interleaved runs, subprocess isolation, memory measurement), see [benchmarking.md](benchmarking.md).

---

## Overview

All benchmarks use datasets with 10 columns of mixed types (strings, integers, floats, booleans) and vary the row count from 1,000 to 100,000. Each configuration is measured with 5 interleaved runs. The minimum time and median RSS delta are reported.

## Read Performance

Reading XLSX files into Python lists. Both libraries read the same openpyxl-generated file to avoid format-specific advantages.

![Read time comparison](assets/bench_read_time.svg)

| Dataset | opensheet_core | openpyxl | Speedup |
|---------|---------------|----------|---------|
| 1K rows x 10 cols | -- | -- | -- |
| 10K rows x 10 cols | -- | -- | -- |
| 50K rows x 10 cols | -- | -- | -- |
| 100K rows x 10 cols | -- | -- | -- |

> Run `python benchmarks/bench_visualize.py` to populate the table above with your own results.

### Scaling analysis

Read performance scales roughly linearly with dataset size for both libraries. The speedup ratio of opensheet_core over openpyxl tends to increase with larger datasets, since the Rust parsing engine amortizes its fixed overhead across more rows while openpyxl's per-row Python overhead grows proportionally.

## Write Performance

Writing XLSX files from Python data. openpyxl uses `write_only=True` (its streaming/fastest mode) for a fair comparison.

![Write time comparison](assets/bench_write_time.svg)

| Dataset | opensheet_core | openpyxl | Speedup |
|---------|---------------|----------|---------|
| 1K rows x 10 cols | -- | -- | -- |
| 10K rows x 10 cols | -- | -- | -- |
| 50K rows x 10 cols | -- | -- | -- |
| 100K rows x 10 cols | -- | -- | -- |

> Run `python benchmarks/bench_visualize.py` to populate the table above with your own results.

### Scaling analysis

Write speedups are typically more modest than read speedups. Both libraries stream data to disk, so the bottleneck shifts toward Python-side row generation and data serialization. The Rust writer still wins on XML serialization and ZIP compression, but the per-row data marshalling from Python narrows the gap.

## Speedup Over openpyxl

How the speedup ratio changes as dataset size grows.

![Speedup chart](assets/bench_speedup.svg)

## Memory Usage

RSS (Resident Set Size) delta measured via platform-specific APIs -- not the high-water mark. See [benchmarking.md](benchmarking.md) for details on why this matters.

![Memory usage comparison](assets/bench_memory.svg)

| Dataset | opensheet_core (read) | openpyxl (read) | opensheet_core (write) | openpyxl (write) |
|---------|----------------------|-----------------|----------------------|------------------|
| 1K rows x 10 cols | -- | -- | -- | -- |
| 10K rows x 10 cols | -- | -- | -- | -- |
| 50K rows x 10 cols | -- | -- | -- | -- |
| 100K rows x 10 cols | -- | -- | -- | -- |

> Run `python benchmarks/bench_visualize.py` to populate the table above with your own results.

### Memory optimization techniques

OpenSheet Core uses several techniques to minimize memory during reads:

1. **Deferred shared-string resolution** -- Shared strings are stored as integer indices during XML parsing rather than cloned string values.
2. **Pre-interned Python strings** -- The shared string table is converted to Python objects once. Cells referencing the same string reuse the existing Python object.
3. **Convert-and-drop** -- Rust row data is consumed during Python conversion. As each row is converted, the Rust memory is freed immediately.
4. **Single-sheet parsing** -- `read_sheet()` only parses the requested worksheet.

## How to Reproduce

### Prerequisites

```bash
pip install openpyxl matplotlib
```

### Run benchmarks and generate charts

```bash
# Full run: 4 dataset sizes, 5 interleaved runs each (takes several minutes)
python benchmarks/bench_visualize.py

# Quick test (1 run per config, useful for verifying the script works)
python benchmarks/bench_visualize.py --quick

# Custom number of runs
python benchmarks/bench_visualize.py --runs 7
```

### Regenerate charts from cached results

If you have already run the benchmarks and just want to regenerate the charts (for example, after tweaking chart styling):

```bash
python benchmarks/bench_visualize.py --no-run
```

### Run individual benchmark scripts

For more detailed output on a specific operation:

```bash
python benchmarks/bench_read.py     # Read benchmarks across multiple sizes
python benchmarks/bench_write.py    # Write benchmarks across multiple sizes
python benchmarks/benchmark.py      # Combined read/write on a single size
```

### Output files

| File | Description |
|------|-------------|
| `docs/assets/benchmark_results.json` | Raw timing and memory data (JSON) |
| `docs/assets/bench_read_time.svg` | Read performance bar chart |
| `docs/assets/bench_write_time.svg` | Write performance bar chart |
| `docs/assets/bench_memory.svg` | Memory usage comparison (read and write) |
| `docs/assets/bench_speedup.svg` | Speedup ratio line chart |

PNG versions of each chart are also generated alongside the SVG files.

## Factors That Affect Results

Benchmark results vary across machines and environments. When sharing or comparing results, note:

- **Hardware**: CPU model, RAM amount, and disk type (SSD vs HDD) all affect results. Write benchmarks are especially sensitive to disk I/O.
- **Operating system**: macOS, Linux, and Windows have different I/O and memory subsystems.
- **Python version**: CPython 3.11+ has measurably faster bytecode execution. PyPy is not tested.
- **Background processes**: Other running applications introduce noise. The interleaved measurement strategy mitigates but does not eliminate this.
- **Thermal throttling**: Long benchmark runs on laptops may trigger CPU throttling.

The JSON results file (`docs/assets/benchmark_results.json`) records the platform, Python version, and library versions for reproducibility.
