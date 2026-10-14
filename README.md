# PowerPoint dataset extractor

Extract data with descriptions and slide provenance, using ordinary Python. Runs on the Databricks driver without Microsoft Office, LibreOffice, Spark jobs, or a model endpoint.

## Start in Databricks

1. Import **PowerPoint_Data_Extractor.ipynb** as a notebook.
2. Run the dependency and restart cells, then the remaining cells. Leaving `INPUT_PATH` blank runs the built-in side-by-side chart demo. Set `SAMPLE_MODE="single_chart"` for the original revenue/table/notes example.
3. Set `INPUT_PATH` to your file, for example `/Volumes/main/analytics/documents/presentations/sales.pptx`.
4. Set `OUTPUT_DIR` to save JSON and JSONL. Optionally enable the Delta cell and choose a table you can write.

The notebook contains the entire extractor. You do not need to upload the module separately. The module is provided for reuse in jobs or other notebooks. Extract one deck at a time; this implementation is bounded driver-side processing, not a distributed ingestion system.

Python 3.10+ is required. Unity Catalog Volume access requires compatible compute and Databricks Runtime 13.3 LTS or later. The Volume must already exist. `dbfs:/Volumes/...` is normalized to `/Volumes/...`. Legacy `dbfs:/...` is supported only where the `/dbfs` FUSE mount is available. Stage cloud URIs into a Volume or supply bytes; the extractor does not create cloud clients.

## Generic Python

```bash
pip install -r requirements.txt
```

```python
from pptx_extract import extract_pptx, write_outputs

report = extract_pptx("presentation.pptx")
paths = write_outputs(report, "extracted")

for record in report["records"]:
    if record["kind"] in {"chart", "table", "embedded_worksheet"}:
        print(record["what_data_represents"])
        print(record["data"])
```

`extract_pptx(pptx_bytes, source_name="sales.pptx")` also accepts in-memory bytes, including bytes already retrieved through a storage SDK or Spark binary-file read.

## Output: description followed by data

Each record contains `what_data_represents` followed by a `data` payload, with source fields that let you trace it back to the original deck. This abbreviated example uses the included synthetic sample:

```json
{
  "slide_number": 1,
  "slide_title": "Quarterly product revenue (USD millions)",
  "kind": "chart",
  "what_data_represents": "Chart 1 of 1: 'Revenue by quarter' on slide 1; series: Product A, Product B; axes: USD millions",
  "status": "saved_values",
  "data": {
    "rows": [
      {"series_name": "Product A", "point_index": 0, "category": "Q1", "value": 12},
      {"series_name": "Product A", "point_index": 1, "category": "Q2", "value": 15}
    ]
  }
}
```

The real output also retains chart axes, series reference formulas, number formats, original numeric strings, and warning messages. Numbers formatted as percentages remain stored fractions; numeric chart dates remain serials with the date system and format retained. No automatic currency, percentage, or date interpretation is applied to chart caches. Workbook dates decoded by openpyxl become ISO strings and retain raw XML values in cell metadata.

Descriptions are composed from source labels and context. Missing meanings, units, denominators, or populations are not invented. `description_evidence` stores slide title/text and alternative text where available. The chart title, series labels, and axis labels are in `data`. If source labels are weak, the description will be correspondingly limited.

| Record kind | Data payload |
| --- | --- |
| `chart` | Long-form `rows`, series names, chart types, categories or x/y/bubble values, axes, reference formulas, and raw values |
| `table` | Original rows as strings, positional columns, first-row text, header styling, and merge attributes |
| `embedded_worksheet` | Stored worksheet rectangle, Excel column letters, original row offset, labels, formulas, cached values, formats, hidden flags, and source references |
| `slide_text` / `speaker_notes` | Text |
| `diagram_text` | SmartArt labels without inferred geometry or relationships |
| `image_reference` | Media part and alternative text; no pixels or OCR results |
| `external_link` | Link target and type; never fetched |
| `extraction_issue` / `unsupported_embedding` | Explicit failure or unsupported-content details |

Native tables preserve every row, including the first. Worksheets do not automatically promote a header row because workbooks may contain titles, multiple tables, blank rows, or duplicate headers. Blank leading/trailing rows are trimmed to stored cells; internal blanks and coordinate offsets remain. Merged cells are not forward-filled. Sparse hierarchical chart labels are retained in XML order, with omitted labels left null.

Full worksheets are separate records from plotted chart series and can include unplotted/hidden data. A workbook shared by several slides has `slide_number=null` and all links in `data.source_references`. Do not sum chart and workbook records together as if they were independent measurements.

## Files and Spark

`write_outputs` writes:

- **JSON report:** complete report, dataset records, deck-level warnings, and scope.
- **JSONL:** one full record per line, for downstream processing.

Filenames contain the input stem and content hash. Repeating extraction of the same named content replaces the corresponding output files. Record IDs are deterministic for the same deck, extraction options, and extractor version; they are not persistent identities across edits or configuration changes.

```python
from pptx_extract import to_spark_dataframe

datasets_df = to_spark_dataframe(spark, report)
display(datasets_df.select("slide_number", "kind", "what_data_represents", "data_json"))
```

The Spark/Delta table has one row per extracted object, with typed metadata and JSON string payloads. `data_json` is the dataset; `record_json` preserves the complete source record. An explicit Spark schema avoids inference failures caused by numbers, strings, dates, empty records, or all-null columns. These JSON strings are intentionally flexible; promote selected datasets into domain-specific tables after checking their meaning and units. The notebook includes a chart-row parsing example and an optional Delta upsert. A batch extraction failure is an envelope returned by `iter_pptx_reports`, not a successful empty extraction.

Version 1.1 adds chart position/index and uncertainty columns. Existing Delta tables created with version 1.0 need an explicit schema migration or a new table name. The notebook checks for missing columns before upserting so new quality fields cannot be silently discarded.

## Multiple charts on a slide

Every native chart instance gets a separate record, even if several charts share a slide or chart part. Chart series are associated with their own chart relationship, never pooled across the slide. Records retain:

- `record_id`, `shape_id`, and `source_part` for object identity.
- `chart_index_on_slide` (one-based source object order) and `chart_count_on_slide`.
- `bounding_box` in slide EMU coordinates and a coarse `position_on_slide`, such as `middle-left` or `middle-right`. Indices are not a visual reading order.
- Its own title, axes, series, numbers, alternative text, and candidate nearby headings.

Unrotated nested group translations/scales are handled. Rotated/flipped group positions and missing/inherited geometry are left unknown; chart data is still extracted. Nearby headings are selected conservatively using horizontal overlap and distance above the chart. They are stored as candidates, not authoritative assignments. A heading spanning several charts can appear as a candidate for each and needs review. Whole-slide text is marked as shared context on multi-chart slides.

Workbooks retain every chart/slide reference in `data.source_references`. A workbook shared by multiple charts is not arbitrarily attributed to the first chart. Multiple plots within one combination-chart object remain series within that one chart record. Image-only charts still require a separate vision/OCR workflow.

## Explicit uncertainty flags

Each record has `is_uncertain`, `needs_review`, `uncertainty_reasons`, and `quality`. For example:

```json
{
  "chart_index_on_slide": 2,
  "position_on_slide": "middle-right",
  "is_uncertain": true,
  "needs_review": true,
  "uncertainty_reasons": [{
    "code": "NEARBY_LABEL_UNVERIFIED",
    "dimension": "meaning",
    "detail": "Nearby text is a candidate label; spatial proximity does not prove it belongs to this chart."
  }],
  "quality": {
    "data_extraction": "no_detected_issue",
    "meaning": "ambiguous",
    "source_freshness": "unverified_saved_values"
  }
}
```

This example demonstrates a key distinction: the numbers can be extracted without a detected problem while their meaning remains ambiguous. Both booleans currently select records with one or more detected uncertainty flags. Rules cover missing chart-specific labels on multi-chart slides, duplicate chart labels, absent series/category labels, unresolved chart positions, missing/mismatched points, Excel error values, missing formula results, uncertain table headers/context, unlinked or unlabeled workbooks, unsupported content, and inferred callback descriptions. `uncertainty_reasons` supplies a code, dimension, and explanation for every flag. The report includes `uncertain_record_count` and `review_record_count`.

```python
uncertain = [r for r in report["records"] if r["is_uncertain"]]
```

These are deterministic diagnostics, **not calibrated probabilities or a guarantee that every ambiguity is detected**. Missing chart points can be intentional blanks. An unflagged record means none of the implemented rules found a problem. Chart/formula freshness is always unverified and reported separately in `quality.source_freshness` and warnings; saved caches alone do not force every otherwise well-labeled record into the review queue. No automated source-vs-workbook reconciliation or comprehensive semantic verification is claimed.

## Optional semantic descriptions

`enrich_descriptions` accepts a function of your choice. No model calls occur by default. You can connect an approved Databricks endpoint, local model, or other service through that callback. It receives a copy of the record, potentially including hidden workbook data, and returns a string. The output is marked `user_callback_inferred`, and the original source-label description is retained. It cannot alter the numeric payload through its input argument.

```python
from pptx_extract import enrich_descriptions

def describe(record):




    return record["what_data_represents"]

enriched = enrich_descriptions(report, describe)
```

This hook is for labeling machine-readable datasets. Recovering data from screenshots requires a separate image/slide rendering and vision workflow; this code inventories those objects so you can route them for review.

## Coverage and limits

- Supports ordinary unencrypted Transitional OOXML `.pptx` and `.pptm` packages. Does not execute macros, formulas, links, or embedded objects. It does not modify or sanitize the source presentation.
- Supports classic native chart series, including categorical, XY/scatter, bubble, multiple series, and series from combination charts. Point indices are preserved; missing values stay null. Saved caches may be stale. The extractor does not reconcile caches with Excel, evaluate formulas, or rebuild a chart from an uncached reference.
- Reads embedded OOXML workbooks regardless of filename extension. Legacy OLE `.bin`/`.xls`, `.xlsb`, non-Excel packages, encrypted decks, Strict OOXML, and chartEx plots need additional decoders. Embedded workbooks attached to chartEx are still attempted.
- Does not extract data from images, handwritten content, video, master/layout text, comments, or every possible proprietary object. SmartArt text is partial extraction. Shape-drawn tables/charts remain slide text. This is not an exhaustive document-forensics or sanitization tool.
- `review_record_count` includes explicit uncertainty flags as well as partial, unsupported, and reference-only records; image references include ordinary logos as well as screenshots. Read all warnings, including warnings on successfully extracted chart/workbook records.
- Hidden slides, hidden worksheets, and unreferenced embedded packages are included by default. Options can exclude them. Excluding hidden slides alone does not exclude their unreferenced embedded workbooks; set `include_unreferenced_embeddings=False` too if that is desired. Shared workbooks may still contain hidden-slide data.
- Default bounds: 256 MiB input, 128 MiB per ZIP member, 512 MiB expanded package, 20,000 archive entries, 250,000 cells per worksheet rectangle, and 250,000 points/rows per chart. A limit failure is reported rather than silently truncating a dataset. These bounds reduce accidental excessive work but are not a process-isolation guarantee for hostile files.

## Validation

The local regression suite generates real PPTX fixtures and checks values, worksheet caches, formulas, hidden content, missing chart points, scatter/bubble series, multi-level categories, presentation order, external-link handling, shared-workbook provenance, size limits, JSON round trips, and callback isolation. Additional cases validate side-by-side charts, separate headings/data, grouped charts, rotated and malformed geometry, duplicate labels, and uncertainty flags.

```bash
pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
```

Databricks workspace permissions, cluster installation, and live Spark/Delta operations require validation in your own workspace.

## Sample files

The included presentations contain synthetic data. Their author metadata, document timestamps, printer settings, and stored thumbnails have been removed, including metadata inside embedded workbooks and notebook demo copies. Notebook outputs and execution metadata are cleared. Sample JSON and JSONL reports use a fixed example timestamp.

## Implementation references

- [Databricks: Unity Catalog Volumes and Python paths](https://docs.databricks.com/aws/en/volumes/)
- [Databricks: Work with files in Volumes](https://docs.databricks.com/aws/en/volumes/volume-files)
- [Microsoft: Open XML chart numeric cache](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.charts.numberreference.numberingcache?view=openxml-3.0.1)
- [openpyxl: load_workbook and saved formula values](https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/reader/excel.html)
