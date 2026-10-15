# macroost

A VBA-compatible macro engine written in Rust. macroost reads the macro project embedded in
macro-enabled documents, parses and runs the code, and gives it an application object model to
work on: a spreadsheet, a word-processor document or a presentation, each backed by a storage
trait that any application can implement.

It is a clean-room implementation, built from published file-format and language specifications
and from black-box conformance testing. It contains no third-party code, assets or
documentation.

## Workspace

| Crate | What it does |
|---|---|
| `macroost-project` (`crates/project`) | Reads and writes macro projects (compressed module streams, the `dir` stream, project metadata) from `.xlsm`, `.docm`, `.pptm` packages and legacy compound files. Writes source-only projects. |
| `macroost-lang` (`crates/lang`) | The language: lexer, parser, interpreter, operators and standard library. Variants and coercion rules, `ByRef`/`ByVal`, arrays and `ReDim Preserve`, `With`, `Select Case`, every loop form, `GoSub`, `On Error Resume Next / GoTo / Resume`, exact 96-bit `Decimal` arithmetic, classes with properties, user-defined types, events, `Collection` and a dictionary object. |
| `macroost-sheet` (`crates/sheet`) | A spreadsheet object model (`Application`, `Workbook`, `Worksheet`, `Range`, `Names`, `WorksheetFunction`, protection, sort, filter, fill, find) over a `Workbook` trait, with an in-memory implementation and a formula evaluator. |
| `macroost-xlsx` (`crates/xlsx`) | Loads and saves OOXML spreadsheets into the in-memory workbook, keeping everything it doesn't model intact on save. |
| `macroost-doc` (`crates/doc`) | A word-processor object model (`Document`, `Range`, `Selection`, `Find`, `Tables`, `Bookmarks`) over a `Document` trait. |
| `macroost-slides` (`crates/slides`) | A presentation object model (`Presentation`, `Slides`, `Shapes`, `TextRange`, tables) over a `Presentation` trait. |
| `macroost-cli` (`crates/cli`) | The `macroost` command: `list`, `extract`, `run` (optionally `--save`), `sweep`. |

The language engine knows nothing about any application: it talks to an object model only
through the `Object` and `Host` traits in `crates/lang/src/object.rs`. An application adds macro
support by implementing a storage trait (`Workbook`, `Document` or `Presentation`).

## Safety

Macros are untrusted code. macroost gives them no file system, shell, network, registry, native
calls (`Declare`) or arbitrary object creation; only safe built-in objects can be created. Every
run is bounded by a step budget and a recursion limit, and array sizes are capped, so a macro
can't hang or exhaust the application. Hosts can run a macro against a scratch copy and apply
the result as one undoable change.

## Command line

```
macroost list    <file>                         modules and procedures
macroost extract <file> [out-dir]               write each module's source
macroost run     <file|dir|.bas> <Macro> [args] [--save out.xlsm]
macroost sweep   <list-of-files>                run every argument-free procedure of each file
```

## Conformance

`conformance/cases/<suite>/*.bas` holds test macros; `conformance/expected/<suite>.json` holds
the results recorded from a reference implementation. The `conformance_*` tests run every case
end to end (parser, interpreter, object model) and compare. Each suite has a ratchet floor that
only goes up.

| Suite | Cases | Matching |
|---|---|---|
| Language: expressions (`lang`) | 194 | 194 |
| Language: programs (`lang2`) | 40 | 40 |
| Language: Decimal type (`decimal`) | 70 | 70 |
| Spreadsheet (`sheet1`–`sheet5`) | 464 | 464 |
| Word processor (`doc`) | 280 | 280 |
| Presentation (`slides`, `slides2`, `slides_window`) | 123 | 123 |

## Building

```
cargo test --workspace
scripts/check-all.sh   # fmt, clippy -D warnings, tests, and compile checks for Linux, macOS and wasm32
```

Pure Rust with no platform code: it builds for Windows, macOS, Linux and WebAssembly.

## Status and disclaimers

macroost is an independent project. It is not affiliated with, endorsed by, or sponsored by the
maker of any office suite or programming language. Product and language names are trademarks of
their respective owners and are used only to describe compatibility.

It is experimental software, provided as is, without warranty of any kind (see the licenses).
Don't rely on it for work where a wrong result matters without checking the result.

## Security

macroost is designed to run untrusted macros safely (see [Safety](#safety)). If you find a way
for a macro to escape the sandbox, please report it privately through the repository's security
advisory page rather than in a public issue.

## License

Licensed under either of

- Apache License, Version 2.0 ([LICENSE-APACHE](LICENSE-APACHE))
- MIT license ([LICENSE-MIT](LICENSE-MIT))

at your option.

Unless you explicitly state otherwise, any contribution intentionally submitted for inclusion in
the work by you, as defined in the Apache-2.0 license, shall be dual licensed as above, without
any additional terms or conditions.
