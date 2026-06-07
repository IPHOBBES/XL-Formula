# XL Formula Generator

A small, dependency-free web app for building reliable, ready-to-paste Excel
formulas without memorizing the syntax.

[Open XL Formula Generator](https://iphobbes.github.io/XL-Formula/)

## Features

- XLOOKUP, including composite keys and multiple fallback sources
- IFERROR + XLOOKUP
- FILTER for ranges and tables
- VSTACK + FILTER
- IF
- Sheet and structured-table reference modes
- Automatic quoting for sheet names and text values
- Automatic table qualification for FILTER rules
- Responsive, keyboard-accessible interface
- One-click formula copying

## How to use

1. Choose **Sheet** or **Table** reference mode.
2. Select the formula you want to build.
3. Complete the generated fields.
4. Copy the finished formula into Excel.

## Local development

Open `index.html` directly, or serve the folder with any static file server.
No build step or package installation is required.

Run the formula-core tests with:

```sh
node --test tests/formula-core.test.js
```

## License

[MIT](LICENSE)
