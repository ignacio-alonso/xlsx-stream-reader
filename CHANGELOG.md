# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Available change types per release: Added|Changed|Deprecated|Fixed|Removed|Security

## Unreleased

### Added

## 1.1.1 - 2024-xx-xx

### Added
- Changelog file to keep track of changes through versions

### Changed

- Typos in Readme were corrected

### Fixed

- Bug when counting the number of rows.  
  Bug Description: when the cells in a row have no values, but at least one cell is formatted, the xml contains an entry for the row and the row counter is incremented.  
  The solution was to check if the row contains cells with value before incrementing the row counter
  
  - When a row has data the xml and the parsed node looks like this:
  ```
  [
    {name: 'row', attributes: {…}},
    {name: 'c', attributes: {…}},
    {name: 'v'},
    '0'
  ]

  <row r="5" customFormat="false" ht="12.8" hidden="false" customHeight="false" …>
    <c r="A5" s="1" t="s">
      <v>0</v>
    </c>
  </row>
  ```

  - Without data:
  ```
  [{
  name: "row",
  attributes: {…},
  }]

  <row r="4" customFormat="false" ht="12.8" hidden="false" customHeight="false" …>
    <c r="A4" s="1" />
  </row>
  ```


