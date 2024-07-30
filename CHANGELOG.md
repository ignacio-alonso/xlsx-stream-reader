# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Available change types per release: Added|Changed|Deprecated|Fixed|Removed|Security

## Unreleased

### Added

## 1.x.x - 2024-xx-xx

### Added
- Changelog file to keep track of changes through versions

### Fixed

- Bug when rows have no data.  
  Bug Description: when the cells in a row have no values, but at least one cell is formatted, the xml contains an entry for the row, and the `row.values` array is filled with empty strings.
  The solution was to only emit the row if at least one cell has a value.  

### Security
- Vulnerabilities  
  Vulnerabilities were addressed by updating dependencies. The only library that underwent a major version change was Mocha.

