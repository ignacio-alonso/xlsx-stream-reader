/* global describe, it */

const XlsxStreamReader = require('../index')
const fs = require('fs')
const assert = require('assert')
const path = require('path')

describe('The xslx stream parser', function () {
  it('parses large files', function (done) {
    var workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'big.xlsx')).pipe(workBookReader)
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert(workSheetReader.rowCount === 80000)
        done()
      })
      workSheetReader.process()
    })
  })
  it('supports predefined formats', function (done) {
    var workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'predefined_formats.xlsx')).pipe(workBookReader)
    const rows = []
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert(rows[1][4] === '9/27/86')
        assert(rows[1][8] === '20064')
        done()
      })
      workSheetReader.on('row', function (r) {
        rows.push(r.values)
      })
      workSheetReader.process()
    })
  })
  it('supports custom formats', function (done) {
    var workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'import.xlsx')).pipe(workBookReader)
    const rows = []
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert(rows[1][2] === '27/09/1986')
        assert(rows[1][3] === '20064')
        done()
      })
      workSheetReader.on('row', function (r) {
        rows.push(r.values)
      })
      workSheetReader.process()
    })
  })
  it('supports date formate 1904', function (done) {
    var workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'date1904.xlsx')).pipe(workBookReader)
    const rows = []
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert(rows[1][2] === '27/09/1986')
        done()
      })
      workSheetReader.on('row', function (r) {
        rows.push(r.values)
      })
      workSheetReader.process()
    })
  })
  it('catches zip format errors', function (done) {
    var workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'notanxlsx')).pipe(workBookReader)
    workBookReader.on('error', function (err) {
      assert(err.message === 'invalid signature: 0x6d612069')
      done()
    })
  })
  it('parses a file with no number format ids', function (done) {
    const workBookReader = new XlsxStreamReader()
    const rows = []
    fs.createReadStream(path.join(__dirname, 'nonumfmt.xlsx')).pipe(workBookReader)
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert(rows[1][1] === 'lambrate')
        done()
      })
      workSheetReader.on('row', function (r) {
        rows.push(r.values)
      })
      workSheetReader.process()
    })
  })
  it('parses two files in parallel', done => {
    const file1 = 'import.xlsx'
    const file2 = 'file_with_2_sheets.xlsx'
    let finishedStreamCount = 0
    const endStream = function () {
      finishedStreamCount++

      if (finishedStreamCount === 2) {
        done()
      }
    }

    fs.createReadStream(path.join(__dirname, file1)).pipe(consumeXlsxFile(endStream))
    fs.createReadStream(path.join(__dirname, file2)).pipe(consumeXlsxFile(endStream))
  })
  it('support rich-text', function (done) {
    const workBookReader = new XlsxStreamReader({ saxTrim: false })
    fs.createReadStream(path.join(__dirname, 'richtext.xlsx')).pipe(workBookReader)
    const rows = []
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert(rows[0][2] === 'B cell')
        assert(rows[0][3] === 'C cell')
        done()
      })
      workSheetReader.on('row', function (r) {
        rows.push(r.values)
      })
      workSheetReader.process()
    })
  })
  it('parses a file having uppercase in sheet name and mixed first node', function (done) {
    const workBookReader = new XlsxStreamReader()
    const rows = []
    fs.createReadStream(path.join(__dirname, 'uppercase_sheet_name.xlsx')).pipe(workBookReader)
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert.strictEqual(rows.length, 24)
        assert.deepStrictEqual(
          rows[0].slice(1),
          ['Category ID', 'Parent category ID', 'Name DE', 'Name FR', 'Name IT', 'Name EN', 'GS1 ID']
        )
        done()
      })
      workSheetReader.on('row', function (r) {
        rows.push(r.values)
      })
      workSheetReader.process()
    })
    workBookReader.on('end', function () {
      rows.length || done(new Error('Read nothing'))
    })
  })
  it('parse 0 as 0', function (done) {
    const workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'issue_44_empty_0.xlsx')).pipe(workBookReader)
    const rows = []
    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('end', function () {
        assert.strictEqual(rows[1][1], 0)
        assert.strictEqual(rows[1][2], 1)
        done()
      })
      workSheetReader.on('row', function (r) {
        rows.push(r.values)
      })
      workSheetReader.process()
    })
  })
})

describe('Empty rows are omitted', function () {
  it('rows with no data are omitted', (done) => {
    const workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'empty_rows.xlsx')).pipe(
      workBookReader
    )

    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('row', function (row) {
        assert.strictEqual(Number(row.attributes.r), 2)
        done()
      })
      workSheetReader.process()
    })
  })
})

describe('Rows are properly counted', function () {
  it('when there are not rows with data the row count is 0', (done) => {
    const workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'row_counter.xlsx')).pipe(
      workBookReader
    )

    workBookReader.on('worksheet', function (workSheetReader) {
      if (workSheetReader.name !== 'no data') {
        workSheetReader.skip()
        return
      }

      workSheetReader.on('end', function () {
        assert.strictEqual(workSheetReader.rowCount, 0)
        done()
      })

      workSheetReader.process()
    })
  })

  it('when there are rows with formatted cells but no data, only the rows with data are counted', (done) => {
    const workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'row_counter.xlsx')).pipe(
      workBookReader
    )

    workBookReader.on('worksheet', function (workSheetReader) {
      if (workSheetReader.name !== 'formatted cells') {
        workSheetReader.skip()
        return
      }

      workSheetReader.on('end', function () {
        assert.strictEqual(workSheetReader.rowCount, 2)
        done()
      })

      workSheetReader.process()
    })
  })
})

describe('Row type node', function () {
  it('contains "number" attribute', (done) => {
    const workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'row_numbers.xlsx')).pipe(
      workBookReader
    )

    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('row', function (row) {
        assert.notStrictEqual(row.number, undefined)
        done()
      })
      workSheetReader.process()
    })
  })

  it('typeof row.number is "number"', (done) => {
    const workBookReader = new XlsxStreamReader()
    fs.createReadStream(path.join(__dirname, 'row_numbers.xlsx')).pipe(
      workBookReader
    )

    workBookReader.on('worksheet', function (workSheetReader) {
      workSheetReader.on('row', function (row) {
        assert.strictEqual(typeof row.number, 'number')
        done()
      })
      workSheetReader.process()
    })
  })
})

function consumeXlsxFile (cb) {
  const workBookReader = new XlsxStreamReader()
  workBookReader.on('worksheet', sheet => sheet.process())
  workBookReader.on('end', cb)
  return workBookReader
}
