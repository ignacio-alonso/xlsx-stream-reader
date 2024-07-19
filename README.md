# xlsx-stream-reader

[![JavaScript Style Guide](https://img.shields.io/badge/code_style-standard-brightgreen.svg)](https://standardjs.com)

======

Memory efficient minimalist streaming XLSX reader that can handle piped 
streams as input. Events are emitted while reading the stream.

Example

More examples can be found if `example` folder

```javascript
    var stream = new XlsxStreamReader({
        verbose: false,
        formatting: false
    })
```

Options

|Key|Default Value|Description|
|---|---|---|
|verbose|true|throw additional exceptions, if `false` - then pass empty string in that places|
|formatting|true|should cells with combined formats be formatted or not|
|saxTrim|true|whether or not to trim text and comment nodes|

-------
```javascript
const XlsxStreamReader = require("xlsx-stream-reader");

var workBookReader = new XlsxStreamReader();
workBookReader.on('error', function (error) {
    throw(error);
});
workBookReader.on('sharedStrings', function () {
    // do not need to do anything with these, 
    // cached and used when processing worksheets
    console.log(workBookReader.workBookSharedStrings);
});

workBookReader.on('styles', function () {
    // do not need to do anything with these
    // but not currently handled in any other way
    console.log(workBookReader.workBookStyles);
});

workBookReader.on('worksheet', function (workSheetReader) {
    if (workSheetReader.id > 1){
        // we only want first sheet
        workSheetReader.skip();
        return; 
    }
    // print worksheet name
    console.log(workSheetReader.name);

    // if we do not listen for rows we will only get end event
    // and have info about the sheet like row count
    workSheetReader.on('row', function (row) {
        if (row.attributes.r == 1){
            // do something with row 1 like save as column names
        }else{
            // second param to forEach colNum is very important as
            // null columns are not defined in the array, ie sparse array
            row.values.forEach(function(rowVal, colNum){
                // do something with row values
            });
        }
    });
    workSheetReader.on('end', function () {
        console.log(workSheetReader.rowCount);
    });

    // call process after registering handlers
    workSheetReader.process();
});
workBookReader.on('end', function () {
    // end of workbook reached
});

fs.createReadStream(fileName).pipe(workBookReader);

```
Beta Warning

-------
This module is currently in use on a live internal business system for product 
management. That being said this should still be considered beta. More usage 
and input from users will be needed due to the numerous differences/incompatibilities/flukes 
I have already run into with XLSX files.

Limitations

-------
The row reader currently returns stored values for formulas (these are normally available)
and does not calculate the formula itself. As time permits the row handler will be more capable 
but was enough for current purposes (loading values from large worksheets fast)
 
Inspiration

-----------
The project originated as a fork of [xlsx-stream-reader], to address some issues encountered while using it. 
The package was no longer maintained and forks were not being merged anymore, so we decided to create a new package instead.

Need a simple XLSX file streaming reader to handle large excel sheets but only
one available/compatible was by guyonroche/exceljs. The stream reader module at
the time was unfinished/unusable and rewrite attempts exposed column shifting I
could not solve

More Information

-----------
Events are emitted as pertinent parts of the workbook and worksheet are received
in the stream. Theoretically you could pause the input stream if events are being
received too fast but this has not been tested

Events can potentially (even though I have not seen it) be received out of order,
if you receive a worksheet end event while still receiving rows be sure to make sure
your number of rows received equals the `workSheetReader.rowCount` 

Theoretically you could process an excel sheet as it is being uploaded, depending
on the sheet type, but untried (I encountered some XLSX files that have a different
zip format that requires having the entire file to read the archive contents properly),
but still probably better to save temp first and read streams from there.

Currently if the zip archive does not have the shared strings at the beginning of the
archive then the input stream for each sheet is pied into a temp file until the shared
string are encountered and processed, then re-read the temp worksheets with the shared
strings.

API Information

-----------
#### new XlsxStreamReader()

Create a new XlsxStreamReader object (workBookReader). After attaching handlers you
can `pipe()` your input stream into the reader to begin processing

#### Event: 'error'

* `error` {Error Object}

Emitted if there was an error in processing (may not catch all errors, 
some may be thrown depending on where the error happened)

#### Event: 'end'

Emitted once the XLSX zip parser has closed and all sheets have been processed

#### Event: 'sharedStrings'

After the workbook shared strings have been parsed this event is emitted. Shared strings 
are available via array `workBookReader.workBookSharedStrings`.

#### Event: 'styles'

After the workbook styles have been parsed this event is emitted. Styles are available
via array `workBookReader.workBookStyles`

#### Event: 'worksheet'

* `workSheetReader` {Object} XlsxStreamReaderWorkSheet object

Emitted when a worksheet is reached. The sheet number is available via 
{Number} `workSheetReader.id`. You can either process or skip at this point, 
but you must do one for the processing to the next sheet to continue/finish.

Once event is received you can attach worksheet on handlers (end, row) then you
would `workSheetReader.process()`. If you do not want to process a sheet and instead
want to skip entirely, you would `workSheetReader.skip()` without attaching any handlers.

#### Worksheet Event: 'end'

Emitted once the end of the worksheet has been reached. The row count is 
available via {Number} `workSheetReader.rowCount`

#### Worksheet Event: 'row'

* `row` {Object} Row object

Emitted on every row encountered in the worksheet. for more details on what 
is in the row object attributes, see the [Row class][msdnRows] on MSDN.  

For example:

* `row.values`: sparse array containing all cell values
* `row.formulas`: sparse array containing all cell formulas
* `row.attributes.r`: row index
* `row.attributes.ht`: Row height measured in point size
* `row.attributes.customFormat`: '1' if the row style should be applied.
* `row.attributes.hidden`: '1' if the row is hidden

References

-----------
* [Working with sheets (Open XML SDK)][msdnSheets]
* [Row class][msdnRows]
* [ExcelJS][ExcelJS]

Used Modules

-----------
* [Path][modPath]
* [Util][modUtil]
* [Stream][modStream]
* [Sax][modSax]
* [unzipper][modUnzipper]
* [Temp][modTemp]

Authors

-----------
Written by [Ignacio Alonso](https://github.com/ignacio-alonso) and [Fernando Quintan](https://github.com/fernandoqcv). Fork of [xlsx-stream-reader]

License

-----------
[MIT](LICENSE)

[xlsx-stream-reader]: https://www.npmjs.com/package/xlsx-stream-reader
[gratipay-image-daspawn]: https://img.shields.io/gratipay/team/daspawn.svg
[msdnRows]: https://msdn.microsoft.com/EN-US/library/office/documentformat.openxml.spreadsheet.row.aspx
[msdnSheets]: https://msdn.microsoft.com/EN-US/library/office/gg278309.aspx
[ExcelJS]: https://github.com/guyonroche/exceljs

[modPath]: https://nodejs.org/api/path.html
[modStream]: https://nodejs.org/api/stream.html
[modUtil]: https://nodejs.org/api/util.html
[modSax]: https://github.com/isaacs/sax-js
[modUnzipper]: https://github.com/ZJONSSON/node-unzipper
[modTemp]: https://github.com/bruce/node-temp