use v6;

need XLSX::Writer::Align;
need XLSX::Writer::Border;
need XLSX::Writer::Pattern;

unit package XLSX::Writer;

constant LIB is export = "xlsxwriter";
constant MAX_ROWS = 1_048_576;
constant MAX_COLS =  16_384;

enum Color is export (
    black => 0x1000000,
    blue => 0x0000FF,
    brown => 0x800000,
    cyan => 0x00FFFF,
    gray => 0x808080,
    green => 0x008000,
    lime => 0x00FF00,
    magenta => 0xFF00FF,
    navy => 0x000080,
    orange => 0xFF6600,
    pink => 0xFF00FF,
    purple => 0x800080,
    red => 0xFF0000,
    silver => 0xC0C0C0,
    white => 0xFFFFFF,
    yellow => 0xFFFF00
);

enum Error is export <
    no-error                       memory-malloc-failed              creating-xlsx-file
    creating-tmpfile               reading-tmpfile                   zip-file-operation
    zip-parameter-error            zip-bad-file                      zip-internal-error
    zip-file-add                   zip-close                         feature-not-supported
    null-parameter-ignored         parameter-validation              sheetname-length-exceeded
    invalid-sheetname-character    sheetname-start-end-apostrophe    sheetname-already-used
    string-length-exceeded-32      string-length-exceeded-128        string-length-exceeded-255
    string-length-exceeded-max     shared-string-index-not-found     worksheet-index-out-of-range
    max-url-length-exceeded        max-number-urls-exceeded          image-dimensions
>;

enum Gridlines is export <hide-all-gridlines show-screen-gridlines show-print-gridlines show-all-gridlines>;

enum Underline is export <<:single(1) double single-accounting double-accounting>>;
enum Script is export <<:superscript(1) subscript>>;


subset RowRange of Range is export where { $_.is-int and $_.min >= 0 and $_.max < MAX_ROWS };
subset ColRange of Range is export where { $_.is-int and $_.min >= 0 and $_.max < MAX_COLS };
