use v6;

use NativeCall;
use XLSX::Writer::Format;
use XLSX::Writer::Common;
use XLSX::Writer::Protection;
need XLSX::Writer::DateTime;
need XLSX::Writer::RowColOptions;
need XLSX::Writer::HeaderFooterOptions;

unit class XLSX::Writer::Worksheet is repr('CPointer') is export;

sub worksheet_write_number(XLSX::Writer::Worksheet, uint32, uint16, num64, Format) returns int32 is native(LIB) {*}
sub worksheet_write_string(XLSX::Writer::Worksheet, uint32, uint16, Str, Format) returns int32 is native(LIB) {...}
sub worksheet_write_formula(XLSX::Writer::Worksheet, uint32, uint16, Str, Format) returns int32 is native(LIB) {...}
sub worksheet_write_array_formula(XLSX::Writer::Worksheet, uint32, uint16, uint32, uint16, Str, Format) returns int32 is native(LIB) {...}
sub worksheet_write_datetime(XLSX::Writer::Worksheet, uint32, uint16, XLSX::Writer::DateTime, Format) returns int32 is native(LIB) {...}
sub worksheet_write_url(XLSX::Writer::Worksheet, uint32, uint16, Str, Format) returns int32 is native(LIB) {...}
sub worksheet_write_boolean(XLSX::Writer::Worksheet, uint32, uint16, int32, Format) returns int32 is native(LIB) {...}
sub worksheet_write_blank(XLSX::Writer::Worksheet, uint32, uint16, Format) returns int32 is native(LIB) {...}
sub worksheet_set_row(XLSX::Writer::Worksheet, uint32, num64, Format) returns int32 is native(LIB) {...}
sub worksheet_set_row_opt(XLSX::Writer::Worksheet, uint32, num64, Format, XLSX::Writer::RowColOptions) returns int32 is native(LIB) {...}
sub worksheet_set_column(XLSX::Writer::Worksheet, uint16, uint16, num64, Format) returns int32 is native(LIB) {...}
sub worksheet_set_column_opt(XLSX::Writer::Worksheet, uint16, uint16, num64, Format, XLSX::Writer::RowColOptions) returns int32 is native(LIB) {...}
sub worksheet_insert_image(XLSX::Writer::Worksheet, uint32, uint16, Str, Format) returns int32 is native(LIB) {...}
sub worksheet_merge_range(XLSX::Writer::Worksheet, uint32, uint16, uint32, uint16, Str, Format) returns int32 is native(LIB) {...}
sub worksheet_autofilter(XLSX::Writer::Worksheet, uint32, uint16, uint32, uint16) returns int32 is native(LIB) {...}
sub worksheet_activate(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_select(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_hide(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_set_first_sheet(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_freeze_panes(XLSX::Writer::Worksheet, uint32, uint16) is native(LIB) {...}
sub worksheet_split_panes(XLSX::Writer::Worksheet, num64, num64) is native(LIB) {...}
sub worksheet_set_selection(XLSX::Writer::Worksheet, uint32, uint16, uint32, uint16) is native(LIB) {...}
sub worksheet_set_landscape(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_set_portrait(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_set_page_view(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_set_paper(XLSX::Writer::Worksheet, uint8) is native(LIB) {...}
sub worksheet_set_margins(XLSX::Writer::Worksheet, num64, num64, num64, num64) is native(LIB) {...}
sub worksheet_set_header(XLSX::Writer::Worksheet, Str) returns int32 is native(LIB) {...}
sub worksheet_set_header_opt(XLSX::Writer::Worksheet, Str, XLSX::Writer::HeaderFooterOptions) returns int32 is native(LIB) {...}
sub worksheet_set_footer(XLSX::Writer::Worksheet, Str) returns int32 is native(LIB) {...}
sub worksheet_set_footer_opt(XLSX::Writer::Worksheet, Str, XLSX::Writer::HeaderFooterOptions) returns int32 is native(LIB) {...}
sub worksheet_set_h_pagebreaks(XLSX::Writer::Worksheet, CArray[uint32]) returns int32 is native(LIB) {...}
sub worksheet_set_v_pagebreaks(XLSX::Writer::Worksheet, CArray[uint16]) returns int32 is native(LIB) {...}
sub worksheet_print_across(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_set_zoom(XLSX::Writer::Worksheet, uint16) is native(LIB) {...}
sub worksheet_gridlines(XLSX::Writer::Worksheet, uint8) is native(LIB) {...}
sub worksheet_center_horizontally(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_center_vertically(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_print_row_col_headers(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_repeat_rows(XLSX::Writer::Worksheet, uint32, uint32) returns int32 is native(LIB) {...}
sub worksheet_repeat_columns(XLSX::Writer::Worksheet, uint16, uint16) returns int32 is native(LIB) {...}
sub worksheet_print_area(XLSX::Writer::Worksheet, uint32, uint16, uint32, uint16) returns int32 is native(LIB) {...}
sub worksheet_fit_to_pages(XLSX::Writer::Worksheet, uint16, uint16) is native(LIB) {...}
sub worksheet_set_start_page(XLSX::Writer::Worksheet, uint16) is native(LIB) {...}
sub worksheet_set_print_scale(XLSX::Writer::Worksheet, uint16) is native(LIB) {...}
sub worksheet_right_to_left(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_hide_zero(XLSX::Writer::Worksheet) is native(LIB) {...}
sub worksheet_set_tab_color(XLSX::Writer::Worksheet, int32) is native(LIB) {...}
sub worksheet_protect(XLSX::Writer::Worksheet, Str is encoded('utf8'), Protection) is native(LIB) {...}
sub worksheet_set_default_row(XLSX::Writer::Worksheet, num64, uint8) is native(LIB) {...}

multi method write-value(UInt:D $row, UInt:D $col, Numeric $number, Format $format?) returns Error {
    Error(worksheet_write_number(self, $row, $col, $number.Num, $format))
}

multi method write-value(UInt:D $row, UInt:D $col, Str $string, Format $format?) returns Error {
    Error(worksheet_write_string(self, $row, $col, $string, $format))
}

multi method write-value(UInt:D $row, UInt:D $col, Bool $bool, Format $format?) returns Error {
    Error(worksheet_write_boolean(self, $row, $col, $bool, $format))
}

method write-formula(UInt:D $row, UInt:D $col, Str $formula, Format $format?) returns Error {
    Error(worksheet_write_formula(self, $row, $col, $formula, $format))
}

# Nice error message...
# Undeclared routine:
#   write-array-formula used at lines 75, 79, 83, 87

multi method write-array-formula(UInt:D $row, UInt:D $col, Str $formula, Format $format?) returns Error {
    self.write-array-formula($row..$row, $col..$col, $formula, $format)
}

multi method write-array-formula(UInt:D $row, ColRange:D $cols, Str $formula, Format $format?) returns Error {
    self.write-array-formula($row..$row, $cols, $formula, $format)
}

multi method write-array-formula(RowRange:D $rows, UInt:D $col, Str $formula, Format $format?) returns Error {
    self.write-array-formula($rows, $col..$col, $formula, $format)
}

multi method write-array-formula(UInt:D $row, ColRange:D $cols, Str $formula, Format $format?) returns Error {
    self.write-array-formula($row..$row, $cols, $formula, $format)
}

multi method write-array-formula(RowRange:D $rows, ColRange:D $cols, Str $formula, Format $format?) returns Error {
    Error(worksheet_write_array_formula(self, $rows.min, $cols.min, $rows.max, $cols.max, $formula, $format))
}

method write-datetime(UInt:D $row, UInt:D $col, Dateish $date, Format $format?) returns Error {
    my $dt = XLSX::Writer::DateTime.new(:year($date.year), :month($date.month), :day($date.day),
            :hour($date.hour), :min($date.minute), :sec($date.second));
    Error(worksheet_write_datetime(self, $row, $col, $dt, $format))
}

method write-url(UInt:D $row, UInt:D $col, Str $url, Format $format?) returns Error {
    Error(worksheet_write_url(self, $row, $col, $url, $format))
}

method write-blank(UInt:D $row, UInt:D $col, Format $format?) returns Error {
    Error(worksheet_write_blank(self, $row, $col, $format))
}

method set-row(UInt:D $row, Numeric $height, Format $format?, XLSX::Writer::RowColOptions $options?) returns Error {
    Error(worksheet_set_row_opt(self, $row, $height.Num, $format, $options))
}

multi method set-column(UInt:D $col, Numeric $width, Format $format?, XLSX::Writer::RowColOptions $options?) returns Error {
    Error(worksheet_set_column_opt(self, $col, $col, $width.Num, $format, $options))
}

multi method set-column(ColRange:D $cols, Numeric $width, Format $format?, XLSX::Writer::RowColOptions $options?) returns Error {
    Error(worksheet_set_column_opt(self, $cols.min, $cols.max, $width.Num, $format, $options))
}

method insert-image(UInt:D $row, UInt:D $col, Str $filename) returns Error {
    Error(worksheet_insert_image(self, $row, $col, $filename))
}

method merge-range(RowRange:D $rows, ColRange:D $cols, Str $string, Format $format?) returns Error {
    Error(worksheet_merge_range(self, $rows.min, $cols.min, $rows.max, $cols.max, $string, $format))
}

method autofilter(RowRange:D $rows, ColRange:D $cols) {
    Error(worksheet_autofilter(self, $rows.min, $cols.min, $rows.max, $cols.max))
}

method activate() {
    worksheet_activate(self)
}

method select() {
    worksheet_select(self)
}

method hide() {
    worksheet_hide(self)
}

method set-first-sheet() {
    worksheet_set_first_sheet(self)
}

method freeze-panes(UInt:D $row, UInt:D $col) {
    worksheet_freeze_panes(self, $row, $col)
}

method split-panes(Numeric $vertical, Numeric $horizontal) {
    worksheet_split_panes(self, $vertical.Num, $horizontal.Num)
}

method set-selection(RowRange:D $rows, ColRange:D $cols) {
    worksheet_set_selection(self, $rows.min, $cols.min, $rows.max, $cols.max)
}

method set-landscape() {
    worksheet_set_landscape(self)
}

method set-portrait() {
    worksheet_set_portrait(self)
}

method set-page-view() {
    worksheet_set_page_view(self)
}

method set-paper(UInt $type where { $type < 42 }) { # gah vim-perl6 is annoying with <=
    worksheet_set_paper(self, $type)
}

method set-margins(Numeric $left, Numeric $right, Numeric $top, Numeric $bottom) {
    worksheet_set_margins(self, $left.Num, $right.Num, $top.Num, $bottom.Num)
}

method set-header(Str $caption, XLSX::Writer::HeaderFooterOptions $options?) returns Error {
    Error(worksheet_set_header_opt(self, $caption, $options))
}

method set-footer(Str $caption, XLSX::Writer::HeaderFooterOptions $options?) returns Error {
    Error(worksheet_set_footer_opt(self, $caption, $options))
}

method set-h-pagebreaks(@breaks) returns Error {
    Error(worksheet_set_h_pagebreaks(self, CArray[int32].new(|@breaks, 0)))
}

method set-v-pagebreaks(@breaks) returns Error {
    Error(worksheet_set_v_pagebreaks(self, CArray[int32].new(|@breaks, 0)))
}

method print-across() {
    worksheet_print_across(self)
}

method set-zoom(UInt:D $zoom where 400 >= * >= 10) {
    worksheet_set_zoom(self, $zoom)
}

method gridlines(Gridlines $gridlines) {
    worksheet_gridlines(self, $gridlines)
}

method center-horizontally() {
    worksheet_center_horizontally(self)
}

method center-vertically() {
    worksheet_center_vertically(self)
}

method print-row-col-headers() {
    worksheet_print_row_col_headers(self)
}

method repeat-rows(RowRange:D $rows) returns Error {
    Error(worksheet_repeat_rows(self, $rows.min, $rows.max))
}

method repeat-columns(ColRange:D $cols) returns Error {
    Error(worksheet_repeat_columns(self, $cols.min, $cols.max))
}

method print-area(RowRange:D $rows, ColRange:D $cols) returns Error {
    Error(worksheet_print_area(self, $rows.min, $cols.min, $rows.max, $cols.max))
}

method fit-to-pages(UInt:D $width, UInt:D $height) {
    worksheet_fit_to_pages(self, $width, $height)
}

method set-start-page(UInt:D $page) {
    worksheet_set_start_page(self, $page)
}

method set-print-scale(UInt:D $scale where 400 >= * >= 10) {
    worksheet_set_print_scale(self, $scale)
}

method right-to-left() {
    worksheet_right_to_left(self)
}

method hide-zero() {
    worksheet_hide_zero(self)
}

method set-tab-color(Color $color) {
    worksheet_set_tab_color(self, $color)
}

method protect(Str :$password, Protection :$protection) {
    worksheet_protect(self, $password, $protection)
}

method set-default-row(Numeric $height, Bool $hide_unused_rows) {
    worksheet_set_default_row(self, $height.Num, $hide_unused_rows)
}
