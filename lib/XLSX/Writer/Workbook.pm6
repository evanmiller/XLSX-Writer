use v6;

use NativeCall;

use XLSX::Writer::Format;
use XLSX::Writer::Common;
use XLSX::Writer::Worksheet;
need XLSX::Writer::DocProperties;

unit class XLSX::Writer::Workbook is repr('CPointer') is export;

sub workbook_new(Str is encoded('utf8')) returns XLSX::Writer::Workbook is native(LIB) {...}
sub workbook_close(XLSX::Writer::Workbook) returns int32 is native(LIB) {...}
sub workbook_add_worksheet(XLSX::Writer::Workbook, Str is encoded('utf8')) returns XLSX::Writer::Worksheet is native(LIB) {...}
sub workbook_add_format(XLSX::Writer::Workbook) returns Format is native(LIB) {...}
sub workbook_set_properties(XLSX::Writer::Workbook, XLSX::Writer::DocProperties) returns int32 is native(LIB) {...}
sub workbook_set_custom_property_string(XLSX::Writer::Workbook, Str, Str) returns int32 is native(LIB) {...}
sub workbook_set_custom_property_number(XLSX::Writer::Workbook, Str, num64) returns int32 is native(LIB) {...}
sub workbook_set_custom_property_boolean(XLSX::Writer::Workbook, Str, uint8) returns int32 is native(LIB) {...}
sub workbook_set_custom_property_datetime(XLSX::Writer::Workbook, Str, XLSX::Writer::DateTime) returns int32 is native(LIB) {...}
sub workbook_define_name(XLSX::Writer::Workbook, Str, Str) returns int32 is native(LIB) {...}
sub workbook_get_worksheet_by_name(XLSX::Writer::Workbook, Str) returns XLSX::Writer::Worksheet is native(LIB) {...}
sub workbook_validate_worksheet_name(XLSX::Writer::Workbook, Str) returns int32 is native(LIB) {...}

method new(Str $path) {
    workbook_new($path)
}

method close() returns Error {
    Error(workbook_close(self))
}

method add-worksheet(Str $sheetname?) returns XLSX::Writer::Worksheet {
    workbook_add_worksheet(self, $sheetname)
}

method add-format() returns Format {
    workbook_add_format(self);
}

method define-name(Str:D $name, Str:D $formula) returns Error {
    Error(workbook_define_name(self, $name, $formula))
}

method get-worksheet-by-name(Str:D $name) returns XLSX::Writer::Worksheet {
    workbook_get_worksheet_by_name(self, $name)
}

method validate-worksheet-name(Str:D $name) returns Error {
    Error(workbook_validate_worksheet_name(self, $name))
}

method set-properties(XLSX::Writer::DocProperties $properties) returns Error {
    Error(workbook_set_properties(self, $properties))
}

multi method set-custom-property(Str:D $name, Str:D $value) returns Error {
    Error(workbook_set_custom_property_string(self, $name, $value))
}

multi method set-custom-property(Str:D $name, Numeric:D $value) returns Error {
    Error(workbook_set_custom_property_number(self, $name, $value.Num))
}

multi method set-custom-property(Str:D $name, Bool:D $value) returns Error {
    Error(workbook_set_custom_property_boolean(self, $name, $value))
}

multi method set-custom-property(Str:D $name, Dateish:D $value) returns Error {
    my $dt = XLSX::Writer::DateTime.new(:year($value.year), :month($value.month), :day($value.day),
            :hour($value.hour), :min($value.minute), :sec($value.second));
    Error(workbook_set_custom_property_datetime(self, $name, $dt))
}
