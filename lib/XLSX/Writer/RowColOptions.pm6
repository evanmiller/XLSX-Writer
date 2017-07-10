use v6;

unit class XLSX::Writer::RowColOptions is repr('CStruct') is export;
has uint8 $.hidden;
has uint8 $.level;
has uint8 $.collapsed;
