use v6;

unit class XLSX::Writer::Protection is repr('CStruct') is export;

has uint8 $.no_select_locked_cells;
has uint8 $.no_select_unlocked_cells;
has uint8 $.format_cells;
has uint8 $.format_columns;
has uint8 $.format_rows;
has uint8 $.insert_columns;
has uint8 $.insert_rows;
has uint8 $.insert_hyperlinks;
has uint8 $.delete_columns;
has uint8 $.delete_rows;
has uint8 $.sort;
has uint8 $.autofilter;
has uint8 $.pivot_tables;
has uint8 $.scenarios;
has uint8 $.objects;

has uint8 $.no_sheet;
has uint8 $.content;
has uint8 $.is_configured;

has int8 $.hash1;
has int8 $.hash2;
has int8 $.hash3;
has int8 $.hash4;
has int8 $.hash5;
