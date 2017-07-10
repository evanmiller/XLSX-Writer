use v6;

unit class XLSX::Writer::DateTime is repr('CStruct') is export;

has int32 $.year;
has int32 $.month;
has int32 $.day;
has int32 $.hour;
has int32 $.min;
# alignment ok
has num64 $.sec;
