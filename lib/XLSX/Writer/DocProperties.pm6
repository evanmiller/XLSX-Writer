use v6;

use NativeCall;

unit class XLSX::Writer::DocProperties is repr('CStruct') is export;

has Str $!title;
has Str $!subject;
has Str $!author;
has Str $!manager;
has Str $!company;
has Str $!category;
has Str $!keywords;
has Str $!comments;
has Str $!status;
has Str $!hyperlink_base;

has longlong $!created;
