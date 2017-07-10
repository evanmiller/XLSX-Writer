use v6;
use Test;
use lib 'lib';
use XLSX::Writer;

my $tmpfile = "$*TMPDIR/XLSX-Writer-test-{time}.xlsx";

my $wbx = XLSX::Writer::Workbook.new($tmpfile);
ok $wbx;

my $fmt = $wbx.add-format();
ok $fmt;

$fmt.set-bold();
$fmt.set-font-color(XLSX::Writer::Color($_)) for XLSX::Writer::Color.enums.values;
$fmt.set-font-size(24);
$fmt.set-font-name("Comic Sans MS");
$fmt.set-italic();
$fmt.set-underline(XLSX::Writer::Underline($_)) for XLSX::Writer::Underline.enums.values;
$fmt.set-font-strikeout();
$fmt.set-font-script(XLSX::Writer::Script($_)) for XLSX::Writer::Script.enums.values;
$fmt.set-num-format("###.00");
$fmt.set-num-format-index(12);
$fmt.set-hidden();
$fmt.set-align(XLSX::Writer::Align($_)) for XLSX::Writer::Align.enums.values;
$fmt.set-text-wrap();
$fmt.set-rotation($_) for -90 .. 90;
$fmt.set-rotation(270);
$fmt.set-indent($_) for 1 .. 10;
$fmt.set-shrink();
$fmt.set-pattern(XLSX::Writer::Pattern($_)) for XLSX::Writer::Pattern.enums.values;
$fmt.set-bg-color(XLSX::Writer::Color($_)) for XLSX::Writer::Color.enums.values;
$fmt.set-fg-color(XLSX::Writer::Color($_)) for XLSX::Writer::Color.enums.values;
$fmt.set-border(XLSX::Writer::Border($_)) for XLSX::Writer::Border.enums.values;
$fmt.set-top(XLSX::Writer::Border($_)) for XLSX::Writer::Border.enums.values;
$fmt.set-left(XLSX::Writer::Border($_)) for XLSX::Writer::Border.enums.values;
$fmt.set-right(XLSX::Writer::Border($_)) for XLSX::Writer::Border.enums.values;
$fmt.set-border-color(XLSX::Writer::Color($_)) for XLSX::Writer::Color.enums.values;
$fmt.set-bottom-color(XLSX::Writer::Color($_)) for XLSX::Writer::Color.enums.values;
$fmt.set-top-color(XLSX::Writer::Color($_)) for XLSX::Writer::Color.enums.values;
$fmt.set-left-color(XLSX::Writer::Color($_)) for XLSX::Writer::Color.enums.values;
$fmt.set-right-color(XLSX::Writer::Color($_)) for XLSX::Writer::Color.enums.values;

is $wbx.close(), XLSX::Writer::Error::no-error;

done-testing;

unlink $tmpfile;
CATCH {
    unlink $tmpfile;
}
