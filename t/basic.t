use v6;
use Test;
use lib 'lib';
use XLSX::Writer;

my $tmpfile = "$*TMPDIR/XLSX-Writer-test-{time}.xlsx";

my $wbx = XLSX::Writer::Workbook.new($tmpfile);
ok $wbx;

is $wbx.validate-worksheet-name("Hello"), XLSX::Writer::Error::no-error;
is $wbx.validate-worksheet-name("Caesar: Then fall"), XLSX::Writer::Error::invalid-sheetname-character;

my $dp = XLSX::Writer::DocProperties.new(:title("Great Spreadsheet"));
ok $dp;
is $wbx.set-properties($dp), XLSX::Writer::Error::no-error;
is $wbx.set-custom-property("Author of all things", "God"), XLSX::Writer::Error::no-error;
is $wbx.set-custom-property("Books in the Bible", 66), XLSX::Writer::Error::no-error;
is $wbx.set-custom-property("Here for a reason", True), XLSX::Writer::Error::no-error;
is $wbx.set-custom-property("Party starts", DateTime.now), XLSX::Writer::Error::no-error;

my $wsx = $wbx.add-worksheet("Hello");
ok $wsx;

ok $wbx.get-worksheet-by-name("Hello");

my $pr = XLSX::Writer::Protection.new(:no_select_locked_cells(1));
ok $pr;
$wsx.protect(:password("Foobar"), :protection($pr));

my $fmt = $wbx.add-format();
ok $fmt;
$fmt.set-unlocked();

is $wsx.write-value(0, 0, "Locked"), XLSX::Writer::Error::no-error;
is $wsx.write-value(1, 0, "Unlocked", $fmt), XLSX::Writer::Error::no-error;

is $wsx.write-value(0, 1, 20.0), XLSX::Writer::Error::no-error;
is $wsx.write-value(1, 1, pi), XLSX::Writer::Error::no-error;

is $wsx.write-value(0, 2, False), XLSX::Writer::Error::no-error;
is $wsx.write-value(1, 2, True), XLSX::Writer::Error::no-error;

is $wsx.write-formula(0, 3, "=1+2"), XLSX::Writer::Error::no-error;

is $wsx.write-array-formula(0..1, 4, "=1+2"), XLSX::Writer::Error::no-error;
is $wsx.write-array-formula(0..1, 5..6, "=1+2"), XLSX::Writer::Error::no-error;
is $wsx.write-array-formula(0, 7..8, "=1+2"), XLSX::Writer::Error::no-error;
is $wsx.write-array-formula(0, 9, "=1+2"), XLSX::Writer::Error::no-error;
is $wsx.write-datetime(0, 10, DateTime.now), XLSX::Writer::Error::no-error;
is $wsx.write-url(0, 11, "http://perl6.org/"), XLSX::Writer::Error::no-error;
is $wsx.write-blank(0, 12), XLSX::Writer::Error::no-error;

is $wsx.set-row(0, 12), XLSX::Writer::Error::no-error;
is $wsx.set-column(0, 100), XLSX::Writer::Error::no-error;

is $wsx.merge-range(3..4, 3..4, ""), XLSX::Writer::Error::no-error;
is $wsx.autofilter(3..4, 3..4), XLSX::Writer::Error::no-error;

$wsx.activate();
$wsx.select();
$wsx.hide();
$wsx.set-first-sheet();
$wsx.freeze-panes(1, 1);
$wsx.split-panes(15, 8.43);
$wsx.set-selection(0..1, 0..1);
$wsx.set-landscape();
$wsx.set-portrait();
$wsx.set-page-view();
$wsx.set-paper($_) for ^42;
$wsx.set-margins(10.0, 10.0, 10.0, 10.0);

is $wsx.set-header("My Header"), XLSX::Writer::Error::no-error;
is $wsx.set-footer("My Footer"), XLSX::Writer::Error::no-error;

is $wsx.set-h-pagebreaks([ 20, 30 ]), XLSX::Writer::Error::no-error;
is $wsx.set-v-pagebreaks([ 40, 50 ]), XLSX::Writer::Error::no-error;

$wsx.print-across();
$wsx.set-zoom($_) for 10, 20 ... 400;
$wsx.gridlines(XLSX::Writer::Gridlines($_)) for XLSX::Writer::Gridlines.enums.values;
$wsx.center-horizontally();
$wsx.center-vertically();
$wsx.print-row-col-headers();

is $wsx.repeat-rows( 10 .. 20 ), XLSX::Writer::Error::no-error;
is $wsx.repeat-columns( 10 .. 20 ), XLSX::Writer::Error::no-error;
is $wsx.print-area( 10 .. 20, 10 .. 20 ), XLSX::Writer::Error::no-error;

$wsx.fit-to-pages(2, 2);
$wsx.set-start-page(2);
$wsx.set-print-scale($_) for 10, 20 ... 400;
$wsx.right-to-left();
$wsx.hide-zero();
$wsx.set-tab-color(XLSX::Writer::Color($_)) for XLSX::Writer::Color.enums.values;
$wsx.set-default-row(100.0, True);

is $wbx.close(), XLSX::Writer::Error::no-error;
ok $tmpfile.IO.f;
ok $tmpfile.IO.s > 0;

done-testing;

unlink $tmpfile;
CATCH {
    unlink $tmpfile;
}

