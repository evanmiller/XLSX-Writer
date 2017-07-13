use v6;

use XLSX::Writer::Common;

use NativeCall;

unit class XLSX::Writer::Format is repr('CPointer') is export;

sub format_set_bold(XLSX::Writer::Format) is native(LIB) {*}
sub format_set_font_color(XLSX::Writer::Format, int32) is native(LIB) {*}
sub format_set_font_name(XLSX::Writer::Format, Str is encoded('utf8')) is native(LIB) {*}
sub format_set_font_size(XLSX::Writer::Format, uint16) is native(LIB) {*}
sub format_set_italic(XLSX::Writer::Format) is native(LIB) {*}
sub format_set_underline(XLSX::Writer::Format, int32) is native(LIB) {*}
sub format_set_font_strikeout(XLSX::Writer::Format) is native(LIB) {*}
sub format_set_font_script(XLSX::Writer::Format, uint8) is native(LIB) {*}
sub format_set_num_format(XLSX::Writer::Format, Str is encoded('utf8')) is native(LIB) {*}
sub format_set_num_format_index(XLSX::Writer::Format, uint8) is native(LIB) {*}
sub format_set_unlocked(XLSX::Writer::Format) is native(LIB) {*}
sub format_set_hidden(XLSX::Writer::Format) is native(LIB) {*}
sub format_set_align(XLSX::Writer::Format, uint8) is native(LIB) {*}
sub format_set_text_wrap(XLSX::Writer::Format) is native(LIB) {*}
sub format_set_rotation(XLSX::Writer::Format, int16) is native(LIB) {*}
sub format_set_indent(XLSX::Writer::Format, uint8) is native(LIB) {*}
sub format_set_shrink(XLSX::Writer::Format) is native(LIB) {*}
sub format_set_pattern(XLSX::Writer::Format, uint8) is native(LIB) {*}
sub format_set_bg_color(XLSX::Writer::Format, int32) is native(LIB) {*}
sub format_set_fg_color(XLSX::Writer::Format, int32) is native(LIB) {*}
sub format_set_border(XLSX::Writer::Format, uint8) is native(LIB) {*}
sub format_set_bottom(XLSX::Writer::Format, uint8) is native(LIB) {*}
sub format_set_top(XLSX::Writer::Format, uint8) is native(LIB) {*}
sub format_set_left(XLSX::Writer::Format, uint8) is native(LIB) {*}
sub format_set_right(XLSX::Writer::Format, uint8) is native(LIB) {*}
sub format_set_border_color(XLSX::Writer::Format, int32) is native(LIB) {*}
sub format_set_bottom_color(XLSX::Writer::Format, int32) is native(LIB) {*}
sub format_set_top_color(XLSX::Writer::Format, int32) is native(LIB) {*}
sub format_set_left_color(XLSX::Writer::Format, int32) is native(LIB) {*}
sub format_set_right_color(XLSX::Writer::Format, int32) is native(LIB) {*}

method set-bold() { format_set_bold(self) }
method set-font-color(Color:D $color) { format_set_font_color(self, $color) }
method set-font-name(Str:D $name) { format_set_font_name(self, $name) }
method set-font-size(UInt:D $size) { format_set_font_size(self, $size) }
method set-italic() { format_set_italic(self) }
method set-underline(Underline:D $underline) { format_set_underline(self, $underline) }
method set-font-strikeout() { format_set_font_strikeout(self) }
method set-font-script(Script:D $script) { format_set_font_script(self, $script) }
method set-num-format(Str:D $format) { format_set_num_format(self, $format) }
method set-num-format-index(UInt:D $index) { format_set_num_format_index(self, $index) }
method set-unlocked() { format_set_unlocked(self) }
method set-hidden() { format_set_hidden(self) }
method set-align(XLSX::Writer::Align:D $align) { format_set_align(self, $align) }
method set-text-wrap() { format_set_text_wrap(self) }
method set-indent(UInt:D $level) { format_set_indent(self, $level) }
method set-shrink() { format_set_shrink(self) }
method set-pattern(XLSX::Writer::Pattern:D $pattern) { format_set_pattern(self, $pattern) }
method set-bg-color(Color:D $color) { format_set_bg_color(self, $color) }
method set-fg-color(Color:D $color) { format_set_fg_color(self, $color) }
method set-border(XLSX::Writer::Border:D $border) { format_set_border(self, $border) }
method set-bottom(XLSX::Writer::Border:D $border) { format_set_bottom(self, $border) }
method set-top(XLSX::Writer::Border:D $border) { format_set_top(self, $border) }
method set-left(XLSX::Writer::Border:D $border) { format_set_left(self, $border) }
method set-right(XLSX::Writer::Border:D $border) { format_set_right(self, $border) }
method set-border-color(Color:D $color) { format_set_border_color(self, $color) }
method set-bottom-color(Color:D $color) { format_set_bottom_color(self, $color) }
method set-top-color(Color:D $color) { format_set_top_color(self, $color) }
method set-left-color(Color:D $color) { format_set_left_color(self, $color) }
method set-right-color(Color:D $color) { format_set_right_color(self, $color) }
method set-rotation(Int:D $rotation where { -90 <= $_ <= 90 or $_ == 270 }) { format_set_rotation(self, $rotation) }
