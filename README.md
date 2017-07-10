Perl 6 wrapper for libxlsxwriter
--

This is not a port of Perl 5's Excel::Writer::XLSX; rather, it's a wrapper for
the [libxlsxwriter](https://libxlsxwriter.github.io/) C library, which itself
was a port of the Perl 5 library. Oh the irony.

Baisc usage:

```perl6
    my $workbook = XLSX::Writer::Workbook.new("/path/to/somewhere.xlsx");
    my $worksheet = $workbook.add-worksheet("Sheeeeeeeeet");
    $worksheet.write-value(0, 0, "Hello");
    $worksheet.write-value(0, 1, pi);
    $worksheet.write-value(0, 2, True);
    $worksheet.write-value(0, 3, DateTime.now);

    $workbook.close();
```

See `t/` directory for more functions, including custom formatting.
