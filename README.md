Music Cleanup Tool

This script came in handy when I needed to recover a ton of music from a disparate 
set of source devices. It traverses a directory tree one folder at a time, finds all the songs,
and places them into a target folder under [album artist]\[album]\[track#][track name].[existing extension]

It has a number of very specific workarounds for some issues I encountered, and I didn't fully
work through powershell escaping rules, so there are some rough edges when songs have strange characters.
It has good logging though, and a "mock" mode so you can see what it will do without making changes.
Notably, it'll log a warning if a file is very small or empty (e.g. questionable data from data recovery).
It quietly skips duplicates, since it's designed to merge results together from various sources.

One particularly helpful tidbit is handling of song metadata. I found some good examples that were functional
but wrong or over-simplified. This code should be more robust, since it dynamically reads property IDs rather
than using hard coded values.

All in all, it was pretty useful for me, and I hope if you're looking for a way to organize your mesic 
collection, you find this to be a helpful starting point.

Typical usage:

    normalize.ps1 -SourceDir . -TargetDir "m:\MergedMusic"

Additional options:

  -Verbose $True
        Detailed output

  -Mock $True
        Show what will happen, but don't copy anything.