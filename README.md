Automatically format document on save. Current release includes:

* RemoveAndSort: Same as Edit > IntelliSense > Remove and Sort Usings
* SmartRemoveAndSort: Apply remove and sort to .cs without #if.
* FormatDoucment: Same as Edit > Advance > Format Document
* UnifyLineBreak: Enforce line break to CRLF or LF.
* UnifyEndOfFile: Enforce only one line return at the end of file.
* TabToSpace: Convert tab to space.
* ForceUtf8WithBom: Force file encoding to UTF8 with BOM.
* RemoveTrailingSpaces: Remove trailing spaces. It is mostly for Visual Sutdio 2012, which won't remove trailing spaces when formatting.
* File extension filters.
* Batch formatting in solution explorer.
* Settings in Tools -> Options -> Format on Save.

I found it convenient to unify source code format throughout the develope team. If you have any suggestion, feel free to tell me.

## Updates

### v1.12

* Reactivate current editing document after saving all.

### v1.11

* Add SmartRemoveAndSort: Apply remove and sort to .cs without #if.
* Add RemoveTrailingSpaces: Remove trailing spaces. It is mostly for Visual Sutdio 2012, which won't remove trailing spaces when formatting.
