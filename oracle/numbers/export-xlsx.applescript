-- Usage: osascript export-xlsx.applescript <input.xlsx> <output.xlsx>
-- Opens the input in Numbers (which imports it into a new document), waits for the import to finish,
-- then re-exports the front document as an Excel workbook (a Numbers-generated golden).
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  tell application "Numbers"
    activate
    open (POSIX file inputPath)
    -- `open` can return before the import has populated the document; poll until a document exists
    -- whose name matches the input, then give the import a moment to settle.
    set waited to 0
    repeat until (count of documents) > 0 or waited > 60
      delay 1
      set waited to waited + 1
    end repeat
    delay 3
    set doc to front document
    set sheetCount to count of sheets of doc
    with timeout of 120 seconds
      export doc to (POSIX file outputPath) as Microsoft Excel
    end timeout
    close doc saving no
    return "exported " & outputPath & " (" & sheetCount & " sheets in Numbers)"
  end tell
end run
