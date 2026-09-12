-- Usage: osascript export-front.applescript <output.xlsx>
-- Exports Numbers' front document as an Excel workbook. Open the input first with `open -a Numbers <file>`
-- (LaunchServices import) - the AppleScript `open` verb produced empty documents for xlsx inputs on Numbers 14.4.
on run argv
  set outputPath to item 1 of argv
  tell application "Numbers"
    set waited to 0
    repeat until (count of documents) > 0 or waited > 90
      delay 1
      set waited to waited + 1
    end repeat
    delay 3
    set doc to front document
    set sheetCount to count of sheets of doc
    set cellCount to 0
    try
      set cellCount to count of cells of table 1 of sheet 1 of doc
    end try
    with timeout of 120 seconds
      export doc to (POSIX file outputPath) as Microsoft Excel
    end timeout
    close doc saving no
    return "exported " & outputPath & " (" & sheetCount & " sheets, table1 cells " & cellCount & ")"
  end tell
end run
