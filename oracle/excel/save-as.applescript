-- Usage: osascript save-as.applescript <input> <output.xlsx> <format> <date1904>
--   format:   xlsx     Excel Workbook (Transitional OOXML, "Excel XML file format")
--             strict   NOT automatable on Excel 16 for Mac (no Strict member in the file-format enumeration and no
--                      ad-hoc VBA hook) - reported as STRICT-UNSUPPORTED; save by hand (STEPS.md)
--             xls      Excel 97-2004 (BIFF8, "Excel98to2004 file format") - renamed-legacy-format hostile source
--             xlsb     Excel Binary Workbook ("Excel binary file format") - renamed-format hostile source
--   ("workbook normal file format" is NOT BIFF8 on Excel 16 - it is the default .xlsx.)
--   date1904: true|false - switch the workbook to the 1904 date system before saving
-- Produces Excel-generated goldens from any input Excel can open (CSV, xlsx, ...). Needs a full Excel license.
on run argv
  set inputPath to item 1 of argv
  set outputPath to item 2 of argv
  set fileFormatName to item 3 of argv
  set useDate1904 to (item 4 of argv is "true")
  set tabChar to "|||"
  set notes to {}
  tell application "Microsoft Excel"
    activate
    set display alerts to false
    try
      close every workbook saving no
    end try
    with timeout of 60 seconds
      open (POSIX file inputPath)
    end timeout
    set wb to active workbook
    set wbName to name of wb
    if useDate1904 then
      set date 1904 of wb to true
      set end of notes to "date1904=true"
    end if
    -- delete any previous output so save-as never prompts about overwriting
    do shell script "rm -f " & quoted form of outputPath
    if fileFormatName is "xlsx" then
      save workbook as wb filename outputPath file format Excel XML file format with overwrite
    else if fileFormatName is "xls" then
      save workbook as wb filename outputPath file format Excel98to2004 file format with overwrite
    else if fileFormatName is "xlsb" then
      save workbook as wb filename outputPath file format Excel binary file format with overwrite
    else if fileFormatName is "strict" then
      -- Excel 16's AppleScript dictionary has no Strict Open XML file-format member and no way to run ad-hoc
      -- VBA (`run VB macro` needs a macro inside the workbook), so the Strict golden stays a manual save.
      set end of notes to "STRICT-UNSUPPORTED" & tabChar & "use File > Save As > Strict Open XML Spreadsheet by hand"
    else
      error "unknown format " & fileFormatName
    end if
    try
      close active workbook saving no
    end try
    set display alerts to true
  end tell
  set AppleScript's text item delimiters to linefeed
  set body to notes as string
  set AppleScript's text item delimiters to ""
  return "saved=" & outputPath & linefeed & "source=" & wbName & linefeed & body
end run
