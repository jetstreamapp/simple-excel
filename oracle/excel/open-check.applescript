-- Usage: osascript open-check.applescript <file.xlsx> [sheetName] [refs comma-separated]
-- Read-only Excel oracle: opens the workbook (a repair dialog blocks `open`, which the caller turns into a
-- timeout = HUNG verdict), reports the workbook name (Excel appends "[Repaired]" after a repair), the sheet
-- names, and the value + displayed text of sentinel cells. Never saves.
on run argv
  set inputPath to item 1 of argv
  set sheetName to "Data"
  if (count of argv) > 1 then set sheetName to item 2 of argv
  set refList to {"A2", "B2", "D5", "E2", "F2", "F3", "F7", "G2", "I2", "J2", "J11", "M2", "M3", "M4", "N2", "O2", "O6", "P2", "S30"}
  if (count of argv) > 2 then
    set AppleScript's text item delimiters to ","
    set refList to text items of item 3 of argv
    set AppleScript's text item delimiters to ""
  end if
  set outputLines to {}
  set tabChar to "|||" -- field separator (`tab` is an Excel term inside the tell block; cell values may contain real tabs)
  set wbName to ""
  set sheetNamesText to ""
  tell application "Microsoft Excel"
    activate
    -- every golden is named canonical.xlsx, so a workbook left open by an earlier run would be addressed instead
    try
      close every workbook saving no
    end try
    try
      with timeout of 25 seconds
        open (POSIX file inputPath)
      end timeout
    on error errMsg
      return "STEP-open" & tabChar & errMsg
    end try
    try
      set wbName to name of active workbook
    on error errMsg
      try
        set wbName to name of workbook 1
      on error errMsg2
        return "STEP-name" & tabChar & errMsg & " / " & errMsg2
      end try
    end try
    try
      set sheetList to name of every worksheet of workbook wbName
      set AppleScript's text item delimiters to "|"
      set sheetNamesText to sheetList as string
      set AppleScript's text item delimiters to ""
    on error errMsg
      set end of outputLines to "STEP-sheets" & tabChar & errMsg
    end try
    try
      repeat with cellRef in refList
        set rawValue to ""
        try
          set rawValue to value of range (cellRef as string) of worksheet sheetName of workbook wbName
        on error errMsg
          set rawValue to "ERR " & errMsg
        end try
        set shownText to ""
        try
          set shownText to string value of range (cellRef as string) of worksheet sheetName of workbook wbName
        on error
          set shownText to ""
        end try
        set end of outputLines to (cellRef as string) & tabChar & (rawValue as string) & tabChar & (shownText as string)
      end repeat
    on error errMsg
      set end of outputLines to "STEP-cells" & tabChar & errMsg
    end try
    try
      close workbook wbName saving no
    on error errMsg
      set end of outputLines to "STEP-close" & tabChar & errMsg
    end try
  end tell
  set AppleScript's text item delimiters to linefeed
  set body to outputLines as string
  set AppleScript's text item delimiters to ""
  return "workbook=" & wbName & linefeed & "sheets=" & sheetNamesText & linefeed & body
end run
