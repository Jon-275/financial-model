on run argv
    set actionName to item 1 of argv
    set expectedName to item 2 of argv
    set expectedPath to item 3 of argv
    set expectedHfs to (POSIX file expectedPath) as text
    set timeoutSeconds to 180
    if actionName is "save" then set timeoutSeconds to 240
    if actionName is "run" then set timeoutSeconds to 7200
    if actionName is "open" then
        set expectedHash to item 4 of argv
        set hashLine to do shell script "/usr/bin/shasum -a 256 " & quoted form of expectedPath
        if hashLine does not start with (expectedHash & "  ") then error "The exact session workbook fingerprint differs."
        set excelBundle to POSIX path of (path to application "Microsoft Excel")
    end if
    with timeout of timeoutSeconds seconds
        tell application "Microsoft Excel"
            if actionName is "quit-empty" then
                if (count workbooks) is not 0 then error "Excel is not empty. It was left open."
                my requireExcelProcess(item 4 of argv)
                ignoring application responses
                    quit
                end ignoring
                return "quit-requested"
            end if
            if actionName is "open" then
                if (count workbooks) is not 0 then error "Use a new empty Excel process."
                set securityBefore to automation security
                open workbook workbook file name expectedPath update links do not update links read only false ignore read only recommended true add to mru false
                if automation security is not securityBefore then error "Excel's automation security setting changed during open."
            end if
            if (count workbooks) is not 1 then error "Expected exactly one benchmark workbook."
            set book to active workbook
            set actualPath to full name of book
            if (name of book) is not expectedName or (actualPath is not expectedPath and actualPath is not expectedHfs) then error "The open workbook does not match the verified session path."

            if actionName is "open" then return "policy-unchanged|" & excelBundle
            if actionName is "close" then
                my requireExcelProcess(item 4 of argv)
                close book saving no
                if (count workbooks) is not 0 then error "Excel is not empty after closing the session workbook."
                return "closed"
            end if
            if actionName is "run" then
                run VB Macro ("'" & expectedName & "'!RunMarketRisk")
                return "returned"
            end if
            if actionName is "profile" then return run VB Macro ("'" & expectedName & "'!ReadMarketRiskProfile")
            if actionName is "guarded" then return run VB Macro ("'" & expectedName & "'!ReadGuardedCalculationCounts")
            if actionName is "proof" then return run VB Macro ("'" & expectedName & "'!ReadFundingProofCounts")
            if actionName is not "save" then error "Unknown benchmark API action."

            set riskData to worksheet "Risk_Data" of book
            if ((value of range "Z1" of riskData) as text) is not "current" then error "Risk results are not current."
            if ((value of range "Z6" of riskData) as text) is not "complete" then error "The native batch is incomplete."
            if ((value of range "Z10" of riskData) as integer) is not 1000 then error "Expected 1,000 pairs."
            set summaryStatus to (value of range "K22" of worksheet "Summary" of book) as text
            if summaryStatus is not "Risk results are current." then error "The report has no current verified result."
            set alertsBefore to display alerts
            try
                set display alerts to false
                save book
            on error saveText number saveNumber
                try
                    set display alerts to alertsBefore
                    if display alerts is not alertsBefore then error "Excel's previous alert setting was not restored."
                on error restoreText
                    error (saveText & " (alert restoration: " & restoreText & ")") number saveNumber
                end try
                error saveText number saveNumber
            end try
            set display alerts to alertsBefore
            if display alerts is not alertsBefore then error "Excel's previous alert setting was not restored."
            return summaryStatus
        end tell
    end timeout
end run

on requireExcelProcess(expectedPid)
    set runningPid to do shell script "/usr/bin/pgrep -x 'Microsoft Excel'"
    if runningPid is not expectedPid then error "The original benchmark Excel process changed."
end requireExcelProcess
