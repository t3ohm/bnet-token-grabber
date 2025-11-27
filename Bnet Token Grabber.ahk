#NoEnv
#SingleInstance Force

; Disable tray icon and default menus
Menu, Tray, NoIcon
#NoTrayIcon

; Your credentials
global EMAIL := ""
global PASSWORD := ""
global url := "https://us.battle.net/login/en/?externalChallenge=login&app=osi"
global ie := ""
global iePID := ""
global shown := false
global extractedToken := ""
global tokenMasked := true
global CLI_MODE := false
global CLI_EXECUTE := false

; Parse command line arguments
if (A_Args.Length() > 0) {
    CLI_MODE := true
    
    ; Parse arguments
    for index, param in A_Args {
        if (param = "--email" || param = "-e") && (index < A_Args.Length()) {
            EMAIL := A_Args[index + 1]
        }
        else if (param = "--password" || param = "-p") && (index < A_Args.Length()) {
            PASSWORD := A_Args[index + 1]
        }
        else if (param = "--execute" || param = "-x") {
            CLI_EXECUTE := true
        }
        else if (param = "--help" || param = "-h") {
            ShowHelp()
            ExitApp
        }
    }
    
    ; If --execute flag is passed, run headless
    if (CLI_EXECUTE) {
        RunHeadless()
        ExitApp
    }
}

ShowHelp() {
    MsgBox, 64, Battle.net Token Grabber - Help, 
    (
Usage:
    TokenGrabber.exe [options]

Options:
    -e, --email <email>       Set email/username
    -p, --password <pass>     Set password
    -x, --execute             Execute login and return token (required for CLI mode)
    -h, --help                Show this help message

Examples:
    Open GUI with prefilled credentials:
        TokenGrabber.exe --email "user@example.com" --password "pass123"
    
    Run headless and output token:
        TokenGrabber.exe -e "user@example.com" -p "pass123" --execute
    
    Open GUI with default credentials:
        TokenGrabber.exe

Notes:
    - Without --execute flag, the GUI will open with provided credentials
    - With --execute flag, runs in CLI mode and outputs token to stdout
    - Exit codes: 0 = success, 1 = failure
    )
}

; --- GUI SETUP ---
if (!CLI_MODE) {
    Gui, Add, Text,, Username:
    Gui, Add, Edit, vUserInput w450, %EMAIL%

    Gui, Add, Text,, Password:
    Gui, Add, Edit, vPassInput w450 Password, %PASSWORD%

    Gui, Add, Checkbox, vToggleIE gToggleIEVisibility, Browser Visible

    Gui, Add, Button, gStartLogin w450, Start Login

    Gui, Add, Button, gCopyToken w220 xm, Copy Token to Clipboard
    Gui, Add, Button, gToggleMask w220 x+10, Toggle Token Masking

    Gui, Add, Text, xm, Status Log:
    Gui, Add, Edit, vStatusLog w450 h200 ReadOnly -Wrap +VScroll

    Gui, Show, , Battle.net Token Grabber
}
return

; --- LOG FUNCTION ---
LogStatus(message) {
    global tokenMasked, extractedToken, CLI_MODE
    
    ; In CLI mode, output to stdout
    if (CLI_MODE) {
        FileAppend, %message%`n, *
        return
    }
    
    GuiControlGet, currentLog,, StatusLog
    timestamp := A_Hour ":" A_Min ":" A_Sec
    
    ; Mask token in log if masking is enabled
    displayMessage := message
    if (tokenMasked && extractedToken != "" && InStr(message, extractedToken)) {
        displayMessage := StrReplace(message, extractedToken, "********" SubStr(extractedToken, -8))
    }
    
    newLog := "[" timestamp "] " displayMessage "`r`n" currentLog
    GuiControl,, StatusLog, %newLog%
}

; --- KILL IE PROCESS ---
KillIEProcess() {
    global iePID
    if (iePID != "") {
        LogStatus("Killing IE process (PID: " iePID ")...")
        Process, Close, %iePID%
        Sleep, 500
        Process, Exist, %iePID%
        if (ErrorLevel = 0) {
            LogStatus("IE process closed successfully")
        } else {
            LogStatus("Warning: IE process may still be running")
        }
        iePID := ""
    }
}

; --- KILL ALL IE PROCESSES (SAFETY CLEANUP) ---
KillAllIEProcesses() {
    LogStatus("Performing safety cleanup of all IE processes...")
    Loop {
        Process, Exist, iexplore.exe
        if (ErrorLevel = 0)
            break
        Process, Close, iexplore.exe
        Sleep, 200
    }
}

DoForceLogout() {
    global iePID
    LogStatus("Starting force logout...")
    try {
        tempIE := ComObjCreate("InternetExplorer.Application")
        tempIE.Visible := false
        
        ; Get PID of this temporary IE instance
        tempPID := GetIEProcessID(tempIE)
        LogStatus("Temporary IE created (PID: " tempPID ")")
        
        tempIE.Navigate("https://us.battle.net/login/logout")
        WaitForLoad(tempIE)
        Sleep, 1500
        tempIE.Quit()
        
        ; Kill the temp IE process
        if (tempPID != "") {
            Process, Close, %tempPID%
            Sleep, 500
        }
        
        LogStatus("Logout complete")
    } catch e {
        LogStatus("Error during logout: " e)
    }
    Sleep, 1000
}

; --- GET IE PROCESS ID ---
GetIEProcessID(ieObj) {
    try {
        ; Get all iexplore.exe processes and their creation times
        ; Find the most recently created one (our new instance)
        latestPID := ""
        latestTime := ""
        
        for proc in ComObjGet("winmgmts:").ExecQuery("SELECT ProcessId, CreationDate FROM Win32_Process WHERE Name='iexplore.exe'") {
            if (proc.ProcessId && proc.CreationDate) {
                if (latestTime = "" || proc.CreationDate > latestTime) {
                    latestTime := proc.CreationDate
                    latestPID := proc.ProcessId
                }
            }
        }
        return latestPID
    }
    return ""
}

AutomateLoginIE() {
    global url, EMAIL, PASSWORD, ie, iePID, CLI_MODE, CLI_EXECUTE
    
    try {
        ; Create Internet Explorer COM object
        LogStatus("Creating Internet Explorer instance...")
        ie := ComObjCreate("InternetExplorer.Application")
        
        ; In CLI execute mode, hide browser; otherwise check the GUI checkbox
        if (CLI_EXECUTE) {
            ie.Visible := false
            LogStatus("Browser hidden (CLI mode)")
        } else {
            GuiControlGet, visible,, ToggleIE
            ie.Visible := visible
        }
        
        ; Get the process ID with retry logic
        Sleep, 500
        maxRetries := 5
        Loop, %maxRetries% {
            iePID := GetIEProcessID(ie)
            if (iePID != "") {
                LogStatus("IE Process ID captured: " iePID)
                break
            }
            Sleep, 500
        }
        
        if (iePID = "") {
            LogStatus("ERROR: Could not capture IE Process ID after " maxRetries " attempts")
            ; Still continue but warn user
        }
        
        ; Navigate to login page
        LogStatus("Navigating to login page...")
        ie.Navigate(url)
        
        ; Wait for page to load completely
        WaitForLoad(ie)
        LogStatus("Page loaded")
        
        ; Check which page we're on and handle accordingly
        doc := ie.document
        
        ; Scenario 1: We're on the email/phone entry page (fresh browser)
        if (IsEmailEntryPage(doc)) {
            LogStatus("Detected email entry page")
            EnterEmailAndContinue(ie, doc)
            
            ; Wait for password page to load
            Sleep, 3000
            WaitForLoad(ie)
            doc := ie.document  ; Refresh document object
        }
        
        ; Scenario 2: We're directly on the password page (browser remembers us)
        if (IsPasswordPage(doc)) {
            LogStatus("Detected password page")
            EnterPasswordAndSubmit(ie, doc)
            
            ; Wait for the URL to change and contain the ST parameter
            LogStatus("Waiting for ST parameter in URL...")
            
            ; Use JavaScript method to get URL (since this worked)
            finalURL := WaitForSTParameterJS(ie, 15000)
            if (finalURL != "") {
                LogStatus("ST parameter found in URL")
                
                ; Extract the ST parameter
                extractedST := ExtractSTParameter(finalURL)
                LogStatus("ST token extracted successfully")
                
                ; Store token globally
                extractedToken := extractedST
                
                ; Close IE properly
                try ie.Quit()
                Sleep, 500
                KillIEProcess()
                
                return extractedST
            } else {
                LogStatus("ERROR: ST parameter not found within timeout")
                ; Fallback to address bar URL for debugging
                addressBarURL := ie.LocationURL
                jsURL := GetJavaScriptURL(ie)
                LogStatus("Address bar URL: " addressBarURL)
                LogStatus("JavaScript URL: " jsURL)
                
                try ie.Quit()
                Sleep, 500
                KillIEProcess()
                
                return ""
            }
        } else {
            LogStatus("ERROR: Unexpected page detected")
            currentURL := GetJavaScriptURL(ie)
            LogStatus("Current URL: " currentURL)
            
            try ie.Quit()
            Sleep, 500
            KillIEProcess()
            
            return ""
        }
        
    } catch e {
        LogStatus("CRITICAL ERROR: " e)
        try ie.Quit()
        Sleep, 500
        KillIEProcess()
        return ""
    }
}

; Function to wait for ST parameter using JavaScript URL method
WaitForSTParameterJS(ie, timeout) {
    startTime := A_TickCount
    while (A_TickCount - startTime < timeout) {
        ; Use JavaScript location.href (the method that worked)
        jsURL := GetJavaScriptURL(ie)
        if (jsURL != "" && InStr(jsURL, "localhost") && InStr(jsURL, "ST=")) {
            return jsURL
        }
        Sleep, 500
    }
    return ""
}

; Function to get URL via JavaScript location.href (the reliable method)
GetJavaScriptURL(ie) {
    try {
        doc := ie.document
        jsURL := doc.parentWindow.location.href
        return jsURL
    } catch {
        return ""
    }
}

; Function to detect if we're on the email entry page
IsEmailEntryPage(doc) {
    try {
        accountNameField := doc.getElementById("accountName")
        return !!accountNameField
    } catch {
        return false
    }
}

; Function to detect if we're on the password page
IsPasswordPage(doc) {
    try {
        passwordField := doc.getElementById("password")
        return !!passwordField
    } catch {
        return false
    }
}

; Function to enter email and click Continue
EnterEmailAndContinue(ie, doc) {
    global EMAIL
    try {
        emailField := doc.getElementById("accountName")
        if emailField {
            emailField.value := EMAIL
            LogStatus("Email entered: " EMAIL)
        } else {
            LogStatus("ERROR: Email field not found")
            return false
        }
        
        continueButton := doc.getElementById("submit")
        if continueButton {
            continueButton.click()
            LogStatus("Continue button clicked")
            return true
        }
        LogStatus("ERROR: Continue button not found")
        return false
    } catch e {
        LogStatus("ERROR entering email: " e)
        return false
    }
}

; Function to enter password and submit
EnterPasswordAndSubmit(ie, doc) {
    global PASSWORD
    try {
        passwordField := doc.getElementById("password")
        if passwordField {
            passwordField.value := PASSWORD
            LogStatus("Password entered")
        } else {
            LogStatus("ERROR: Password field not found")
            return false
        }
        
        submitButton := doc.getElementById("submit")
        if submitButton {
            submitButton.click()
            LogStatus("Submit button clicked")
            return true
        }
        LogStatus("ERROR: Submit button not found")
        return false
    } catch e {
        LogStatus("ERROR entering password: " e)
        return false
    }
}

; Function to extract ST parameter
ExtractSTParameter(token_url) {
    if RegExMatch(token_url, "[?&]ST=([^&]+)", match) {
        stValue := match1
        return stValue
    } else {
        return ""
    }
}

; Helper function to wait for page load
WaitForLoad(ie) {
    while ie.readyState != 4 || ie.document.readyState != "complete" || ie.busy
        Sleep, 100
}

; --- CLEANUP AND EXIT ---
CleanupAndExit(exitCode := 0) {
    global ie, iePID
    LogStatus("Cleaning up resources...")
    
    ; Try to close IE gracefully
    try {
        if (ie)
            ie.Quit()
    }
    
    Sleep, 500
    
    ; Force kill IE process if still running
    KillIEProcess()
    
    ; Also kill any orphaned IE processes (safety measure)
    KillAllIEProcesses()
    
    ExitApp, %exitCode%
}

; --- HEADLESS EXECUTION ---
RunHeadless() {
    global EMAIL, PASSWORD, extractedToken
    
    LogStatus("=== CLI Mode: Executing Login ===")
    LogStatus("Email: " EMAIL)
    
    ; Force logout first
    DoForceLogout()
    
    ; Run login flow
    extractedToken := AutomateLoginIE()
    
    if (extractedToken != "") {
        ; Output token to stdout
        FileAppend, TOKEN=%extractedToken%`n, *
        LogStatus("SUCCESS: Token retrieved")
        ; Cleanup before exit
        CleanupAndExit(0)
    } else {
        LogStatus("ERROR: Failed to retrieve token")
        ; Cleanup before exit
        CleanupAndExit(1)
    }
}

; --- TOGGLE IE VISIBILITY ---
ToggleIEVisibility:
    GuiControlGet, visible,, ToggleIE
    if (ie) {
        ie.Visible := visible
        LogStatus("IE visibility toggled: " (visible ? "Visible" : "Hidden"))
    }
return

; --- COPY TOKEN TO CLIPBOARD ---
CopyToken:
    global extractedToken
    if (extractedToken != "") {
        Clipboard := extractedToken
        LogStatus("Token copied to clipboard manually")
        MsgBox, 64, Success, Token copied to clipboard!
    } else {
        LogStatus("ERROR: No token available to copy")
        MsgBox, 48, No Token, No token has been extracted yet. Please run the login process first.
    }
return

; --- TOGGLE TOKEN MASKING ---
ToggleMask:
    global tokenMasked, extractedToken
    tokenMasked := !tokenMasked
    
    ; Rebuild the log with new masking state
    GuiControlGet, currentLog,, StatusLog
    if (extractedToken != "") {
        if (tokenMasked) {
            ; Mask the token
            maskedLog := StrReplace(currentLog, extractedToken, "********" SubStr(extractedToken, -8))
            GuiControl,, StatusLog, %maskedLog%
            LogStatus("Token masking enabled")
        } else {
            ; Unmask the token
            maskedToken := "********" SubStr(extractedToken, -8)
            unmaskedLog := StrReplace(currentLog, maskedToken, extractedToken)
            GuiControl,, StatusLog, %unmaskedLog%
            LogStatus("Token masking disabled")
        }
    } else {
        LogStatus("Token masking toggled: " (tokenMasked ? "Enabled" : "Disabled"))
    }
return

; --- START LOGIN ---
StartLogin:
    GuiControlGet, EMAIL,, UserInput
    GuiControlGet, PASSWORD,, PassInput

    if (EMAIL = "" || PASSWORD = "") {
        LogStatus("ERROR: Missing username or password")
        MsgBox, 48, Missing Info, Please enter both Username and Password.
        return
    }

    LogStatus("=== Login Process Started ===")
    LogStatus("Username: " EMAIL)

    ; Force logout first
    DoForceLogout()
    
    ; Run your existing login flow
    extractedST := AutomateLoginIE()

    if (extractedST != "") {
        extractedToken := extractedST
        LogStatus("SUCCESS: Token retrieved and ready to copy")
        LogStatus("Token: " (tokenMasked ? "********" SubStr(extractedST, -8) : extractedST))
        MsgBox, 64, Success, Token retrieved successfully! Use 'Copy Token to Clipboard' button to copy it.
    } else {
        LogStatus("FAILED: Could not retrieve ST token")
        MsgBox, 16, Failed, Could not retrieve ST token.
    }
    
    LogStatus("=== Login Process Completed ===")
return

GuiClose:
    LogStatus("Closing application...")
    try ie.Quit()
    Sleep, 500
    KillIEProcess()
    KillAllIEProcesses()
ExitApp

; Prevent script from exiting in GUI mode
#If (!CLI_MODE)
return
#If