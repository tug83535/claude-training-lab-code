Attribute VB_Name = "modDirector"
Option Explicit

'===============================================================================
' modDirector - Master Director Macro for Video Demo Automation
' iPipeline Finance Automation - Video Recording Puppeteer
'===============================================================================
' PURPOSE:  Automates the entire video demo for OBS screen recording.
'           Plays AI narration audio, navigates sheets, triggers macros,
'           scrolls, types, and pauses — all timed to match the script.
'
' USAGE:    1. Import this module into the demo Excel file
'           2. Set AUDIO_BASE_PATH below to your AudioClips folder
'           3. Run "RunPreflight" to verify everything is ready
'           4. Run "RunVideo1", "RunVideo2", or "RunVideo3"
'           5. Or run "RunAllVideos" for the full show
'
' AUDIO:    Uses Windows mciSendString API for invisible MP3 playback.
'           No media player window appears on screen.
'           Clip durations are measured automatically at runtime.
'
' TIMING:   Uses kernel32 Sleep API. All durations in milliseconds.
'           Audio clip lengths are auto-detected via mciSendString —
'           no hardcoded duration constants needed.
'
' SAFETY:   - Pre-flight checks verify audio files, sheets, and window
'             state before each video starts
'           - MCI audio device is reset at every entry point to prevent
'             stuck state from interrupted runs
'           - Clips 19, 22, 23 bypass InputBox/FileDialog directly
'             (no fragile SendKeys for modal dialogs)
'
' PUBLIC SUBS:
'   RunAllVideos          - Run all 3 videos back-to-back
'   RunVideo1             - Video 1: "What's Possible" (Clips 1-7)
'   RunVideo2             - Video 2: "Full Demo Walkthrough" (Clips 8-26)
'   RunVideo3             - Video 3: "Universal Tools" (Clips 27-39)
'   TestClip N            - Test any single clip by number (1-38)
'   QuickTest             - Quick audio + scroll + preflight test
'   RunPreflight          - Full system check before recording
'   CleanupAllOutputSheets - Delete all macro-generated sheets
'
' VERSION:  2.0.0
' DATE:     2026-03-25
' CHANGES:  v1.0 -> v2.0:
'           + Runtime MP3 duration detection (no hardcoded constants)
'           + MCI cleanup/reset at every entry point
'           + Clips 19/22/23 bypass modal dialogs directly
'           + Pre-flight checks (audio files, sheets, window state)
'           + RunPreflight public verification sub
'===============================================================================

'===============================================================================
' WINDOWS API DECLARATIONS
'===============================================================================
#If VBA7 Then
    ' 64-bit Office / VBA7
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function mciSendStringA Lib "winmm.dll" _
        (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
         ByVal uReturnLength As Long, ByVal hwndCallback As LongPtr) As Long
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
    ' 32-bit Office
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function mciSendStringA Lib "winmm.dll" _
        (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
         ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

'===============================================================================
' CONFIGURATION — EDIT THESE BEFORE RUNNING
'===============================================================================

' >>> AUDIO FILE PATH — Set this to YOUR AudioClips folder <<<
' Use the FULL path ending with a backslash.
' Example: "C:\Users\connor.atlee\Desktop\FinalExport\AudioClips\"
Private Const AUDIO_BASE_PATH As String = "C:\Users\connor.atlee\.claude\projects\claude-training-lab-code\FinalExport\AudioClips\"

' >>> TYPING SPEED — milliseconds between each character <<<
Private Const TYPING_DELAY_MS As Long = 90

' >>> SCROLL SPEED — milliseconds between each scroll step <<<
Private Const SCROLL_STEP_DELAY As Long = 250

' >>> SILENCE PADDING — seconds of silence at start/end of each clip <<<
Private Const SILENCE_PAD_SEC As Double = 2#

' >>> ABORT KEY — Hold Escape during playback to abort <<<
' (checked between clips, not during Sleep)

'===============================================================================
' MODULE-LEVEL STATE
'===============================================================================
' m_ClipDurSec is auto-measured each time PlayAudio is called.
' It replaces all hardcoded duration constants from v1.0.
' Every clip sub can reference it for "remaining time" calculations.
Private m_ClipDurSec As Double
Private m_Aborted As Boolean

'===============================================================================
'
'    SECTION 1: HELPER FUNCTIONS
'
'===============================================================================

'===============================================================================
' WaitMs - Pause execution for N milliseconds, pumping events
'===============================================================================
Private Sub WaitMs(ByVal ms As Long)
    If ms <= 0 Then Exit Sub
    Dim endTick As Long
    endTick = GetTickCount() + ms
    Do While GetTickCount() < endTick
        DoEvents
        Sleep 50   ' yield CPU in 50ms chunks so Excel stays responsive
    Loop
End Sub

'===============================================================================
' WaitSec - Pause execution for N seconds
'===============================================================================
Private Sub WaitSec(ByVal sec As Double)
    WaitMs CLng(sec * 1000)
End Sub

'===============================================================================
' ResetMCI - Close any existing MCI audio device to prevent stuck state.
' MUST be called at the top of every public entry point (RunVideo1, etc.)
' If a prior run was interrupted (Ctrl+Break, error, crash), the MCI device
' stays open and all future audio silently fails. This fixes that.
'===============================================================================
Private Sub ResetMCI()
    mciSendStringA "stop director_audio", vbNullString, 0, 0
    mciSendStringA "close director_audio", vbNullString, 0, 0
    m_ClipDurSec = 0
End Sub

'===============================================================================
' GetOpenClipDurationSec - Query the duration of the currently-open MCI clip.
' Returns duration in seconds. Falls back to 30 sec if query fails.
' Uses mciSendString "status <alias> length" which returns milliseconds.
'===============================================================================
Private Function GetOpenClipDurationSec() As Double
    Dim buf As String * 128
    Dim ret As Long

    ' Ensure time format is milliseconds
    mciSendStringA "set director_audio time format milliseconds", vbNullString, 0, 0

    ' Query length
    ret = mciSendStringA("status director_audio length", buf, 128, 0)

    Dim ms As Long
    ms = Val(buf)

    If ms > 0 Then
        GetOpenClipDurationSec = ms / 1000#
        Debug.Print "[Director] Clip duration measured: " & Format(GetOpenClipDurationSec, "0.0") & "s"
    Else
        ' Fallback — measurement failed (corrupt file, unsupported format, etc.)
        GetOpenClipDurationSec = 30
        Debug.Print "[Director] WARNING: Could not measure clip duration (ret=" & ret & "). Using 30s fallback."
    End If
End Function

'===============================================================================
' PlayAudio - Play an MP3 file in the background (invisible, no player window)
' Also measures the clip duration and stores it in m_ClipDurSec.
' Parameters:
'   clipPath  - Full path to the MP3 file
'===============================================================================
Private Sub PlayAudio(ByVal clipPath As String)
    ' Close any previous audio
    mciSendStringA "close director_audio", vbNullString, 0, 0

    ' Open the new clip
    Dim cmd As String
    cmd = "open """ & clipPath & """ alias director_audio"
    Dim ret As Long
    ret = mciSendStringA(cmd, vbNullString, 0, 0)

    If ret <> 0 Then
        Debug.Print "[Director] WARNING: MCI open failed (ret=" & ret & ") for: " & clipPath
        m_ClipDurSec = 30   ' fallback
    Else
        ' Measure duration BEFORE playing
        m_ClipDurSec = GetOpenClipDurationSec()
    End If

    ' Play
    mciSendStringA "play director_audio", vbNullString, 0, 0
End Sub

'===============================================================================
' StopAudio - Stop and close any playing audio
'===============================================================================
Private Sub StopAudio()
    mciSendStringA "stop director_audio", vbNullString, 0, 0
    mciSendStringA "close director_audio", vbNullString, 0, 0
End Sub

'===============================================================================
' PlayClip - Play an audio clip by subfolder and filename, wait for it to finish.
' Duration is auto-measured — no manual timing needed.
' Parameters:
'   subFolder  - "Video1", "Video2", or "Video3"
'   fileName   - e.g., "V1_S1_Opening_Hook.mp3"
'===============================================================================
Private Sub PlayClip(ByVal subFolder As String, ByVal fileName As String)
    Dim fullPath As String
    fullPath = AUDIO_BASE_PATH & subFolder & "\" & fileName

    ' Verify file exists
    If Dir(fullPath) = "" Then
        Debug.Print "[Director] WARNING: Audio file not found: " & fullPath
        ' Still wait a default duration so timing stays correct
        m_ClipDurSec = 30
        WaitSec m_ClipDurSec
        Exit Sub
    End If

    PlayAudio fullPath
    ' Wait for the measured duration
    WaitSec m_ClipDurSec
End Sub

'===============================================================================
' SimulateType - Type text character-by-character into the active cell or control
' Uses SendKeys for visual typing effect on camera.
'===============================================================================
Private Sub SimulateType(ByVal text As String, Optional ByVal delayMs As Long = 0)
    If delayMs = 0 Then delayMs = TYPING_DELAY_MS
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(text)
        ch = Mid(text, i, 1)
        ' SendKeys special characters need braces
        Select Case ch
            Case "+", "^", "%", "~", "(", ")", "{", "}"
                SendKeys "{" & ch & "}", True
            Case Else
                SendKeys ch, True
        End Select
        DoEvents
        Sleep CLng(delayMs)
    Next i
End Sub

'===============================================================================
' GoToSheet - Activate a sheet by name (safe — won't crash if missing)
'===============================================================================
Private Sub GoToSheet(ByVal shName As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(shName).Activate
    DoEvents
    If Err.Number <> 0 Then
        Debug.Print "[Director] Sheet not found: " & shName
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' SelectCell - Select a cell on the active sheet
'===============================================================================
Private Sub SelectCell(ByVal addr As String)
    On Error Resume Next
    ActiveSheet.Range(addr).Select
    DoEvents
    On Error GoTo 0
End Sub

'===============================================================================
' SmoothScrollDown - Scroll down slowly for camera effect
' Parameters:
'   steps     - Number of scroll increments
'   delayMs   - Milliseconds between each scroll (default SCROLL_STEP_DELAY)
'===============================================================================
Private Sub SmoothScrollDown(ByVal steps As Long, Optional ByVal delayMs As Long = 0)
    If delayMs = 0 Then delayMs = SCROLL_STEP_DELAY
    Dim i As Long
    For i = 1 To steps
        ActiveWindow.SmallScroll Down:=3
        DoEvents
        Sleep CLng(delayMs)
    Next i
End Sub

'===============================================================================
' SmoothScrollUp - Scroll up slowly
'===============================================================================
Private Sub SmoothScrollUp(ByVal steps As Long, Optional ByVal delayMs As Long = 0)
    If delayMs = 0 Then delayMs = SCROLL_STEP_DELAY
    Dim i As Long
    For i = 1 To steps
        ActiveWindow.SmallScroll Up:=3
        DoEvents
        Sleep CLng(delayMs)
    Next i
End Sub

'===============================================================================
' ScrollToTop - Reset scroll position to row 1
'===============================================================================
Private Sub ScrollToTop()
    On Error Resume Next
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    DoEvents
    On Error GoTo 0
End Sub

'===============================================================================
' SilencePad - Standard silence padding between clips
'===============================================================================
Private Sub SilencePad()
    WaitSec SILENCE_PAD_SEC
End Sub

'===============================================================================
' ClipTransition - Standard transition between clips (silence + cleanup)
'===============================================================================
Private Sub ClipTransition()
    StopAudio
    SilencePad
    DoEvents
End Sub

'===============================================================================
' SafeDeleteSheet - Delete a sheet by name if it exists (no confirmation dialog)
'===============================================================================
Private Sub SafeDeleteSheet(ByVal shName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(shName)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' SheetExistsLocal - Check if a sheet exists (local helper, no dependency)
'===============================================================================
Private Function SheetExistsLocal(ByVal nm As String) As Boolean
    On Error Resume Next
    SheetExistsLocal = (Not ThisWorkbook.Worksheets(nm) Is Nothing)
    On Error GoTo 0
End Function

'===============================================================================
' CleanupAllOutputSheets - Delete all macro-generated output sheets
' Call this to reset the file to a clean state between videos.
'===============================================================================
Public Sub CleanupAllOutputSheets()
    Application.DisplayAlerts = False

    Dim sheetsToDelete As Variant
    sheetsToDelete = Array( _
        "Data Quality Report", _
        "Variance Analysis", _
        "Variance Commentary", _
        "Executive Dashboard", _
        "YoY Variance Analysis", _
        "Sensitivity Analysis", _
        "Time Saved Analysis", _
        "Executive Brief", _
        "Integration Test Report", _
        "What-If Impact", _
        "Search Results", _
        "WhatIf_Baseline")

    Dim i As Long
    For i = LBound(sheetsToDelete) To UBound(sheetsToDelete)
        SafeDeleteSheet CStr(sheetsToDelete(i))
    Next i

    ' Delete VER_ and BKP_ sheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 4) = "VER_" Or Left(ws.Name, 4) = "BKP_" Then
            ws.Delete
        End If
    Next ws

    Application.DisplayAlerts = True

    ' Clear audit log data (leave headers)
    On Error Resume Next
    Dim wsLog As Worksheet
    Set wsLog = ThisWorkbook.Worksheets("VBA_AuditLog")
    If Not wsLog Is Nothing Then
        wsLog.Visible = xlSheetVisible
        If wsLog.UsedRange.Rows.Count > 1 Then
            wsLog.Rows("2:" & wsLog.UsedRange.Rows.Count).Delete
        End If
        wsLog.Visible = xlSheetVeryHidden
    End If
    On Error GoTo 0

    ' Clear Checks sheet data
    On Error Resume Next
    Dim wsChk As Worksheet
    Set wsChk = ThisWorkbook.Worksheets("Checks")
    If Not wsChk Is Nothing Then
        If wsChk.UsedRange.Rows.Count > 1 Then
            wsChk.Rows("2:" & wsChk.UsedRange.Rows.Count).ClearContents
        End If
    End If
    On Error GoTo 0

    Debug.Print "[Director] All output sheets cleaned up."
End Sub

'===============================================================================
' ShowCommandCenterBriefly - Open the Command Center modeless for camera,
' wait, then close it. Does NOT execute any action.
'===============================================================================
Private Sub ShowCommandCenterBriefly(ByVal showSeconds As Double)
    On Error Resume Next
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmCommandCenter")
    If Not frm Is Nothing Then
        frm.Show vbModeless
        DoEvents
        WaitSec showSeconds
        Unload frm
        DoEvents
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' ShowCCAndSearch - Open CC modeless, type a search term, pause, then close
'===============================================================================
Private Sub ShowCCAndSearch(ByVal searchText As String, ByVal showSeconds As Double)
    On Error Resume Next
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmCommandCenter")
    If Not frm Is Nothing Then
        frm.Show vbModeless
        DoEvents
        WaitSec 1.5   ' let CC render

        ' Type into the search box (txtSearch control)
        On Error Resume Next
        frm.txtSearch.SetFocus
        DoEvents
        SimulateType searchText
        DoEvents
        WaitSec showSeconds  ' pause on filtered results

        ' Clear search
        On Error Resume Next
        frm.txtSearch.Value = ""
        DoEvents
        WaitSec 0.5

        Unload frm
        DoEvents
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' ShowCCTypeActionAndRun - Open CC, type action number, click Run, close
' This gives the visual effect of using the Command Center on camera.
' The actual macro is called AFTER the CC closes for reliability.
'===============================================================================
Private Sub ShowCCTypeActionAndRun(ByVal actionNum As Long, ByVal pauseSec As Double)
    On Error Resume Next
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmCommandCenter")
    If Not frm Is Nothing Then
        frm.Show vbModeless
        DoEvents
        WaitSec 0.8  ' let form render

        ' Type the action number into search box
        frm.txtSearch.SetFocus
        DoEvents
        SimulateType CStr(actionNum)
        DoEvents
        WaitSec pauseSec  ' pause so viewer sees the number

        ' Close the form — we'll run the macro directly
        Unload frm
        DoEvents
    End If
    On Error GoTo 0
End Sub

'===============================================================================
' CheckAbort - Check if user wants to abort (Escape key)
'===============================================================================
Private Function CheckAbort() As Boolean
    DoEvents
    CheckAbort = m_Aborted
End Function

'===============================================================================
' StatusMsg - Show a message in the Excel status bar
'===============================================================================
Private Sub StatusMsg(ByVal msg As String)
    Application.StatusBar = "[Director] " & msg
    DoEvents
End Sub

'===============================================================================
' PreFlightCheck - Verify recording readiness before a video starts.
' Returns True if all checks pass. Shows a single dialog listing all
' failures if any are found.
' Parameters:
'   videoNum  - 1, 2, or 3
'===============================================================================
Private Function PreFlightCheck(ByVal videoNum As Long) As Boolean
    Dim issues As String
    Dim issueCount As Long
    issues = ""
    issueCount = 0

    ' --- Check 1: Excel window maximized ---
    On Error Resume Next
    If Application.WindowState <> xlMaximized Then
        issueCount = issueCount + 1
        issues = issues & issueCount & ". Excel is NOT maximized (fix: maximize the window)" & vbCrLf
        ' Auto-fix
        Application.WindowState = xlMaximized
        issues = issues & "   -> AUTO-FIXED: Window maximized" & vbCrLf
    End If
    On Error GoTo 0

    ' --- Check 2: Zoom is 100% ---
    On Error Resume Next
    If ActiveWindow.Zoom <> 100 Then
        issueCount = issueCount + 1
        issues = issues & issueCount & ". Zoom is " & ActiveWindow.Zoom & "% (should be 100%)" & vbCrLf
        ' Auto-fix
        ActiveWindow.Zoom = 100
        issues = issues & "   -> AUTO-FIXED: Zoom set to 100%" & vbCrLf
    End If
    On Error GoTo 0

    ' --- Check 3: Active sheet is correct ---
    If videoNum <= 2 Then
        If ActiveSheet.Name <> "Report-->" Then
            issueCount = issueCount + 1
            issues = issues & issueCount & ". Active sheet is '" & ActiveSheet.Name & "' (should be 'Report-->')" & vbCrLf
            ' Auto-fix
            GoToSheet "Report-->"
            issues = issues & "   -> AUTO-FIXED: Navigated to Report-->" & vbCrLf
        End If
    End If

    ' --- Check 4: Cell A1 selected ---
    On Error Resume Next
    If ActiveCell.Address <> "$A$1" Then
        issueCount = issueCount + 1
        issues = issues & issueCount & ". Selected cell is " & ActiveCell.Address & " (should be A1)" & vbCrLf
        ' Auto-fix
        SelectCell "A1"
        ScrollToTop
        issues = issues & "   -> AUTO-FIXED: Selected A1 and scrolled to top" & vbCrLf
    End If
    On Error GoTo 0

    ' --- Check 5: Audio files exist ---
    Dim audioFolder As String
    Select Case videoNum
        Case 1: audioFolder = "Video1"
        Case 2: audioFolder = "Video2"
        Case 3: audioFolder = "Video3"
    End Select

    Dim audioPath As String
    audioPath = AUDIO_BASE_PATH & audioFolder & "\"
    If Dir(audioPath, vbDirectory) = "" Then
        issueCount = issueCount + 1
        issues = issues & issueCount & ". Audio folder not found: " & audioPath & vbCrLf
    Else
        ' Count MP3 files
        Dim mp3Count As Long: mp3Count = 0
        Dim f As String: f = Dir(audioPath & "*.mp3")
        Do While f <> ""
            mp3Count = mp3Count + 1
            f = Dir()
        Loop
        If mp3Count = 0 Then
            issueCount = issueCount + 1
            issues = issues & issueCount & ". No MP3 files found in: " & audioPath & vbCrLf
        Else
            Debug.Print "[Director] PreFlight: Found " & mp3Count & " audio clips in " & audioFolder
        End If
    End If

    ' --- Check 6: Required sheets exist (Videos 1 & 2 only) ---
    If videoNum <= 2 Then
        Dim reqSheets As Variant
        reqSheets = Array("Report-->", "P&L - Monthly Trend", "Product Line Summary", _
                          "Assumptions", "General Ledger", "Checks")
        Dim s As Long
        For s = 0 To UBound(reqSheets)
            If Not SheetExistsLocal(CStr(reqSheets(s))) Then
                issueCount = issueCount + 1
                issues = issues & issueCount & ". Required sheet missing: " & reqSheets(s) & vbCrLf
            End If
        Next s
    End If

    ' --- Check 7: MCI audio subsystem works ---
    Dim testRet As Long
    testRet = mciSendStringA("open """ & AUDIO_BASE_PATH & audioFolder & "\" & Dir(AUDIO_BASE_PATH & audioFolder & "\*.mp3") & """ alias director_test", vbNullString, 0, 0)
    If testRet = 0 Then
        ' Subsystem works — clean up
        mciSendStringA "close director_test", vbNullString, 0, 0
        Debug.Print "[Director] PreFlight: MCI audio subsystem OK"
    Else
        issueCount = issueCount + 1
        issues = issues & issueCount & ". MCI audio subsystem error (ret=" & testRet & "). Audio may not play." & vbCrLf
    End If

    ' --- Report ---
    If issueCount = 0 Then
        Debug.Print "[Director] PreFlight PASSED — all checks OK for Video " & videoNum
        PreFlightCheck = True
    Else
        Dim msg As String
        msg = "PRE-FLIGHT CHECK for Video " & videoNum & vbCrLf & vbCrLf & _
              issues & vbCrLf & _
              "Items marked AUTO-FIXED have been corrected." & vbCrLf & _
              "Click OK to proceed, or Cancel to abort."
        If MsgBox(msg, vbOKCancel + vbExclamation, "Director - Pre-Flight") = vbOK Then
            PreFlightCheck = True
        Else
            PreFlightCheck = False
        End If
    End If
End Function

'===============================================================================
' VerifyAudioFiles - Check that all expected MP3 files exist for a video.
' Returns a string listing any missing files (empty = all present).
'===============================================================================
Private Function VerifyAudioFiles(ByVal videoNum As Long) As String
    Dim missing As String: missing = ""
    Dim folder As String
    Dim files As Variant

    Select Case videoNum
        Case 1
            folder = "Video1\"
            files = Array("V1_S1_Opening_Hook.mp3", "V1_S2_Command_Center.mp3", _
                          "V1_S3_Data_Quality.mp3", "V1_S4_Variance_Commentary.mp3", _
                          "V1_S5_Dashboard.mp3", "V1_S6_Bridge.mp3", "V1_S7_Closing.mp3")
        Case 2
            folder = "Video2\"
            files = Array("V2_S0_Opening.mp3", "V2_S1a_Workbook.mp3", _
                          "V2_S1b_CommandCenter.mp3", "V2_S2_GL_Import.mp3", _
                          "V2_S3_Data_Quality.mp3", "V2_S4_Reconciliation.mp3", _
                          "V2_S5_Variance_Analysis.mp3", "V2_S6_Variance_Commentary.mp3", _
                          "V2_S7_YoY_Variance.mp3", "V2_S8_Dashboard_Charts.mp3", _
                          "V2_S9_Executive_Dashboard.mp3", "V2_S10_PDF_Export.mp3", _
                          "V2_S10b_ExecBrief.mp3", "V2_S11_Executive_Mode.mp3", _
                          "V2_S12_Version_Control.mp3", "V2_S13_WhatIf.mp3", _
                          "V2_S13b_Sensitivity.mp3", "V2_S13c_TimeSaved.mp3", _
                          "V2_S14_Integration_Test.mp3", "V2_S15_Audit_Log.mp3", _
                          "V2_S16_Closing.mp3")
        Case 3
            folder = "Video3\"
            files = Array("V3_S0_Opening.mp3", "V3_C1A_DataSanitizer.mp3", _
                          "V3_C1B_Highlights.mp3", "V3_C1C_Comments.mp3", _
                          "V3_C2A_TabOrganizer.mp3", "V3_C2B_ColumnOps.mp3", _
                          "V3_C2C_SheetTools.mp3", "V3_C3A_Compare.mp3", _
                          "V3_C3B_Consolidate.mp3", "V3_C3C_PivotTools.mp3", _
                          "V3_C3D_LookupValidation.mp3", "V3_C4_CommandCenter.mp3", _
                          "V3_Closing.mp3")
        Case Else
            VerifyAudioFiles = "Invalid video number"
            Exit Function
    End Select

    Dim i As Long
    For i = 0 To UBound(files)
        If Dir(AUDIO_BASE_PATH & folder & files(i)) = "" Then
            missing = missing & "  MISSING: " & folder & files(i) & vbCrLf
        End If
    Next i

    VerifyAudioFiles = missing
End Function

'===============================================================================
' RunPreflight - Public entry point: run ALL checks before recording.
' Call this manually before starting OBS to verify everything is ready.
'===============================================================================
Public Sub RunPreflight()
    ResetMCI

    Dim report As String
    report = "=== DIRECTOR PRE-FLIGHT REPORT ===" & vbCrLf & vbCrLf

    ' --- Audio file inventory ---
    Dim v As Long
    For v = 1 To 3
        Dim miss As String
        miss = VerifyAudioFiles(v)
        If miss = "" Then
            report = report & "Video " & v & " audio files: ALL PRESENT" & vbCrLf
        Else
            report = report & "Video " & v & " audio files: MISSING:" & vbCrLf & miss
        End If
    Next v
    report = report & vbCrLf

    ' --- MCI audio test ---
    Dim testFile As String
    testFile = Dir(AUDIO_BASE_PATH & "Video1\*.mp3")
    If testFile <> "" Then
        Dim ret As Long
        ret = mciSendStringA("open """ & AUDIO_BASE_PATH & "Video1\" & testFile & """ alias director_test", vbNullString, 0, 0)
        If ret = 0 Then
            ' Measure duration to prove it works
            mciSendStringA "set director_test time format milliseconds", vbNullString, 0, 0
            Dim buf As String * 128
            mciSendStringA "status director_test length", buf, 128, 0
            Dim testMs As Long: testMs = Val(buf)
            report = report & "MCI audio subsystem: OK (test clip = " & Format(testMs / 1000, "0.0") & "s)" & vbCrLf
            mciSendStringA "close director_test", vbNullString, 0, 0
        Else
            report = report & "MCI audio subsystem: FAILED (error " & ret & ")" & vbCrLf
        End If
    Else
        report = report & "MCI audio test: SKIPPED (no Video1 clips found)" & vbCrLf
    End If

    ' --- Window state ---
    report = report & vbCrLf
    report = report & "Excel maximized: " & IIf(Application.WindowState = xlMaximized, "YES", "NO") & vbCrLf

    On Error Resume Next
    report = report & "Zoom level: " & ActiveWindow.Zoom & "%" & vbCrLf
    report = report & "Active sheet: " & ActiveSheet.Name & vbCrLf
    report = report & "Active cell: " & ActiveCell.Address & vbCrLf
    On Error GoTo 0

    ' --- Required sheets (demo file) ---
    report = report & vbCrLf & "Required sheets:" & vbCrLf
    Dim reqSheets As Variant
    reqSheets = Array("Report-->", "P&L - Monthly Trend", "Product Line Summary", _
                      "Assumptions", "General Ledger", "Checks")
    Dim s As Long
    For s = 0 To UBound(reqSheets)
        report = report & "  " & reqSheets(s) & ": " & _
                 IIf(SheetExistsLocal(CStr(reqSheets(s))), "FOUND", "MISSING") & vbCrLf
    Next s

    ' --- Summary ---
    report = report & vbCrLf & "=== PRE-FLIGHT COMPLETE ==="

    Debug.Print report
    MsgBox report, vbInformation, "Director - Pre-Flight Report"
End Sub

'===============================================================================
'
'    SECTION 2: VIDEO 1 — "WHAT'S POSSIBLE" (Clips 1-7)
'
'===============================================================================

'===============================================================================
' RunVideo1 - Execute the full Video 1 sequence
'===============================================================================
Public Sub RunVideo1()
    On Error GoTo ErrHandler
    ResetMCI
    m_Aborted = False

    ' Pre-flight check
    If Not PreFlightCheck(1) Then
        MsgBox "Pre-flight check failed or cancelled. Video 1 not started.", vbExclamation, "Director"
        Exit Sub
    End If

    StatusMsg "VIDEO 1 starting — What's Possible"
    Debug.Print "========================================"
    Debug.Print "[Director] VIDEO 1 START: " & Now()
    Debug.Print "========================================"

    ' Ensure clean state
    CleanupAllOutputSheets
    GoToSheet "Report-->"
    SelectCell "A1"
    ScrollToTop
    WaitSec 1

    ' --- CLIP 1: Title Card (5 sec) ---
    StatusMsg "V1 Clip 1/7 — Title Card"
    Debug.Print "[Director] Clip 1: Title Card"
    V1_Clip1_TitleCard
    If CheckAbort Then GoTo Aborted

    ' --- CLIP 2: Opening Hook (30 sec) ---
    StatusMsg "V1 Clip 2/7 — Opening Hook"
    Debug.Print "[Director] Clip 2: Opening Hook"
    V1_Clip2_OpeningHook
    If CheckAbort Then GoTo Aborted

    ' --- CLIP 3: Command Center Introduction (40 sec) ---
    StatusMsg "V1 Clip 3/7 — Command Center"
    Debug.Print "[Director] Clip 3: Command Center"
    V1_Clip3_CommandCenter
    If CheckAbort Then GoTo Aborted

    ' --- CLIP 4: Data Quality Scan (40 sec) ---
    StatusMsg "V1 Clip 4/7 — Data Quality Scan"
    Debug.Print "[Director] Clip 4: Data Quality Scan"
    V1_Clip4_DataQuality
    If CheckAbort Then GoTo Aborted

    ' --- CLIP 5: Variance Commentary (45 sec) ---
    StatusMsg "V1 Clip 5/7 — Variance Commentary"
    Debug.Print "[Director] Clip 5: Variance Commentary"
    V1_Clip5_VarianceCommentary
    If CheckAbort Then GoTo Aborted

    ' --- CLIP 6: Executive Dashboard (40 sec) ---
    StatusMsg "V1 Clip 6/7 — Executive Dashboard"
    Debug.Print "[Director] Clip 6: Executive Dashboard"
    V1_Clip6_Dashboard
    If CheckAbort Then GoTo Aborted

    ' --- CLIP 7: Bridge + Closing (60 sec) ---
    StatusMsg "V1 Clip 7/7 — Bridge + Closing"
    Debug.Print "[Director] Clip 7: Bridge + Closing"
    V1_Clip7_BridgeClosing

    StatusMsg "VIDEO 1 COMPLETE"
    Debug.Print "========================================"
    Debug.Print "[Director] VIDEO 1 COMPLETE: " & Now()
    Debug.Print "========================================"

    StopAudio
    Application.StatusBar = False
    MsgBox "Video 1 recording complete!" & vbCrLf & vbCrLf & _
           "Stop OBS recording now." & vbCrLf & _
           "Then run CleanupAllOutputSheets before Video 2.", _
           vbInformation, "Director - Video 1 Done"
    Exit Sub

Aborted:
    StopAudio
    Application.StatusBar = False
    Debug.Print "[Director] VIDEO 1 ABORTED by user."
    MsgBox "Video 1 aborted.", vbExclamation, "Director"
    Exit Sub

ErrHandler:
    StopAudio
    Application.StatusBar = False
    Debug.Print "[Director] ERROR in Video 1: " & Err.Description
    MsgBox "Director Error: " & Err.Description, vbCritical, "Director"
End Sub

'===============================================================================
' CLIP 1 — Title Card (5 seconds)
' No audio. Just a static pause. Title card is added in video editor.
'===============================================================================
Private Sub V1_Clip1_TitleCard()
    ' The title card gets overlaid in editing.
    ' Just hold still on the Report--> page for 5 seconds.
    GoToSheet "Report-->"
    SelectCell "A1"
    ScrollToTop
    WaitSec 5
End Sub

'===============================================================================
' CLIP 2 — Opening Hook: Landing Page (~30 sec)
' Audio: V1_S1_Opening_Hook.mp3
' Action: Slow scroll down the Report--> landing page
'===============================================================================
Private Sub V1_Clip2_OpeningHook()
    GoToSheet "Report-->"
    SelectCell "A1"
    ScrollToTop

    ' Silence lead-in
    SilencePad

    ' Play narration: "This is a single Excel file..."
    PlayAudio AUDIO_BASE_PATH & "Video1\V1_S1_Opening_Hook.mp3"

    ' Wait 2 sec for first sentence, then start scrolling
    WaitSec 2

    ' Slowly scroll down through the landing page (~12 seconds of scrolling)
    SmoothScrollDown 15, 500

    ' Pause on "62 automated actions" — hold still
    WaitSec 3

    ' Continue gentle scroll
    SmoothScrollDown 5, 400

    ' Hold still for "show you what that looks like"
    WaitSec m_ClipDurSec - 12   ' remaining audio time

    ' Silence tail
    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 3 — Command Center Introduction (~40 sec)
' Audio: V1_S2_Command_Center.mp3
' Action: Open CC, scroll categories, search "variance", close CC
'===============================================================================
Private Sub V1_Clip3_CommandCenter()
    GoToSheet "Report-->"
    ScrollToTop

    SilencePad

    ' Play narration: "Everything runs from one place..."
    PlayAudio AUDIO_BASE_PATH & "Video1\V1_S2_Command_Center.mp3"

    ' Wait for "the Command Center" line (~3 sec)
    WaitSec 3

    ' Open the Command Center modeless for camera
    On Error Resume Next
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmCommandCenter")
    If Not frm Is Nothing Then
        frm.Show vbModeless
        DoEvents

        ' Let viewer absorb the CC (1.5 sec)
        WaitSec 1.5

        ' Scroll through categories in the listbox (~6 sec)
        Dim catIdx As Long
        On Error Resume Next
        For catIdx = 0 To frm.lstCategories.ListCount - 1
            frm.lstCategories.ListIndex = catIdx
            DoEvents
            WaitSec 0.8
        Next catIdx
        On Error GoTo 0

        ' Now type "variance" in search box (~5 sec)
        WaitSec 1
        On Error Resume Next
        frm.txtSearch.SetFocus
        DoEvents
        On Error GoTo 0
        SimulateType "variance", 120
        DoEvents

        ' Pause on filtered results
        WaitSec 3

        ' Clear search
        On Error Resume Next
        frm.txtSearch.Value = ""
        DoEvents
        On Error GoTo 0
        WaitSec 1

        ' Close the Command Center
        Unload frm
        DoEvents
    Else
        ' Fallback: CC form doesn't exist — just wait for audio
        WaitSec m_ClipDurSec
    End If
    On Error GoTo 0

    ' Wait for remaining audio
    WaitSec 2

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 4 — Data Quality Scan + Letter Grade (~40 sec)
' Audio: V1_S3_Data_Quality.mp3
' Action: Open CC briefly, run ScanAll, scroll report, hold on letter grade
'===============================================================================
Private Sub V1_Clip4_DataQuality()
    SilencePad

    ' Play narration: "First — data quality..."
    PlayAudio AUDIO_BASE_PATH & "Video1\V1_S3_Data_Quality.mp3"

    ' Wait for "Before you do anything" (~3 sec)
    WaitSec 3

    ' Show CC briefly with action number for camera
    ShowCCTypeActionAndRun 7, 1.5

    ' Run the actual Data Quality Scan
    On Error Resume Next
    modDataQuality.ScanAll
    On Error GoTo 0
    DoEvents

    ' The macro navigates to "Data Quality Report" sheet
    ' Pause on the letter grade badge (2-3 sec)
    WaitSec 3

    ' Slowly scroll down through the category breakdown
    SmoothScrollDown 8, 400

    ' Hold still for "Fifteen seconds, start to finish"
    WaitSec 4

    SilencePad
    StopAudio

    ' RESET: Delete the output sheet for next clip
    SafeDeleteSheet "Data Quality Report"
    WaitSec 0.5
End Sub

'===============================================================================
' CLIP 5 — Variance Commentary — JAW-DROP Feature (~45 sec)
' Audio: V1_S4_Variance_Commentary.mp3
' Action: Run GenerateCommentary, pause to let viewer read, scroll narratives
'===============================================================================
Private Sub V1_Clip5_VarianceCommentary()
    SilencePad

    ' Play narration: "Next — one of the most useful features..."
    PlayAudio AUDIO_BASE_PATH & "Video1\V1_S4_Variance_Commentary.mp3"

    ' Wait for intro line (~3 sec)
    WaitSec 3

    ' Show CC briefly
    ShowCCTypeActionAndRun 46, 1.5

    ' Run Variance Commentary
    On Error Resume Next
    modVarianceAnalysis.GenerateCommentary
    On Error GoTo 0
    DoEvents

    ' JAW-DROP MOMENT: Pause 3 seconds in silence — let viewer READ
    WaitSec 3

    ' Slowly scroll through the narratives
    SmoothScrollDown 6, 500

    ' Hover near a narrative (just select a cell in the middle)
    On Error Resume Next
    ActiveSheet.Range("B8").Select
    On Error GoTo 0
    DoEvents

    ' Hold still for "One click" line
    WaitSec 4

    SilencePad
    StopAudio

    ' RESET: Delete output sheet
    SafeDeleteSheet "Variance Commentary"
    WaitSec 0.5
End Sub

'===============================================================================
' CLIP 6 — Executive Dashboard (~40 sec)
' Audio: V1_S5_Dashboard.mp3
' Action: Run CreateExecutiveDashboard, scroll through KPIs/waterfall/products
'===============================================================================
Private Sub V1_Clip6_Dashboard()
    SilencePad

    ' Play narration: "When it's time to present to leadership..."
    PlayAudio AUDIO_BASE_PATH & "Video1\V1_S5_Dashboard.mp3"

    ' Wait for "One click builds" (~4 sec)
    WaitSec 4

    ' Show CC briefly
    ShowCCTypeActionAndRun 12, 1

    ' Run Executive Dashboard (takes 5-10 seconds)
    On Error Resume Next
    modDashboardAdvanced.CreateExecutiveDashboard
    On Error GoTo 0
    DoEvents

    ' Dashboard appears — pause 2 sec to let it register
    WaitSec 2

    ' KPI cards at top — pause
    WaitSec 2

    ' Scroll to waterfall chart
    SmoothScrollDown 6, 400
    WaitSec 2   ' hold on waterfall

    ' Scroll to product comparison
    SmoothScrollDown 6, 400
    WaitSec 2   ' hold on product comparison

    ' Hold still for remaining audio (PDF mention)
    WaitSec 3

    SilencePad
    StopAudio

    ' RESET: Delete output sheet
    SafeDeleteSheet "Executive Dashboard"
    WaitSec 0.5
End Sub

'===============================================================================
' CLIP 7 — Bridge to Universal Tools + Closing (~60 sec)
' Audio: V1_S6_Bridge.mp3 then V1_S7_Closing.mp3
' Action: Navigate back to Report--> page, hold static
'===============================================================================
Private Sub V1_Clip7_BridgeClosing()
    ' Navigate to Report--> landing page
    GoToSheet "Report-->"
    SelectCell "A1"
    ScrollToTop

    SilencePad

    ' Play Bridge narration: "That's a sample of what this file can do..."
    PlayClip "Video1", "V1_S6_Bridge.mp3"

    ' Brief pause between clips
    WaitSec 1

    ' Play Closing narration: "Everything you just saw runs from this one Excel file..."
    PlayClip "Video1", "V1_S7_Closing.mp3"

    ' Hold still for "Thanks for watching" — 3 seconds
    WaitSec 3

    SilencePad
    StopAudio
End Sub

'===============================================================================
'
'    SECTION 3: VIDEO 2 — "FULL DEMO WALKTHROUGH" (Clips 8-26)
'
'===============================================================================

'===============================================================================
' RunVideo2 - Execute the full Video 2 sequence
'===============================================================================
Public Sub RunVideo2()
    On Error GoTo ErrHandler
    ResetMCI
    m_Aborted = False

    ' Pre-flight check
    If Not PreFlightCheck(2) Then
        MsgBox "Pre-flight check failed or cancelled. Video 2 not started.", vbExclamation, "Director"
        Exit Sub
    End If

    StatusMsg "VIDEO 2 starting — Full Demo Walkthrough"
    Debug.Print "========================================"
    Debug.Print "[Director] VIDEO 2 START: " & Now()
    Debug.Print "========================================"

    ' Ensure clean state
    CleanupAllOutputSheets
    GoToSheet "Report-->"
    SelectCell "A1"
    ScrollToTop
    WaitSec 1

    ' --- Clip 8: Opening ---
    StatusMsg "V2 Clip 8 — Opening"
    V2_Clip8_Opening
    If CheckAbort Then GoTo Aborted

    ' --- Clip 9: Workbook Tour ---
    StatusMsg "V2 Clip 9 — Workbook Tour"
    V2_Clip9_WorkbookTour
    If CheckAbort Then GoTo Aborted

    ' --- Clip 10: Command Center ---
    StatusMsg "V2 Clip 10 — Command Center"
    V2_Clip10_CommandCenter
    If CheckAbort Then GoTo Aborted

    ' --- Clip 11: GL Import ---
    StatusMsg "V2 Clip 11 — GL Import"
    V2_Clip11_GLImport
    If CheckAbort Then GoTo Aborted

    ' --- Clip 12: Data Quality ---
    StatusMsg "V2 Clip 12 — Data Quality"
    V2_Clip12_DataQuality
    If CheckAbort Then GoTo Aborted

    ' --- Clip 13: Reconciliation ---
    StatusMsg "V2 Clip 13 — Reconciliation"
    V2_Clip13_Reconciliation
    If CheckAbort Then GoTo Aborted

    ' --- Clip 14: Variance Analysis ---
    StatusMsg "V2 Clip 14 — Variance Analysis"
    V2_Clip14_VarianceAnalysis
    If CheckAbort Then GoTo Aborted

    ' --- Clip 15: Variance Commentary ---
    StatusMsg "V2 Clip 15 — Variance Commentary"
    V2_Clip15_VarianceCommentary
    If CheckAbort Then GoTo Aborted

    ' --- Clip 16: YoY Variance ---
    StatusMsg "V2 Clip 16 — YoY Variance"
    V2_Clip16_YoYVariance
    If CheckAbort Then GoTo Aborted

    ' --- Clip 17: Dashboard Charts ---
    StatusMsg "V2 Clip 17 — Dashboard Charts"
    V2_Clip17_DashboardCharts
    If CheckAbort Then GoTo Aborted

    ' --- Clip 18: Executive Dashboard ---
    StatusMsg "V2 Clip 18 — Executive Dashboard"
    V2_Clip18_ExecDashboard
    If CheckAbort Then GoTo Aborted

    ' --- Clip 19: PDF Export ---
    StatusMsg "V2 Clip 19 — PDF Export"
    V2_Clip19_PDFExport
    If CheckAbort Then GoTo Aborted

    ' --- Clip 20: Executive Brief ---
    StatusMsg "V2 Clip 20 — Executive Brief"
    V2_Clip20_ExecBrief
    If CheckAbort Then GoTo Aborted

    ' --- Clip 21: Executive Mode ---
    StatusMsg "V2 Clip 21 — Executive Mode"
    V2_Clip21_ExecutiveMode
    If CheckAbort Then GoTo Aborted

    ' --- Clip 22: Version Control ---
    StatusMsg "V2 Clip 22 — Version Control"
    V2_Clip22_VersionControl
    If CheckAbort Then GoTo Aborted

    ' --- Clip 23: What-If Scenario ---
    StatusMsg "V2 Clip 23 — What-If Scenario"
    V2_Clip23_WhatIf
    If CheckAbort Then GoTo Aborted

    ' --- Clip 24: Sensitivity Analysis ---
    StatusMsg "V2 Clip 24 — Sensitivity Analysis"
    V2_Clip24_Sensitivity
    If CheckAbort Then GoTo Aborted

    ' --- Clip 25: Integration Test ---
    StatusMsg "V2 Clip 25 — Integration Test"
    V2_Clip25_IntegrationTest
    If CheckAbort Then GoTo Aborted

    ' --- Clip 26: Audit Log + Closing ---
    StatusMsg "V2 Clip 26 — Audit Log + Closing"
    V2_Clip26_AuditLogClosing

    StatusMsg "VIDEO 2 COMPLETE"
    Debug.Print "========================================"
    Debug.Print "[Director] VIDEO 2 COMPLETE: " & Now()
    Debug.Print "========================================"

    StopAudio
    Application.StatusBar = False
    MsgBox "Video 2 recording complete!" & vbCrLf & vbCrLf & _
           "Stop OBS recording now." & vbCrLf & _
           "Then run CleanupAllOutputSheets before Video 3.", _
           vbInformation, "Director - Video 2 Done"
    Exit Sub

Aborted:
    StopAudio
    Application.StatusBar = False
    MsgBox "Video 2 aborted.", vbExclamation, "Director"
    Exit Sub

ErrHandler:
    StopAudio
    Application.StatusBar = False
    MsgBox "Director Error in Video 2: " & Err.Description, vbCritical, "Director"
End Sub

'===============================================================================
' CLIP 8 — Opening (~40 sec)
' Audio: V2_S0_Opening.mp3
' Action: Slow scroll on Report--> page
'===============================================================================
Private Sub V2_Clip8_Opening()
    GoToSheet "Report-->"
    SelectCell "A1"
    ScrollToTop

    SilencePad

    ' "Welcome to the full walkthrough..."
    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S0_Opening.mp3"

    WaitSec 3
    SmoothScrollDown 12, 500
    WaitSec m_ClipDurSec - 9

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 9 — Workbook Tour (~85 sec)
' Audio: V2_S1a_Workbook.mp3
' Action: Click through sheet tabs to show workbook structure
'===============================================================================
Private Sub V2_Clip9_WorkbookTour()
    GoToSheet "Report-->"
    ScrollToTop

    SilencePad

    ' "Let's start with what's inside this file..."
    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S1a_Workbook.mp3"

    WaitSec 4   ' let intro line play

    ' Click through sheet tabs with pauses
    GoToSheet "P&L - Monthly Trend"
    ScrollToTop
    WaitSec 3   ' pause on P&L Trend

    ' Click to a Functional P&L Summary tab (fall back to Jan 25 if missing)
    If SheetExistsLocal("Functional P&L Summary") Then
        GoToSheet "Functional P&L Summary"
    Else
        GoToSheet "Jan 25"
    End If
    WaitSec 2

    GoToSheet "Product Line Summary"
    ScrollToTop
    WaitSec 3

    GoToSheet "Assumptions"
    ScrollToTop
    WaitSec 2

    GoToSheet "General Ledger"
    ScrollToTop
    SmoothScrollDown 3, 300
    WaitSec 2

    ' Wait for remaining audio
    WaitSec m_ClipDurSec - 20

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 10 — Command Center Overview (~45 sec)
' Audio: V2_S1b_CommandCenter.mp3
' Action: Open CC, scroll categories, search "reconciliation", close
'===============================================================================
Private Sub V2_Clip10_CommandCenter()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S1b_CommandCenter.mp3"

    WaitSec 2

    ' Open CC and browse it
    ShowCCAndSearch "reconciliation", 3

    WaitSec m_ClipDurSec - 10

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 11 — GL Import (~45 sec)
' Audio: V2_S2_GL_Import.mp3
' Action: Show CC, then navigate to General Ledger to show data
' NOTE: We skip the actual file dialog import to avoid UI blocking.
'       Instead we show the GL sheet which already has data loaded.
'===============================================================================
Private Sub V2_Clip11_GLImport()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S2_GL_Import.mp3"

    WaitSec 3

    ' Show CC briefly with action 17
    ShowCCTypeActionAndRun 17, 1.5

    ' Navigate to GL sheet to show data (skip actual import dialog)
    GoToSheet "General Ledger"
    ScrollToTop
    WaitSec 2

    ' Scroll through GL data
    SmoothScrollDown 8, 350

    WaitSec m_ClipDurSec - 10

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 12 — Data Quality Scan (~50 sec)
' Audio: V2_S3_Data_Quality.mp3
' Action: Run ScanAll, show report with letter grade
'===============================================================================
Private Sub V2_Clip12_DataQuality()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S3_Data_Quality.mp3"

    WaitSec 3

    ShowCCTypeActionAndRun 7, 1.5

    ' Run scan
    On Error Resume Next
    modDataQuality.ScanAll
    On Error GoTo 0
    DoEvents

    ' Pause on letter grade
    WaitSec 3

    ' Scroll through category breakdown
    SmoothScrollDown 10, 400

    WaitSec m_ClipDurSec - 12

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 13 — Reconciliation Checks (~45 sec)
' Audio: V2_S4_Reconciliation.mp3
' Action: Run RunAllChecks, show PASS/FAIL results on Checks sheet
'===============================================================================
Private Sub V2_Clip13_Reconciliation()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S4_Reconciliation.mp3"

    WaitSec 3

    ShowCCTypeActionAndRun 3, 1.5

    ' Run reconciliation checks
    On Error Resume Next
    modReconciliation.RunAllChecks
    On Error GoTo 0
    DoEvents

    ' Show Checks sheet with PASS/FAIL
    GoToSheet "Checks"
    ScrollToTop
    WaitSec 3

    SmoothScrollDown 6, 400

    WaitSec m_ClipDurSec - 10

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 14 — Variance Analysis (~40 sec)
' Audio: V2_S5_Variance_Analysis.mp3
' Action: Run RunVarianceAnalysis, scroll flagged items
'===============================================================================
Private Sub V2_Clip14_VarianceAnalysis()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S5_Variance_Analysis.mp3"

    WaitSec 3

    ShowCCTypeActionAndRun 6, 1.5

    On Error Resume Next
    modVarianceAnalysis.RunVarianceAnalysis
    On Error GoTo 0
    DoEvents

    ' Pause on headers
    WaitSec 2

    ' Scroll through flagged items
    SmoothScrollDown 8, 400

    WaitSec m_ClipDurSec - 10

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 15 — Variance Commentary — JAW-DROP (~45 sec)
' Audio: V2_S6_Variance_Commentary.mp3
' Action: Run GenerateCommentary, silent pause, scroll narratives
'===============================================================================
Private Sub V2_Clip15_VarianceCommentary()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S6_Variance_Commentary.mp3"

    WaitSec 3

    ShowCCTypeActionAndRun 46, 1.5

    On Error Resume Next
    modVarianceAnalysis.GenerateCommentary
    On Error GoTo 0
    DoEvents

    ' JAW-DROP: Silent pause (let viewer read)
    WaitSec 3

    ' Scroll narratives slowly
    SmoothScrollDown 6, 600

    ' Select a narrative cell to draw eye
    On Error Resume Next
    ActiveSheet.Range("B6").Select
    On Error GoTo 0

    WaitSec m_ClipDurSec - 12

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 16 — YoY Variance (~30 sec)
' Audio: V2_S7_YoY_Variance.mp3
' Action: Run RunYoYVarianceAnalysis, scroll results
'===============================================================================
Private Sub V2_Clip16_YoYVariance()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S7_YoY_Variance.mp3"

    WaitSec 3

    ShowCCTypeActionAndRun 47, 1.5

    On Error Resume Next
    modVarianceAnalysis.RunYoYVarianceAnalysis
    On Error GoTo 0
    DoEvents

    WaitSec 2
    SmoothScrollDown 8, 350

    WaitSec m_ClipDurSec - 10

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 17 — Dashboard Charts (~45 sec)
' Audio: V2_S8_Dashboard_Charts.mp3
' Action: Run BuildDashboard, scroll through chart grid
'===============================================================================
Private Sub V2_Clip17_DashboardCharts()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S8_Dashboard_Charts.mp3"

    WaitSec 3

    ShowCCTypeActionAndRun 12, 1.5

    On Error Resume Next
    modDashboard.BuildDashboard
    On Error GoTo 0
    DoEvents

    ' Pause on chart grid
    WaitSec 3

    ' Scroll through charts
    SmoothScrollDown 10, 400

    WaitSec m_ClipDurSec - 12

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 18 — Executive Dashboard (~30 sec)
' Audio: V2_S9_Executive_Dashboard.mp3
' Action: Run CreateExecutiveDashboard, scroll KPIs/waterfall/products
'===============================================================================
Private Sub V2_Clip18_ExecDashboard()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S9_Executive_Dashboard.mp3"

    WaitSec 3

    On Error Resume Next
    modDashboardAdvanced.CreateExecutiveDashboard
    On Error GoTo 0
    DoEvents

    ' KPI cards — hold
    WaitSec 2

    ' Waterfall chart
    SmoothScrollDown 6, 350
    WaitSec 2

    ' Product comparison
    SmoothScrollDown 6, 350
    WaitSec 2

    WaitSec m_ClipDurSec - 12

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 19 — PDF Export (~30 sec)
' Audio: V2_S10_PDF_Export.mp3
' Action: Export report sheets to PDF
' v2.0: BYPASSES the file dialog entirely — exports directly to Desktop.
'       This eliminates the fragile SendKeys workaround for the SaveAs dialog.
'===============================================================================
Private Sub V2_Clip19_PDFExport()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S10_PDF_Export.mp3"

    WaitSec 3

    ShowCCTypeActionAndRun 10, 1.5

    ' Direct PDF export — bypasses GetSaveAsFilename dialog entirely
    On Error Resume Next

    Dim pdfPath As String
    pdfPath = Environ("USERPROFILE") & "\Desktop\KBT_Report_Package_" & Format(Now, "yyyymmdd") & ".pdf"

    ' Collect valid report sheets
    Dim candidates As Variant
    candidates = Array("Report-->", "P&L - Monthly Trend", "Product Line Summary", _
                       "Functional P&L Summary", "Assumptions", "General Ledger", "Checks")
    Dim valid() As String
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = 0 To UBound(candidates)
        If SheetExistsLocal(CStr(candidates(i))) Then
            ReDim Preserve valid(cnt)
            valid(cnt) = CStr(candidates(i))
            cnt = cnt + 1
        End If
    Next i

    If cnt > 0 Then
        ' Select all valid sheets for combined export
        ThisWorkbook.Worksheets(valid(0)).Select
        For i = 1 To UBound(valid)
            ThisWorkbook.Worksheets(valid(i)).Select Replace:=False
        Next i

        ' Export selected sheets to PDF
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
            Quality:=xlQualityStandard, IncludeDocProperties:=True

        Debug.Print "[Director] PDF exported to: " & pdfPath
    End If
    On Error GoTo 0
    DoEvents

    ' Re-select single sheet for clean view
    GoToSheet "Report-->"

    WaitSec m_ClipDurSec - 6

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 20 — Executive Brief (~40 sec)
' Audio: V2_S10b_ExecBrief.mp3
' Action: Run GenerateExecBrief, scroll through brief
'===============================================================================
Private Sub V2_Clip20_ExecBrief()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S10b_ExecBrief.mp3"

    WaitSec 3

    On Error Resume Next
    modExecBrief.GenerateExecBrief
    On Error GoTo 0
    DoEvents

    WaitSec 2
    SmoothScrollDown 8, 400

    WaitSec m_ClipDurSec - 8

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 21 — Executive Mode Toggle (~20 sec)
' Audio: V2_S11_Executive_Mode.mp3
' Action: Toggle Executive Mode ON (tabs disappear), then OFF (tabs return)
'===============================================================================
Private Sub V2_Clip21_ExecutiveMode()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S11_Executive_Mode.mp3"

    WaitSec 3

    ' Toggle ON — sheets hide
    On Error Resume Next
    modNavigation.ToggleExecutiveMode
    On Error GoTo 0
    DoEvents

    ' Pause — let viewer see clean tab bar
    WaitSec 3

    ' Toggle OFF — sheets return
    On Error Resume Next
    modNavigation.ToggleExecutiveMode
    On Error GoTo 0
    DoEvents

    WaitSec m_ClipDurSec - 8

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 22 — Version Control (~25 sec)
' Audio: V2_S12_Version_Control.mp3
' Action: Save a version snapshot with a name
' v2.0: BYPASSES the InputBox by calling SaveCopyAs directly.
'       This eliminates the fragile SendKeys workaround.
'===============================================================================
Private Sub V2_Clip22_VersionControl()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S12_Version_Control.mp3"

    WaitSec 3

    ShowCCTypeActionAndRun 32, 1.5

    ' Direct version save — bypasses InputBox entirely
    ' Replicates the core logic of modVersionControl.SaveVersion
    On Error Resume Next

    Dim ts As String: ts = Format(Now, "yyyymmdd_hhmmss")
    Dim basePath As String: basePath = ThisWorkbook.Path
    If basePath = "" Then basePath = Environ("USERPROFILE") & "\Desktop"

    ' Create versions folder if needed
    Dim versDir As String: versDir = basePath & "\versions\"
    If Dir(versDir, vbDirectory) = "" Then MkDir versDir

    ' Build version filename matching modVersionControl format
    Dim vFile As String
    vFile = versDir & "v1_" & ts & "_March_Draft_1.xlsx"

    ' Save a copy
    ThisWorkbook.SaveCopyAs vFile
    Debug.Print "[Director] Version saved: " & vFile

    ' Show status message (what the viewer sees)
    Application.StatusBar = "Version saved: March Draft 1"
    DoEvents

    On Error GoTo 0

    WaitSec m_ClipDurSec - 5

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 23 — What-If Scenario Demo — THE WOW MOMENT (~90 sec)
' Audio: V2_S13_WhatIf.mp3
' Action: Run What-If preset #1 (Revenue drops 15%), show impact, restore
' v2.0: Calls RunWhatIfPreset(1) and RestoreBaselineSilent() directly —
'       NO SendKeys, NO InputBox, NO MsgBox. Zero dialog risk.
'       These are public wrappers added to modWhatIf v2.1.1 specifically
'       for unattended Director playback.
'===============================================================================
Private Sub V2_Clip23_WhatIf()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S13_WhatIf.mp3"

    WaitSec 4

    ShowCCTypeActionAndRun 63, 1.5

    ' Run preset #1 (Revenue Drops 15%) — no InputBox
    On Error Resume Next
    modWhatIf.RunWhatIfPreset 1
    On Error GoTo 0
    DoEvents

    ' What-If Impact sheet appears — pause
    WaitSec 4

    ' Scroll through the impact analysis
    SmoothScrollDown 8, 500
    WaitSec 3

    ' Navigate to Assumptions to show changed values
    GoToSheet "Assumptions"
    ScrollToTop
    WaitSec 3
    SmoothScrollDown 4, 400
    WaitSec 2

    ' Go back to impact sheet
    GoToSheet "What-If Impact"
    ScrollToTop
    WaitSec 2

    ' Restore baseline — no confirmation MsgBox
    On Error Resume Next
    modWhatIf.RestoreBaselineSilent
    On Error GoTo 0
    DoEvents

    WaitSec m_ClipDurSec - 30

    SilencePad
    StopAudio

    ' Clean up (RestoreBaselineSilent already deletes these, but safe to double-check)
    SafeDeleteSheet "What-If Impact"
    SafeDeleteSheet "WhatIf_Baseline"
End Sub

'===============================================================================
' CLIP 24 — Sensitivity Analysis (~35 sec)
' Audio: V2_S13b_Sensitivity.mp3
' Action: Run RunSensitivityAnalysis, scroll results
'===============================================================================
Private Sub V2_Clip24_Sensitivity()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S13b_Sensitivity.mp3"

    WaitSec 3

    ShowCCTypeActionAndRun 5, 1.5

    On Error Resume Next
    modSensitivity.RunSensitivityAnalysis
    On Error GoTo 0
    DoEvents

    WaitSec 2
    SmoothScrollDown 8, 400

    WaitSec m_ClipDurSec - 10

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 25 — Integration Test (~30 sec)
' Audio: V2_S14_Integration_Test.mp3
' Action: Run RunFullTest, show 18/18 PASS results
'===============================================================================
Private Sub V2_Clip25_IntegrationTest()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S14_Integration_Test.mp3"

    WaitSec 3

    ShowCCTypeActionAndRun 44, 1.5

    On Error Resume Next
    modIntegrationTest.RunFullTest
    On Error GoTo 0
    DoEvents

    ' Pause on 18/18 PASS result
    WaitSec 3

    SmoothScrollDown 6, 350

    WaitSec m_ClipDurSec - 10

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 26 — Audit Log + Time Saved + Closing
' Audio: V2_S15_Audit_Log.mp3, V2_S13c_TimeSaved.mp3, V2_S16_Closing.mp3
' Action: Show audit log, show time saved, navigate to Report--> for closing
'===============================================================================
Private Sub V2_Clip26_AuditLogClosing()
    ' --- Part A: Audit Log ---
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S15_Audit_Log.mp3"

    WaitSec 2

    ' Show the audit log
    On Error Resume Next
    modLogger.ViewLog
    On Error GoTo 0
    DoEvents

    WaitSec 2
    SmoothScrollDown 6, 350

    WaitSec m_ClipDurSec - 8
    StopAudio
    WaitSec 1

    ' --- Part B: Time Saved Calculator ---
    PlayAudio AUDIO_BASE_PATH & "Video2\V2_S13c_TimeSaved.mp3"

    WaitSec 2

    On Error Resume Next
    modTimeSaved.ShowTimeSavedReport
    On Error GoTo 0
    DoEvents

    WaitSec 2
    SmoothScrollDown 8, 400

    WaitSec m_ClipDurSec - 8
    StopAudio
    WaitSec 1

    ' --- Part C: Closing ---
    GoToSheet "Report-->"
    SelectCell "A1"
    ScrollToTop

    PlayClip "Video2", "V2_S16_Closing.mp3"

    WaitSec 3

    SilencePad
    StopAudio
End Sub

'===============================================================================
'
'    SECTION 4: VIDEO 3 — "UNIVERSAL TOOLS" (Clips 27-39)
'    NOTE: Video 3 runs on the Sample_Quarterly_Report.xlsm file,
'    NOT the demo file. The user must switch files before running.
'
'===============================================================================

'===============================================================================
' RunVideo3 - Execute the full Video 3 sequence
' IMPORTANT: Run this on the Sample_Quarterly_Report.xlsm file,
'            NOT the demo file. The universal toolkit modules must be
'            imported into the sample file first.
'===============================================================================
Public Sub RunVideo3()
    On Error GoTo ErrHandler
    ResetMCI
    m_Aborted = False

    ' Verify we're not on the demo file
    If InStr(1, ThisWorkbook.Name, "Demo", vbTextCompare) > 0 Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("This appears to be the DEMO file." & vbCrLf & _
                       "Video 3 should run on Sample_Quarterly_Report.xlsm." & vbCrLf & vbCrLf & _
                       "Continue anyway?", vbYesNo + vbExclamation, "Director")
        If resp = vbNo Then Exit Sub
    End If

    ' Pre-flight check (lighter — Video 3 has different sheet requirements)
    If Not PreFlightCheck(3) Then
        MsgBox "Pre-flight check failed or cancelled. Video 3 not started.", vbExclamation, "Director"
        Exit Sub
    End If

    StatusMsg "VIDEO 3 starting — Universal Tools"
    Debug.Print "========================================"
    Debug.Print "[Director] VIDEO 3 START: " & Now()
    Debug.Print "========================================"

    ' --- Clip 27: Opening ---
    StatusMsg "V3 Clip 27 — Opening"
    V3_Clip27_Opening
    If CheckAbort Then GoTo Aborted

    ' --- Clip 28: Data Sanitizer ---
    StatusMsg "V3 Clip 28 — Data Sanitizer"
    V3_Clip28_DataSanitizer
    If CheckAbort Then GoTo Aborted

    ' --- Clip 29: Highlights ---
    StatusMsg "V3 Clip 29 — Highlights"
    V3_Clip29_Highlights
    If CheckAbort Then GoTo Aborted

    ' --- Clip 30: Comments ---
    StatusMsg "V3 Clip 30 — Comments"
    V3_Clip30_Comments
    If CheckAbort Then GoTo Aborted

    ' --- Clip 31: Tab Organizer ---
    StatusMsg "V3 Clip 31 — Tab Organizer"
    V3_Clip31_TabOrganizer
    If CheckAbort Then GoTo Aborted

    ' --- Clip 32: Column Ops ---
    StatusMsg "V3 Clip 32 — Column Ops"
    V3_Clip32_ColumnOps
    If CheckAbort Then GoTo Aborted

    ' --- Clip 33: Sheet Tools ---
    StatusMsg "V3 Clip 33 — Sheet Tools"
    V3_Clip33_SheetTools
    If CheckAbort Then GoTo Aborted

    ' --- Clip 34: Compare ---
    StatusMsg "V3 Clip 34 — Compare Sheets"
    V3_Clip34_Compare
    If CheckAbort Then GoTo Aborted

    ' --- Clip 35: Consolidate ---
    StatusMsg "V3 Clip 35 — Consolidate"
    V3_Clip35_Consolidate
    If CheckAbort Then GoTo Aborted

    ' --- Clip 36: Pivot Tools + Lookup/Validation ---
    StatusMsg "V3 Clip 36 — Pivot & Lookup Tools"
    V3_Clip36_PivotLookup
    If CheckAbort Then GoTo Aborted

    ' --- Clip 37: Universal Command Center ---
    StatusMsg "V3 Clip 37 — Universal Command Center"
    V3_Clip37_CommandCenter
    If CheckAbort Then GoTo Aborted

    ' --- Clip 38: Closing ---
    StatusMsg "V3 Clip 38 — Closing"
    V3_Clip38_Closing

    StatusMsg "VIDEO 3 COMPLETE"
    Debug.Print "========================================"
    Debug.Print "[Director] VIDEO 3 COMPLETE: " & Now()
    Debug.Print "========================================"

    StopAudio
    Application.StatusBar = False
    MsgBox "Video 3 recording complete!" & vbCrLf & vbCrLf & _
           "Stop OBS recording now.", _
           vbInformation, "Director - Video 3 Done"
    Exit Sub

Aborted:
    StopAudio
    Application.StatusBar = False
    MsgBox "Video 3 aborted.", vbExclamation, "Director"
    Exit Sub

ErrHandler:
    StopAudio
    Application.StatusBar = False
    MsgBox "Director Error in Video 3: " & Err.Description, vbCritical, "Director"
End Sub

'===============================================================================
' CLIP 27 — Opening (~45 sec)
' Audio: V3_S0_Opening.mp3
' Action: Show messy sample file, slow scroll
'===============================================================================
Private Sub V3_Clip27_Opening()
    ' Should already be on the sample file's first sheet
    On Error Resume Next
    ActiveWorkbook.Worksheets(1).Activate
    On Error GoTo 0
    ScrollToTop

    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video3\V3_S0_Opening.mp3"

    WaitSec 3

    ' Slowly scroll to show the messy data
    SmoothScrollDown 12, 500

    WaitSec m_ClipDurSec - 9

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 28 — Data Sanitizer (~60 sec)
' Audio: V3_C1A_DataSanitizer.mp3
' Action: Run PreviewSanitizeChanges, then RunFullSanitize
'===============================================================================
Private Sub V3_Clip28_DataSanitizer()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video3\V3_C1A_DataSanitizer.mp3"

    WaitSec 3

    ' Preview first
    On Error Resume Next
    modUTL_DataSanitizer.PreviewSanitizeChanges
    On Error GoTo 0
    DoEvents
    WaitSec 4

    ' Now run full sanitize
    On Error Resume Next
    modUTL_DataSanitizer.RunFullSanitize
    On Error GoTo 0
    DoEvents

    WaitSec 3
    SmoothScrollDown 6, 400

    WaitSec m_ClipDurSec - 14

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 29 — Highlights (~35 sec)
' Audio: V3_C1B_Highlights.mp3
' Action: Run HighlightByThreshold, HighlightDuplicateValues
'===============================================================================
Private Sub V3_Clip29_Highlights()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video3\V3_C1B_Highlights.mp3"

    WaitSec 3

    ' Run threshold highlighting
    ' These macros may prompt for input — use SendKeys
    Application.SendKeys "5000{ENTER}", True
    On Error Resume Next
    modUTL_Highlights.HighlightByThreshold
    On Error GoTo 0
    DoEvents
    WaitSec 3

    ' Run duplicate highlighting
    On Error Resume Next
    modUTL_Highlights.HighlightDuplicateValues
    On Error GoTo 0
    DoEvents
    WaitSec 3

    SmoothScrollDown 4, 400

    WaitSec m_ClipDurSec - 12

    SilencePad
    StopAudio

    ' Clear highlights for next clip
    On Error Resume Next
    modUTL_Highlights.ClearHighlights
    On Error GoTo 0
End Sub

'===============================================================================
' CLIP 30 — Comments (~40 sec)
' Audio: V3_C1C_Comments.mp3
' Action: Run CountComments, ExtractAllComments
'===============================================================================
Private Sub V3_Clip30_Comments()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video3\V3_C1C_Comments.mp3"

    WaitSec 3

    On Error Resume Next
    modUTL_Comments.CountComments
    On Error GoTo 0
    DoEvents
    WaitSec 3

    ' Dismiss any MsgBox
    Application.SendKeys "{ENTER}", True
    DoEvents
    WaitSec 1

    On Error Resume Next
    modUTL_Comments.ExtractAllComments
    On Error GoTo 0
    DoEvents

    WaitSec 3

    WaitSec m_ClipDurSec - 12

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 31 — Tab Organizer (~50 sec)
' Audio: V3_C2A_TabOrganizer.mp3
' Action: Run ColorTabsByKeyword, ReorderTabs
'===============================================================================
Private Sub V3_Clip31_TabOrganizer()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video3\V3_C2A_TabOrganizer.mp3"

    WaitSec 3

    On Error Resume Next
    modUTL_TabOrganizer.ColorTabsByKeyword
    On Error GoTo 0
    DoEvents
    WaitSec 3

    On Error Resume Next
    modUTL_TabOrganizer.ReorderTabs
    On Error GoTo 0
    DoEvents
    WaitSec 3

    WaitSec m_ClipDurSec - 10

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 32 — Column Ops (~50 sec)
' Audio: V3_C2B_ColumnOps.mp3
' Action: Demo SplitColumn and CombineColumns
'===============================================================================
Private Sub V3_Clip32_ColumnOps()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video3\V3_C2B_ColumnOps.mp3"

    WaitSec 3

    On Error Resume Next
    modUTL_ColumnOps.SplitColumn
    On Error GoTo 0
    DoEvents
    WaitSec 4

    On Error Resume Next
    modUTL_ColumnOps.CombineColumns
    On Error GoTo 0
    DoEvents
    WaitSec 4

    WaitSec m_ClipDurSec - 14

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 33 — Sheet Tools (~50 sec)
' Audio: V3_C2C_SheetTools.mp3
' Action: Run ListAllSheetsWithLinks, TemplateCloner
'===============================================================================
Private Sub V3_Clip33_SheetTools()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video3\V3_C2C_SheetTools.mp3"

    WaitSec 3

    On Error Resume Next
    modUTL_SheetTools.ListAllSheetsWithLinks
    On Error GoTo 0
    DoEvents
    WaitSec 4

    SmoothScrollDown 4, 350
    WaitSec 2

    ' TemplateCloner prompts for sheet name and count
    Application.SendKeys "Sheet1{ENTER}", True
    WaitSec 0.5
    Application.SendKeys "2{ENTER}", True

    On Error Resume Next
    modUTL_SheetTools.TemplateCloner
    On Error GoTo 0
    DoEvents
    WaitSec 3

    WaitSec m_ClipDurSec - 14

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 34 — Compare Sheets (~50 sec)
' Audio: V3_C3A_Compare.mp3
' Action: Run CompareSheets
'===============================================================================
Private Sub V3_Clip34_Compare()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video3\V3_C3A_Compare.mp3"

    WaitSec 3

    On Error Resume Next
    modUTL_Compare.CompareSheets
    On Error GoTo 0
    DoEvents
    WaitSec 3

    SmoothScrollDown 6, 400

    WaitSec m_ClipDurSec - 10

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 35 — Consolidate (~40 sec)
' Audio: V3_C3B_Consolidate.mp3
' Action: Run ConsolidateSheets
'===============================================================================
Private Sub V3_Clip35_Consolidate()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video3\V3_C3B_Consolidate.mp3"

    WaitSec 3

    On Error Resume Next
    modUTL_Consolidate.ConsolidateSheets
    On Error GoTo 0
    DoEvents
    WaitSec 3

    SmoothScrollDown 6, 400

    WaitSec m_ClipDurSec - 10

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 36 — Pivot Tools + Lookup/Validation (~60 sec)
' Audio: V3_C3C_PivotTools.mp3 then V3_C3D_LookupValidation.mp3
' Action: Demo PivotTools and LookupBuilder
'===============================================================================
Private Sub V3_Clip36_PivotLookup()
    ' --- Part A: Pivot Tools ---
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video3\V3_C3C_PivotTools.mp3"

    WaitSec 3

    On Error Resume Next
    modUTL_PivotTools.ListAllPivots
    On Error GoTo 0
    DoEvents
    WaitSec 3

    WaitSec m_ClipDurSec - 8
    StopAudio
    WaitSec 1

    ' --- Part B: Lookup/Validation ---
    PlayAudio AUDIO_BASE_PATH & "Video3\V3_C3D_LookupValidation.mp3"

    WaitSec 3

    On Error Resume Next
    modUTL_LookupBuilder.BuildVLOOKUP
    On Error GoTo 0
    DoEvents
    WaitSec 3

    On Error Resume Next
    modUTL_ValidationBuilder.CreateDropdownList
    On Error GoTo 0
    DoEvents
    WaitSec 3

    WaitSec m_ClipDurSec - 10

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 37 — Universal Command Center (~50 sec)
' Audio: V3_C4_CommandCenter.mp3
' Action: Launch the Universal Command Center
'===============================================================================
Private Sub V3_Clip37_CommandCenter()
    SilencePad

    PlayAudio AUDIO_BASE_PATH & "Video3\V3_C4_CommandCenter.mp3"

    WaitSec 3

    ' Show the Universal Command Center
    On Error Resume Next
    modUTL_CommandCenter.LaunchUTLCommandCenter
    On Error GoTo 0
    DoEvents

    WaitSec m_ClipDurSec - 5

    SilencePad
    StopAudio
End Sub

'===============================================================================
' CLIP 38 — Closing (~45 sec)
' Audio: V3_Closing.mp3
' Action: Hold on cleaned-up sample file, static
'===============================================================================
Private Sub V3_Clip38_Closing()
    ' Navigate to first sheet
    On Error Resume Next
    ActiveWorkbook.Worksheets(1).Activate
    On Error GoTo 0
    ScrollToTop

    SilencePad

    PlayClip "Video3", "V3_Closing.mp3"

    ' Hold still after "Thanks for watching"
    WaitSec 3

    SilencePad
    StopAudio
End Sub

'===============================================================================
'
'    SECTION 5: MASTER CONTROLS
'
'===============================================================================

'===============================================================================
' RunAllVideos - Run all 3 videos back-to-back with pauses between
'===============================================================================
Public Sub RunAllVideos()
    ResetMCI

    MsgBox "Starting full demo recording sequence." & vbCrLf & vbCrLf & _
           "Video 1: What's Possible (~5 min)" & vbCrLf & _
           "Video 2: Full Demo Walkthrough (~18 min)" & vbCrLf & _
           "Video 3: Universal Tools (~10 min)" & vbCrLf & vbCrLf & _
           "Total: ~33 min of recording." & vbCrLf & vbCrLf & _
           "Click OK when OBS is recording.", _
           vbInformation, "Director - Full Demo"

    RunVideo1

    ' Pause between videos
    MsgBox "Video 1 complete. Clean up and prepare for Video 2." & vbCrLf & _
           "Click OK when ready to continue.", vbInformation, "Director"
    CleanupAllOutputSheets

    RunVideo2

    MsgBox "Video 2 complete. Switch to Sample_Quarterly_Report.xlsm for Video 3." & vbCrLf & _
           "Click OK when ready to continue.", vbInformation, "Director"

    ' Note: User must manually switch to the sample file
    ' RunVideo3 should be called from the sample file
End Sub

'===============================================================================
' TestClip - Test any single clip by number (1-38)
' Usage: TestClip 4  (tests Data Quality Scan clip)
'===============================================================================
Public Sub TestClip(ByVal clipNum As Long)
    ResetMCI
    m_Aborted = False

    Select Case clipNum
        ' Video 1
        Case 1: V1_Clip1_TitleCard
        Case 2: V1_Clip2_OpeningHook
        Case 3: V1_Clip3_CommandCenter
        Case 4: V1_Clip4_DataQuality
        Case 5: V1_Clip5_VarianceCommentary
        Case 6: V1_Clip6_Dashboard
        Case 7: V1_Clip7_BridgeClosing

        ' Video 2
        Case 8: V2_Clip8_Opening
        Case 9: V2_Clip9_WorkbookTour
        Case 10: V2_Clip10_CommandCenter
        Case 11: V2_Clip11_GLImport
        Case 12: V2_Clip12_DataQuality
        Case 13: V2_Clip13_Reconciliation
        Case 14: V2_Clip14_VarianceAnalysis
        Case 15: V2_Clip15_VarianceCommentary
        Case 16: V2_Clip16_YoYVariance
        Case 17: V2_Clip17_DashboardCharts
        Case 18: V2_Clip18_ExecDashboard
        Case 19: V2_Clip19_PDFExport
        Case 20: V2_Clip20_ExecBrief
        Case 21: V2_Clip21_ExecutiveMode
        Case 22: V2_Clip22_VersionControl
        Case 23: V2_Clip23_WhatIf
        Case 24: V2_Clip24_Sensitivity
        Case 25: V2_Clip25_IntegrationTest
        Case 26: V2_Clip26_AuditLogClosing

        ' Video 3
        Case 27: V3_Clip27_Opening
        Case 28: V3_Clip28_DataSanitizer
        Case 29: V3_Clip29_Highlights
        Case 30: V3_Clip30_Comments
        Case 31: V3_Clip31_TabOrganizer
        Case 32: V3_Clip32_ColumnOps
        Case 33: V3_Clip33_SheetTools
        Case 34: V3_Clip34_Compare
        Case 35: V3_Clip35_Consolidate
        Case 36: V3_Clip36_PivotLookup
        Case 37: V3_Clip37_CommandCenter
        Case 38: V3_Clip38_Closing

        Case Else
            MsgBox "Invalid clip number. Use 1-38.", vbExclamation, "Director"
    End Select

    StopAudio
    Application.StatusBar = False
End Sub

'===============================================================================
' QuickTest - Run a quick test of the audio + scroll + preflight system
'===============================================================================
Public Sub QuickTest()
    ResetMCI

    StatusMsg "Quick test — audio + scroll + preflight"

    ' --- Test 1: Pre-flight ---
    Debug.Print "[Director] QuickTest: Running pre-flight..."
    Dim pfOK As Boolean
    pfOK = PreFlightCheck(1)
    If Not pfOK Then
        MsgBox "Pre-flight failed. Fix issues and try again.", vbExclamation, "Director"
        Application.StatusBar = False
        Exit Sub
    End If

    ' --- Test 2: Audio playback + duration measurement ---
    Debug.Print "[Director] QuickTest: Testing audio..."
    GoToSheet "Report-->"
    SelectCell "A1"
    ScrollToTop

    Dim testFile As String
    testFile = AUDIO_BASE_PATH & "Video1\V1_S1_Opening_Hook.mp3"
    If Dir(testFile) = "" Then
        MsgBox "Test audio file not found:" & vbCrLf & testFile, vbExclamation, "Director"
        Application.StatusBar = False
        Exit Sub
    End If

    PlayAudio testFile
    Dim measuredDur As Double: measuredDur = m_ClipDurSec

    ' --- Test 3: Scroll ---
    WaitSec 2
    SmoothScrollDown 5, 400
    WaitSec 2

    StopAudio
    StatusMsg "Quick test complete"
    Application.StatusBar = False

    MsgBox "Quick test complete!" & vbCrLf & vbCrLf & _
           "Audio: " & IIf(measuredDur > 0 And measuredDur <> 30, "WORKING", "CHECK AUDIO") & vbCrLf & _
           "Measured duration: " & Format(measuredDur, "0.0") & " seconds" & vbCrLf & _
           "Scroll: Did the screen scroll smoothly?" & vbCrLf & _
           "Pre-flight: PASSED" & vbCrLf & vbCrLf & _
           "If audio played and scroll was smooth, you are ready to record.", _
           vbInformation, "Director - Quick Test"
End Sub
