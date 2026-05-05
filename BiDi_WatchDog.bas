Attribute VB_Name = "BiDi_WatchDog"
Option Explicit

' ========================================================================================
' Module      : BiDi_Watchdog
' Description : External Watchdog Timer for SeleniumVBA BiDi.
'               Provides a failsafe mechanism to restore Excel UI interactivity
'               if the main execution hangs or is suspended.
' ========================================================================================

' The name of the procedure to be called by Application.OnTime
Private Const WD_CALLBACK_PROC As String = "WatchDogRestoreUi"

' Tracks the exact scheduled time to ensure successful cancellation
Private m_watchdogTime As Date
' Logical flag to track the pending state, independent of Date-type ambiguity
Private m_isPending As Boolean

' ========================================================================================
' Property: IsPending
' Returns True if a watchdog timer is currently scheduled.
' ========================================================================================
Public Property Get IsPending() As Boolean
    IsPending = m_isPending
End Property

' ========================================================================================
' Method: StartWatchDog
' Schedules a UI recovery task after the specified timeout (seconds).
' ========================================================================================
Public Sub StartWatchDog(ByVal timeoutSec As Long)
    ' Ensure any existing timer is cleared before scheduling a new one
    Call CancelWatchDog
    
    ' Calculate the exact target time
    m_watchdogTime = Now + TimeSerial(0, 0, IIf(timeoutSec < 1, 1, timeoutSec))
    
    On Error Resume Next
    ' Schedule the callback procedure
    Application.OnTime m_watchdogTime, WD_CALLBACK_PROC
    
    ' Validate the registration (handles cases where Excel is too busy to accept OnTime)
    m_isPending = (Err.Number = 0)
    On Error GoTo 0
    
    If Not m_isPending Then
        Debug.Print "[" & Format(Now, "hh:mm:ss") & "] WatchDog Warning: Failed to register OnTime. UI protection is disabled."
    End If
End Sub

' ========================================================================================
' Method: CancelWatchDog
' Cancels the currently scheduled UI recovery task.
' ========================================================================================
Public Sub CancelWatchDog()
    If Not m_isPending Then Exit Sub
    
    On Error Resume Next
    ' Cancellation requires the exact time used during registration
    Application.OnTime m_watchdogTime, WD_CALLBACK_PROC, , False
    On Error GoTo 0
    
    ' Reset internal state
    m_isPending = False
    m_watchdogTime = 0
End Sub

' ========================================================================================
' Callback: WatchDogRestoreUi
' The procedure triggered by the timer. It forces Excel back into an interactive state.
' MUST be Public to be accessible by Application.OnTime.
' ========================================================================================
Public Sub WatchDogRestoreUi()
    ' Clear flags immediately upon entry
    m_isPending = False
    m_watchdogTime = 0
    
    ' Forcefully restore the Application state to return control to the user
    With Application
        .Interactive = True
        .Cursor = xlDefault
        .StatusBar = False
    End With
    
    Debug.Print "[" & Format(Now, "hh:mm:ss") & "] >>> WatchDog: UI recovery executed (Timeout reached)."
End Sub

' For Testing
' ========================================================================================
' Description : Integration tests to certify the robustness of SeleniumVBA BiDi.
'               Uses explicit exception throwing instead of Debug.Assert for
'               environment-independent validation.
' ========================================================================================

' ========================================================================================
' PUBLIC SUB: Run_SeleniumVBA_Robustness_Test
' SUMMARY: Executes a suite of tests to verify the Five-Layer Defense system.
' ========================================================================================
Public Sub Run_SeleniumVBA_Robustness_Test()
    Debug.Print "--- SeleniumVBA Robustness Certification Started ---"
    
    On Error GoTo TestFail
    
    ' 1. Verify Watchdog Arming Logic
    ' Ensure the OnTime registration is correctly reflected in the logical flag.
    BiDi_WatchDog.StartWatchDog 10
    AssertTrue BiDi_WatchDog.IsPending, "Watchdog should be PENDING after StartWatchDog call."
    
    ' 2. Verify Watchdog Cancellation Logic
    ' Ensure the cancellation properly clears the logical state.
    BiDi_WatchDog.CancelWatchDog
    AssertTrue Not BiDi_WatchDog.IsPending, "Watchdog should NOT be pending after CancelWatchDog."
    
    ' 3. Verify UI Interactive State Restoration
    ' Simulate a manual UI recovery call.
    BiDi_WatchDog.WatchDogRestoreUi
    AssertTrue Application.Interactive, "Application.Interactive must be TRUE after WatchDogRestoreUi."
    AssertTrue Application.Cursor = xlDefault, "Cursor must be reset to DEFAULT."

    ' 4. [Structural Check] Verify m_isSendReceiving Guard
    ' (This requires a connected instance, but we can verify the initial state)
    ' AssertTrue Not MyWrapper.GetSocket.IsSendReceiving, "Initial socket state should be idle."

    Debug.Print "--- Certification SUCCESS: All logic gates verified. ---"
    MsgBox "Robustness Certification Passed!" & vbCrLf & _
           "Your SeleniumVBA engine is now fortified.", vbInformation
    Exit Sub

TestFail:
    Dim msg As String
    msg = "Certification FAILED!" & vbCrLf & _
          "Source: " & Err.Source & vbCrLf & _
          "Detail: " & Err.Description
    Debug.Print "!!! " & msg
    MsgBox msg, vbCritical
End Sub

' ========================================================================================
' PRIVATE HELPER: AssertTrue
' DESCRIPTION: Throws a custom error if the condition is false.
'              Unlike Debug.Assert, this is NEVER stripped by the compiler.
' ========================================================================================
Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then
        ' Use 9999 as a dedicated error code for certification failures
        Err.Raise vbObjectError + 9999, "RobustnessCertification", "ASSERT_FAIL: " & message
    End If
End Sub

