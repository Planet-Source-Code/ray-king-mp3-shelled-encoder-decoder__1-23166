VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'------------------------------------------------
'-- clsMP3. Class Module.
'--
'-- Support Class For Using The Lame Frontend
'-- To Encode And Decode .Wav And .MP3 Files,
'-- In Visual Basic.
'--
'-- Lame.exe and Lame_enc.dll Are Required
'-- And Should Be In The Applications Directory.
'--
'-- You Can Download These From:
'-- http://www.dors.de/razorlane
'--
'--
'------------------------------------------------

'------------------------------------------------
'-- STARTUPINFO Structure Used By The Win32 API
'-- CreateProccessA.
'------------------------------------------------
Private Type STARTUPINFO
  cb              As Long      '-- Specifies the size, in bytes, of the structure.
  lpReserved      As String    '-- Reserved. Set this member to NULL before passing the structure to CreateProcess.
  lpDesktop       As String    '-- Windows NT/2000 or later:
  lpTitle         As String    '-- For console processes, this is the title displayed in the title bar.
  dwX             As Long      '-- Ignored unless dwFlags specifies STARTF_USEPOSITION.
  dwY             As Long      '-- Ignored unless dwFlags specifies STARTF_USEPOSITION.
  dwXSize         As Long      '-- Ignored unless dwFlags specifies STARTF_USESIZE.
  dwYSize         As Long      '-- Ignored unless dwFlags specifies STARTF_USESIZE.
  dwXCountChars   As Long      '-- Windows NT/2000 or later:
  dwYCountChars   As Long      '-- Windows NT/2000 or later:
  dwFillAttribute As Long      '-- Ignored unless dwFlags specifies STARTF_USEFILLATTRIBUTE.
  dwFlags         As Long      '-- This is a bit field that determines whether certain STARTUPINFO
                               '-- members are used when the process creates a window.
  wShowWindow     As Integer   '-- Ignored unless dwFlags specifies STARTF_USESHOWWINDOW.
  cbReserved2     As Integer   '-- Reserved; must be zero.
  lpReserved2     As Long      '-- Reserved; must be NULL.
  hStdInput       As Long      '-- Ignored unless dwFlags specifies STARTF_USESTDHANDLES.
  hStdOutput      As Long      '-- Ignored unless dwFlags specifies STARTF_USESTDHANDLES.
  hStdError       As Long      '-- Ignored unless dwFlags specifies STARTF_USESTDHANDLES.
End Type

'------------------------------------------------
'-- PROCCESS_INFORMATION Structure Used By The
'-- Win32 API CreateProccessA.
'------------------------------------------------
Private Type PROCESS_INFORMATION
  hProcess    As Long          '-- A handle to the newly created process.
  hThread     As Long          '-- A handle to the primary thread of the newly created process.
  dwProcessID As Long          '-- A global process identifier that can be used to identify a process.
  dwThreadID  As Long          '-- A global thread identifiers that can be used to identify a thread.
End Type

'------------------------------------------------
'-- Constants For SetWindowPos API.
'------------------------------------------------
Private Const HWND_TOPMOST = -1                      '-- Places the window above all non-topmost windows
Private Const HWND_NOTOPMOST = -2                    '-- Places the window above all non-topmost windows
                                                     '-- (that is, behind all topmost windows).
Private Const SWP_NOMOVE = &H2                       '-- Retains the current position (ignores the X and Y parameters).
Private Const SWP_NOSIZE = &H1                       '-- Retains the current size (ignores the cx and cy parameters).
Private Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE   '-- See Above Constants

'------------------------------------------------
'-- Constants For CreateProccessA And
'-- WaitForSingleOblect API's.
'------------------------------------------------
Private Const NORMAL_PRIORITY_CLASS = &H20&          '-- Indicates a normal process with no special scheduling needs.
Private Const INFINITE = -1&                         '-- Wait Forever.

'------------------------------------------------
'-- Win32 API Declarations.
'------------------------------------------------

'------------------------------------------------
'-- Win32 SetWindowPos API.
'------------------------------------------------
Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, _
         ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'------------------------------------------------
'-- Win32 WaitForSingleObject API.
'------------------------------------------------
Private Declare Function WaitForSingleObject Lib "kernel32" _
        (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
        
'------------------------------------------------
'-- Win32 CreateProccessA API.
'------------------------------------------------
Private Declare Function CreateProcessA Lib "kernel32" _
        (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, _
         ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
         ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
         ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
         lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
         
'------------------------------------------------
'-- Win32 TerminateProccess API.
'------------------------------------------------
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

'------------------------------------------------
'-- Win32 CloseHandle API.
'------------------------------------------------
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'------------------------------------------------
'-- Private Variables Constants, Etc.
'------------------------------------------------
Private Const APP_WAIT_TIME = 700                   '-- Application Wait Time (TIMEOUT) Change If You Need To:

Private udtProcInfo  As PROCESS_INFORMATION         '-- For Accessing The PROCESS_INFORMATION Structure
Private udtStartInfo As STARTUPINFO                 '-- For Accessing The STARTUPINFO Structure

'-- Enum Switch For Encoding And Decoding
Public Enum EncodeSwitch
  Encode = 1                                        '-- Encode The File
  Decode = 2                                        '-- Decode The File
End Enum

Private strSrcFileName  As String                   '-- The Source Filename
Private strDestFileName As String                   '-- The Destination Filename
Private strMP3Switches  As String                   '-- The MP3 Switches USed For Encoding And Decoding
Private lngEncodeSwitch As Long                     '-- Encoding Decoding Switch

'------------------------------------------------
'-- StartApp.
'--
'--Starts An Application And Runs It.
'--
'-- strAppName - The Application Name To Run.
'------------------------------------------------
Public Function StartApp(ByVal strAppName As String) As Long
  
  Dim lngRet As Long                   '-- For API Returned Value.
  
  '-- Create A New Proccess For Our Shelled Application.
  lngRet = CreateProcessA(0&, strAppName, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, 0&, udtStartInfo, udtProcInfo)
    
End Function

'------------------------------------------------
'-- EndApp.
'--
'-- Ends And Terminates The Running Application.
'------------------------------------------------
Public Sub EndApp()
  
  Dim lngRet As Long                                 '-- For API Returned Value
  
  lngRet = TerminateProcess(udtProcInfo.hProcess, 1) '-- Terminate The Shelled Application!.
  
End Sub

'------------------------------------------------
'-- WaitForAppToEnd.
'--
'-- Waits For The Running Application To End Or
'-- Until The Specified Timeout Has Ended.
'-- You Can Change The APP_WAIT_TIME Constant
'-- If You Need TO:
'------------------------------------------------
Public Sub WaitForAppToEnd()
    
  Dim lngRet As Long                   '-- For API Returned Value
  
  lngRet = 1                           '-- Initialize To 1 For Do Loop.
  
  Do '-- Loop Until The Application Returns.
    lngRet = WaitForSingleObject(udtProcInfo.hProcess, APP_WAIT_TIME) '-- Wait For The Application To End.
    DoEvents                                                          '-- Give Windows Some Proccessing Time.
  Loop Until lngRet = 0

End Sub

'------------------------------------------------
'-- KeepFormOnTop.
'--
'-- Keeps A Form On Top As The Top Most Window.
'--
'-- frm      - The Form To Keep On Top.
'-- blnOnTop - Flag For Switching From Top
'--            Or To Undo Form From Top.
'------------------------------------------------
Public Sub KeepFormOnTop(frm As Form, blnOnTop As Boolean)
  
  Dim lngPos As Long                   '-- For Window Position We Want

  If blnOnTop Then
    lngPos = HWND_TOPMOST              '-- We Want A Top Most Window.
  Else
    lngPos = HWND_NOTOPMOST            '-- We Do Not Want A Top Most Window.
  End If

  '-- Call SetWindowPos API To Keep Our Form On Top.
  Call SetWindowPos(frm.hwnd, lngPos, 0, 0, 0, 0, SWP_FLAGS)
  
End Sub

'-------------------------------------------------------
'-- MP3Encode.
'--
'-- Encodes Or Decodes A .MP3 Or A .Wav File.
'--
'-- Function Return:
'-- See ERR Constants Above.
'--
'-- NOTE: For Decoding And Encoding Switches See The
'--       Switches.Html File.
'-------------------------------------------------------
Public Function MP3Encode(frm As Form) As Long

  Dim lngRet As Long                         '-- For Function Returned Value.
  Dim lngErr As Long                         '-- For Error Detection.
  
  '-- Check To Make Sure We Have All Our Parameters!
  If strSrcFileName = "" Then
    lngErr = ERR_NO_SRC_FILE                 '-- No Source Filename Passed In.
  ElseIf strDestFileName = "" Then
    lngErr = ERR_NO_DEST_FILE                '-- No Destination Filename Passed In.
  Else
    lngErr = ERR_NONE                        '-- No ERRORS.
  End If
  
  '-- If We Have An Error Exit This Function With The Error Code...
  If lngErr <> ERR_NONE Then
    MP3Encode = lngErr                       '-- Assign The Error Code.
    Exit Function                            '-- And Exit This Function.
  End If
  
  Call KeepFormOnTop(frm, True)       '-- Keep The Main Application Form On Top.
  
  '-- Decode The .MP3 File To A .Wav File.
  lngRet = StartApp(App.Path & "\lame " & strMP3Switches & strSrcFileName & " " & strDestFileName)
  
  '-- Check For Error Starting The Application!
  If Not lngRet Then MP3Encode = ERR_NO_START_APP
  
  Call WaitForAppToEnd                       '-- Wait For The Started Application To End.
  
  Call KeepFormOnTop(frm, False)      '-- Release The Appliction Form From The Top.
    
End Function

'------------------------------------------------
'-- Property Get SrcFilename
'------------------------------------------------
Public Property Get SrcFilename() As String

  '-- Get The Source Filename.
  SrcFilename = strSrcFileName
  
End Property

'------------------------------------------------
'-- Property Let SrcFilename
'------------------------------------------------
Public Property Let SrcFilename(ByVal strNewValue As String)

  '-- Set The Source Filename.
  '-- Check For An Invalid Source Filename.
  If strNewValue = "" Then                    '-- We Have An Error So Raise It
    Err.Raise vbObjectError + 4000, "Class MP3 - Invalid Source Filename!"
  Else
    strSrcFileName = strNewValue              '-- No Error Assign String
  End If
  
End Property

'------------------------------------------------
'-- Property Get DestFilename
'------------------------------------------------
Public Property Get DestFilename() As String
  
  '-- Get The Destination Filename.
  DestFilename = strDestFileName
  
End Property

'------------------------------------------------
'-- Property Let DestFilename
'------------------------------------------------
Public Property Let DestFilename(ByVal strNewValue As String)

  '-- Set The Destination Filename.
  '-- Check For An Invalid Destination Filename.
  If strNewValue = "" Then                     '-- We Have An Error So Raise It.
    Err.Raise vbObjectError + 4001, "Class MP3 Invalid Destination Filename!."
  Else
    strDestFileName = strNewValue              '-- No Error So Assign The String
  End If
  
End Property

'------------------------------------------------
'-- Property Get EncodeDecode
'------------------------------------------------
Public Property Get EncodeDecode() As EncodeSwitch

  '-- Get Encode Decode Switch
  EncodeDecode = lngEncodeSwitch
  
End Property

'------------------------------------------------
'-- Property Let Encode Decode
'------------------------------------------------
Public Property Let EncodeDecode(ByVal enuNewValue As EncodeSwitch)
  
  '-- Assign New Encode Decode Switch
  lngEncodeSwitch = enuNewValue
  
End Property

'------------------------------------------------
'-- Property Get MP3Switches
'------------------------------------------------
Public Property Get MP3Switches() As String

  '-- Assign The New MP3 Switches
  strMP3Switches = MP3Switches
  
End Property

'------------------------------------------------
'-- Property Let MP3Switches
'------------------------------------------------
Public Property Let MP3Switches(ByVal strNewValue As String)

  Dim strTemp As String
  
  '-- Set The New MP3 Switches
  
  '-- If Were Decoding Then There's Only One Switch --decode.
  'If EncodeDecode = Decode Then strNewValue = "--decode"
  
  '-- We Need To Check Here To Make Sure We Have A Space On
  '-- The End Of The String.
  strTemp = Right(strNewValue, 1)
  
  If strTemp = " " Then
    strMP3Switches = strNewValue              '-- Have Space So Assign String
  Else
    strTemp = strNewValue & " "
    strMP3Switches = strTemp                  '-- No Space So Add One
  End If
  
  strTemp = ""                                '-- Clean Up
  
End Property

