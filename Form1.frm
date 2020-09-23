VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Form1 
   Caption         =   "MP3 Test Of Encoding And Decoding"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7530
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   31
      Top             =   7200
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17965
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame4 
      Caption         =   "Audio Processing Options"
      Height          =   3390
      Left            =   5100
      TabIndex        =   30
      Top             =   3750
      Width           =   5220
      Begin VB.Frame Frame15 
         Caption         =   "Low Pass Filter"
         Height          =   1110
         Left            =   105
         TabIndex        =   47
         Top             =   2145
         Width           =   4995
         Begin VB.TextBox txtLowPassWidth 
            Enabled         =   0   'False
            Height          =   330
            Left            =   3900
            TabIndex        =   24
            Text            =   "0"
            ToolTipText     =   "Sets The Low Pass Width Filter To kHz"
            Top             =   645
            Width           =   975
         End
         Begin VB.TextBox txtLowPassFreq 
            Enabled         =   0   'False
            Height          =   330
            Left            =   3900
            TabIndex        =   23
            Text            =   "0"
            ToolTipText     =   "Sets The Low Pass Filter Frequency In kHz"
            Top             =   270
            Width           =   975
         End
         Begin VB.CheckBox chkLowPassWidth 
            Caption         =   "Width Of Low Pass Filter In (kHz)."
            Enabled         =   0   'False
            Height          =   255
            Left            =   210
            TabIndex        =   22
            ToolTipText     =   "Enables The Low Pass Width Filer"
            Top             =   675
            Width           =   3285
         End
         Begin VB.CheckBox chkLowPassFreq 
            Caption         =   "Low Pass Filtering Frequency In (kHz)."
            Height          =   255
            Left            =   210
            TabIndex        =   21
            ToolTipText     =   "Enables The Low Pass Filter Frequency"
            Top             =   285
            Width           =   3615
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "High Pass Filter"
         Height          =   1110
         Left            =   105
         TabIndex        =   46
         Top             =   1005
         Width           =   4995
         Begin VB.TextBox txtHighPassWidth 
            Enabled         =   0   'False
            Height          =   330
            Left            =   3900
            TabIndex        =   20
            Text            =   "0"
            ToolTipText     =   "Sets The High Pass Width Filter In kHz"
            Top             =   645
            Width           =   975
         End
         Begin VB.TextBox txtHighPassFreq 
            Enabled         =   0   'False
            Height          =   330
            Left            =   3900
            TabIndex        =   19
            Text            =   "0"
            ToolTipText     =   "Sets The High Pass Filter Frequency In kHz"
            Top             =   255
            Width           =   975
         End
         Begin VB.CheckBox chkHighPassWidth 
            Caption         =   "Width Of High Pass Filter In (kHz)."
            Enabled         =   0   'False
            Height          =   270
            Left            =   210
            TabIndex        =   18
            ToolTipText     =   "Enables The High Pass Width Filter"
            Top             =   675
            Width           =   3330
         End
         Begin VB.CheckBox chkHighPassFreq 
            Caption         =   "High Pass Filtering Frequency in (kHz)."
            Height          =   270
            Left            =   210
            TabIndex        =   17
            ToolTipText     =   "Enables The Use Of High Pass Filter Frequency"
            Top             =   300
            Width           =   3645
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Output Sampling Frequency"
         Height          =   720
         Left            =   105
         TabIndex        =   45
         Top             =   285
         Width           =   3225
         Begin VB.ComboBox cboOutputFreq 
            Height          =   315
            ItemData        =   "Form1.frx":0000
            Left            =   165
            List            =   "Form1.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   16
            ToolTipText     =   "Sets The Output Frequency Value In kHz"
            Top             =   255
            Width           =   2850
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "VBR Options"
      Height          =   3645
      Left            =   5115
      TabIndex        =   29
      Top             =   45
      Width           =   5220
      Begin ComCtl2.UpDown UpDown2 
         Height          =   405
         Left            =   3855
         TabIndex        =   15
         Top             =   3045
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   714
         _Version        =   327681
         Value           =   4
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtABRBitrate"
         BuddyDispid     =   196623
         OrigLeft        =   3870
         OrigTop         =   3045
         OrigRight       =   4065
         OrigBottom      =   3450
         Max             =   310
         Min             =   4
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtABRBitrate 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   345
         Left            =   3390
         TabIndex        =   14
         Text            =   "128"
         ToolTipText     =   "Sets The ABR Bitrate"
         Top             =   3090
         Width           =   480
      End
      Begin VB.CheckBox chkUseABR 
         Caption         =   "Use ABR Instead Of VBR."
         Enabled         =   0   'False
         Height          =   255
         Left            =   1275
         TabIndex        =   13
         ToolTipText     =   "Use ABR Bitrate Instead Of VBR Bitrate"
         Top             =   2730
         Width           =   2550
      End
      Begin VB.CheckBox chkEnforceMinBitrate 
         Caption         =   "Enforce Minimum Bitrate."
         Enabled         =   0   'False
         Height          =   270
         Left            =   1275
         TabIndex        =   12
         ToolTipText     =   "Enforces The Minimum Bitrate"
         Top             =   2445
         Width           =   2490
      End
      Begin VB.CheckBox chkDisableVBRTag 
         Caption         =   "Disable Writing Of VBR Tag."
         Enabled         =   0   'False
         Height          =   300
         Left            =   1275
         TabIndex        =   11
         ToolTipText     =   "Disables The Writing Of The VBR Header Tag"
         Top             =   2160
         Width           =   2790
      End
      Begin VB.Frame Frame12 
         Caption         =   "Quality"
         Enabled         =   0   'False
         Height          =   885
         Left            =   105
         TabIndex        =   43
         Top             =   2115
         Width           =   975
         Begin ComCtl2.UpDown UpDown1 
            Height          =   405
            Left            =   630
            TabIndex        =   10
            Top             =   330
            Width           =   195
            _ExtentX        =   344
            _ExtentY        =   714
            _Version        =   327681
            AutoBuddy       =   -1  'True
            BuddyControl    =   "txtQuality"
            BuddyDispid     =   196628
            OrigLeft        =   570
            OrigTop         =   315
            OrigRight       =   765
            OrigBottom      =   720
            Max             =   9
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtQuality 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   345
            Left            =   150
            TabIndex        =   9
            Text            =   "4"
            ToolTipText     =   "Sets The Quality Value 0 - 9"
            Top             =   345
            Width           =   480
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Maximum VBR Bitrate"
         Enabled         =   0   'False
         Height          =   1395
         Left            =   120
         TabIndex        =   40
         Top             =   660
         Width           =   4965
         Begin MSComctlLib.Slider sldBitrate 
            Height          =   435
            Index           =   1
            Left            =   180
            TabIndex        =   8
            ToolTipText     =   "Sets The VBR Maximum Bitrate Value"
            Top             =   525
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   767
            _Version        =   393216
            Enabled         =   0   'False
            LargeChange     =   2
            Min             =   1
            Max             =   18
            SelectRange     =   -1  'True
            SelStart        =   12
            Value           =   12
            TextPosition    =   1
         End
         Begin VB.Label Label4 
            Caption         =   "Poor      Quality     Best"
            Enabled         =   0   'False
            Height          =   270
            Left            =   315
            TabIndex        =   42
            Top             =   1005
            Width           =   2070
         End
         Begin VB.Label Label3 
            Caption         =   "Current Bitrate: 128 kbits"
            Enabled         =   0   'False
            Height          =   225
            Left            =   240
            TabIndex        =   41
            Top             =   285
            Width           =   2175
         End
      End
      Begin VB.CheckBox chkEnableVBR 
         Caption         =   "Enable Vairable Bitrate (VBR)."
         Height          =   300
         Left            =   210
         TabIndex        =   7
         ToolTipText     =   "Enables The VBR Variable Bitrate And Options"
         Top             =   285
         Width           =   2895
      End
      Begin VB.Label Label5 
         Caption         =   "Target Bitrate For ABR:"
         Enabled         =   0   'False
         Height          =   270
         Left            =   1305
         TabIndex        =   44
         Top             =   3120
         Width           =   2025
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Advanced Options"
      Height          =   3945
      Left            =   105
      TabIndex        =   28
      Top             =   1875
      Width           =   4920
      Begin VB.Frame Frame10 
         Caption         =   "Filing"
         Height          =   720
         Left            =   105
         TabIndex        =   39
         Top             =   3075
         Width           =   4710
         Begin VB.CheckBox chkDelSource 
            Caption         =   "Delete Source File After Processing."
            Height          =   315
            Left            =   195
            TabIndex        =   6
            ToolTipText     =   "Deletes The Source File After Encoding It"
            Top             =   255
            Width           =   3450
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Flags"
         Height          =   1155
         Left            =   105
         TabIndex        =   38
         Top             =   1875
         Width           =   4710
         Begin VB.CheckBox chkCopyright 
            Caption         =   "Mark The Encoded File As Copyrighted."
            Height          =   300
            Left            =   195
            TabIndex        =   5
            ToolTipText     =   "Marks The MP3 File As Being Copyrighted"
            Top             =   660
            Width           =   3705
         End
         Begin VB.CheckBox chkCopy 
            Caption         =   "Mark The Encoded File As A Copy."
            Height          =   270
            Left            =   195
            TabIndex        =   4
            ToolTipText     =   "Marks The File As A Copy"
            Top             =   330
            Width           =   3315
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Options"
         Height          =   720
         Left            =   105
         TabIndex        =   37
         Top             =   1110
         Width           =   4710
         Begin VB.CheckBox chkCRC 
            Caption         =   "Include CRC Checksums."
            Height          =   300
            Left            =   195
            TabIndex        =   3
            ToolTipText     =   "Include CRC Checksums Adds 16 Bytes To Each Frame"
            Top             =   300
            Width           =   2520
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Optimizations"
         Height          =   720
         Left            =   105
         TabIndex        =   36
         Top             =   360
         Width           =   4710
         Begin VB.ComboBox cboOptimize 
            Height          =   315
            ItemData        =   "Form1.frx":005B
            Left            =   210
            List            =   "Form1.frx":0068
            Style           =   2  'Dropdown List
            TabIndex        =   2
            ToolTipText     =   "Set The Optimization For Speed Or Quality"
            Top             =   270
            Width           =   2205
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "General Options"
      Height          =   1815
      Left            =   105
      TabIndex        =   27
      Top             =   45
      Width           =   4920
      Begin VB.Frame Frame6 
         Caption         =   "Mode"
         Height          =   1395
         Left            =   2670
         TabIndex        =   35
         Top             =   270
         Width           =   2145
         Begin VB.ComboBox cboMode 
            Height          =   315
            ItemData        =   "Form1.frx":0082
            Left            =   135
            List            =   "Form1.frx":0095
            Style           =   2  'Dropdown List
            TabIndex        =   1
            ToolTipText     =   "Set The Encoding Mode"
            Top             =   330
            Width           =   1890
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Bitrate"
         Height          =   1395
         Left            =   105
         TabIndex        =   32
         Top             =   270
         Width           =   2490
         Begin MSComctlLib.Slider sldBitrate 
            Height          =   435
            Index           =   0
            Left            =   150
            TabIndex        =   0
            ToolTipText     =   "Set The Encoding Bitrate Value"
            Top             =   525
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   767
            _Version        =   393216
            LargeChange     =   2
            Min             =   1
            Max             =   18
            SelectRange     =   -1  'True
            SelStart        =   12
            Value           =   12
            TextPosition    =   1
         End
         Begin VB.Label Label2 
            Caption         =   "Poor     Quality      Best"
            Height          =   240
            Left            =   270
            TabIndex        =   34
            Top             =   1005
            Width           =   2070
         End
         Begin VB.Label Label1 
            Caption         =   "Current Bitrate: 128 kbits"
            Height          =   240
            Left            =   195
            TabIndex        =   33
            Top             =   315
            Width           =   2160
         End
      End
   End
   Begin MSComDlg.CommonDialog dlgFiles 
      Left            =   4605
      Top             =   6540
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDecode 
      Caption         =   "Decode Wav"
      Height          =   420
      Left            =   2685
      TabIndex        =   26
      ToolTipText     =   "Decodes A MP3 File Into A WAV File Using Options"
      Top             =   6585
      Width           =   1575
   End
   Begin VB.CommandButton cmdEncode 
      Caption         =   "Encode MP3"
      Height          =   420
      Left            =   735
      TabIndex        =   25
      ToolTipText     =   "Encodes A WAV File Into A MP3 File Using Options"
      Top             =   6585
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Sample: For Using Lame.exe To Encode And Decode                 MP3 And WAV Files. In Visual Basic 6.0"
      Height          =   525
      Left            =   150
      TabIndex        =   48
      Top             =   5865
      Width           =   4815
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MP3 As New clsMP3                                '-- Initialize Our MP3 Class

Private blnCancelFlag As Boolean                     '-- Flag For DialogBox Cancel

Private Sub cboMode_Click()
  
  udtMP3.Mode = cboMode.ListIndex                    '-- Set New Mode
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub cboOptimize_Click()
  
  udtMP3.Optimization = cboOptimize.ListIndex        '-- Set Optimization
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub cboOutputFreq_Click()

  udtMP3.OutputFrequency = cboOutputFreq.ListIndex   '-- Set New Output Sampling Frequency
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub chkCopy_Click()

  udtMP3.Copy = chkCopy.Value                        '-- Set Copy Flag
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub chkCopyright_Click()

  udtMP3.Copyright = chkCopyright.Value              '-- Set Copyright Flag
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub chkCRC_Click()

  udtMP3.IncludeCRC = chkCRC.Value                   '-- Set CRC Flag
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub chkDelSource_Click()

  udtMP3.Filing = chkDelSource.Value                 '-- Set Filing Flag
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub chkDisableVBRTag_Click()

  udtMP3.DisableVBRTag = chkDisableVBRTag.Value      '-- Set Disable Writing VBR Tag
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub chkEnableVBR_Click()

  Select Case chkEnableVBR.Value
    Case vbUnchecked                                 '-- Disable Controls
      Frame11.Enabled = False
      Label3.Enabled = False
      sldBitrate(1).Enabled = False
      Label4.Enabled = False
      Frame12.Enabled = False
      txtQuality.Enabled = False
      chkDisableVBRTag.Enabled = False
      chkEnforceMinBitrate.Enabled = False
      chkUseABR.Enabled = False
      Label5.Enabled = False
      txtABRBitrate.Enabled = False
    Case vbChecked                                   '-- Enable Controls
      Frame11.Enabled = True
      Label3.Enabled = True
      sldBitrate(1).Enabled = True
      Label4.Enabled = True
      Frame12.Enabled = True
      txtQuality.Enabled = True
      chkDisableVBRTag.Enabled = True
      chkEnforceMinBitrate.Enabled = True
      chkUseABR.Enabled = True
      If chkUseABR.Value = vbChecked Then
        Label5.Enabled = True
        txtABRBitrate.Enabled = True
      End If
      udtMP3.EnableVBR = 1
    Case vbGrayed
  End Select
  
  udtMP3.EnableVBR = chkEnableVBR.Value               '-- Set Enable VBR Flag
  Call DisplaySwitches                                '-- Display New Switches
  
End Sub

Private Sub chkEnforceMinBitrate_Click()

  udtMP3.EnforceBitrate = chkEnforceMinBitrate.Value  '-- Set Enforce Bitrate Flag
  Call DisplaySwitches                                '-- Display New Switches
  
End Sub

Private Sub chkHighPassFreq_Click()

  Select Case chkHighPassFreq.Value
    Case vbUnchecked                                  '-- Disable Controls
      txtHighPassWidth.Enabled = False
      chkHighPassWidth.Enabled = False
      txtHighPassFreq.Enabled = False
    Case vbChecked                                    '-- Enable Controls
      If chkHighPassWidth.Value = vbChecked Then
        txtHighPassWidth.Enabled = True
      Else
        txtHighPassWidth.Enabled = False
      End If
      chkHighPassWidth.Enabled = True
      txtHighPassFreq.Enabled = True
    Case vbGrayed
  End Select
  
  udtMP3.UseHighPassFilter = chkHighPassFreq.Value    '-- Set Use High Pass Filter Flag
  Call DisplaySwitches                                '-- Display New Switches
  
End Sub

Private Sub chkHighPassWidth_Click()

  Select Case chkHighPassWidth.Value
    Case vbUnchecked                                  '-- Disable Controls
      txtHighPassWidth.Enabled = False
    Case vbChecked                                    '-- Enable Controls
      txtHighPassWidth.Enabled = True
    Case vbGrayed
  End Select
  
  udtMP3.UseHighPassWidth = chkHighPassWidth.Value    '-- Set Use High Pass Width Flag
  Call DisplaySwitches                                '-- Display New Switches
  
End Sub

Private Sub chkLowPassFreq_Click()
  
  Select Case chkLowPassFreq.Value
    Case vbUnchecked                                  '-- Disable Controls
      txtLowPassWidth.Enabled = False
      txtLowPassFreq.Enabled = False
      chkLowPassWidth.Enabled = False
    Case vbChecked                                    '-- Enable Controls
      If chkLowPassWidth.Value = vbChecked Then
        txtLowPassWidth.Enabled = True
      Else
        txtLowPassWidth.Enabled = False
      End If
      txtLowPassFreq.Enabled = True
      chkLowPassWidth.Enabled = True
    Case vbGrayed
  End Select
    
  udtMP3.UseLowPassFilter = chkLowPassFreq.Value      '-- Set Use Low Pass Filter Flag
  Call DisplaySwitches                                '-- Display New Switches
  
End Sub

Private Sub chkLowPassWidth_Click()

  Select Case chkLowPassWidth.Value
    Case vbUnchecked                                  '-- Disable Controls
      txtLowPassWidth.Enabled = False
    Case vbChecked                                    '-- Enable Controls
      txtLowPassWidth.Enabled = True
    Case vbGrayed
  End Select
  
  udtMP3.UseLowPassWidth = chkLowPassWidth.Value      '-- Set Use Low Pass Width Flag
  Call DisplaySwitches                                '-- Display New Switches
  
End Sub

Private Sub chkUseABR_Click()

  Select Case chkUseABR.Value
    Case vbUnchecked                                  '-- Disable Controls
      Label5.Enabled = False
      txtABRBitrate.Enabled = False
      txtQuality.Enabled = True
    Case vbChecked                                    '-- Enable Controls
      Label5.Enabled = True
      txtABRBitrate.Enabled = True
      txtQuality.Enabled = False
    Case vbGrayed
  End Select
  
  udtMP3.UseABR = chkUseABR.Value                     '-- Set Use ABR Flag
  Call DisplaySwitches                                '-- Display New Switches
  
End Sub

Private Sub cmdDecode_Click()
  
  '----------------------------------------------
  '-- Decode A MP3 File Back To A Wav File.
  '----------------------------------------------
  MP3.EncodeDecode = Decode                           '-- We Want To Decode The File
  MP3.MP3Switches = "--decode -S " & strCommands      '-- Set Decoding Switchs The -S Switch Is Silent
  
  blnCancelFlag = False                               '-- Set Cancel To False
  
  Call GetFileNames                                   '-- Gets The Filenames
  
  '-- Check To See If The User Canceled File Selection.
  If blnCancelFlag Then                               '-- User Canceled!
    MsgBox "User Canceled File Selection!", vbInformation
    Exit Sub
  Else                                                '-- User Selected Files
    StatusBar1.Panels(1).Text = " Please Wait Decoding MP3 File To WAV!"
    MP3.MP3Encode Me                                  '-- Decode The File
    MsgBox "Finished Decoding File!", vbInformation, "Done Decoding File."
    Call DisplaySwitches                              '-- Display New Switches
  End If
  
End Sub

Private Sub cmdEncode_Click()

  '----------------------------------------------
  '-- Decode A MP3 File Back To A Wav File.
  '----------------------------------------------
  MP3.EncodeDecode = Encode                           '-- We Want To Encode The File
  MP3.MP3Switches = "-S " & strCommands               '-- Set Encoding Switchs The -S Switch Is Silent
  
  blnCancelFlag = False                               '-- Set Cancel To False
  
  Call GetFileNames                                   '-- Gets The Filenames
  
  '-- Check To See If The User Canceled File Selection.
  If blnCancelFlag Then                               '-- User Canceled!
    MsgBox "User Canceled File Selection!", vbInformation
    Exit Sub
  Else                                                '-- User Selected Files
    StatusBar1.Panels(1).Text = " Please Wait Encoding WAV File To MP3!"
    MP3.MP3Encode Me                                  '-- Decode The File
    MsgBox "Finished Encoding File!", vbInformation, "Done Encoding File."
    Call DisplaySwitches                              '-- Display New Switches
  End If
  
End Sub

Private Sub Form_Load()

  '-- Initialize Program Controls
  Call InitializeOptions
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim lngRet As Long                                 '-- MsgBox Return Value
  
  '-- Dispaly Message To See If User Wants To Save Program Options.
  lngRet = MsgBox("Do you want to Save Program Options?", vbYesNo + vbInformation, "Save Program Options")
  
  If lngRet = vbYes Then                             '-- User Wants To Save Options
    Call SaveOptions                                 '-- Save Program Options
    Set MP3 = Nothing                                '-- De-Initialize Our MP3 Class
  End If
  
  Set MP3 = Nothing                                  '-- De-Initialize Our MP3 Class
  
End Sub

Private Sub GetFileNames()
  
  '-----------------------------------------------------
  '-- Opens The File Open Dialog Box And Sets The MP3
  '-- Source And Destination Filenames. It Uses The
  '-- Same Source Filename For The Destination Filename,
  '-- But With A Different Filename Extension.
  '-----------------------------------------------------
  
  Dim lngLen As Long                                 '-- Length Of Our String
  Dim strTmp As String                               '-- Temp String Storage
  
  
  On Error Resume Next                               '-- Let Us Handle ERROR'S

  dlgFiles.CancelError = True                        '-- Causes A Trappable Error When The User Presses Cancel
  
  '-- See If Were Encoding Or Decoding A File
  If MP3.EncodeDecode = Decode Then
    dlgFiles.DialogTitle = "Open a MP3 File"         '-- Dialog Title For MP3 Files
    dlgFiles.Filter = "MP3 Files(*.mp3)|*.mp3"       '-- Dialog Filter For MP3 Files
  Else
    dlgFiles.DialogTitle = "Open a WAV File"         '-- Dialog Title For WAV Files
    dlgFiles.Filter = "WAV Files(*.wav)|*.wav"       '-- Dialog Filter For WAV Files
  End If
  
  dlgFiles.FileName = ""                             '-- Dialog Returned Filename Selected
  dlgFiles.FilterIndex = 0                           '-- Dialog Filter Index
  dlgFiles.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly  '-- Dialog Show Flags

  dlgFiles.ShowOpen                                  '-- Show Dialog Open
  
  If Err = cdlCancel Then                            '-- Cancel Button Was Hit
    blnCancelFlag = True
    Exit Sub                                         '-- Exit The Function
  End If
  
  '-- Set Our Source And Destination Filenames.
  If MP3.EncodeDecode = Decode Then
    MP3.SrcFilename = dlgFiles.FileName              '-- Set The Source Filename
    lngLen = Len(dlgFiles.FileName)                  '-- Get The Length Of The Filename
    strTmp = Left(dlgFiles.FileName, lngLen - 3)     '-- Strip The Filename Ext Off
    strTmp = strTmp & "wav"                          '-- Add New Filename Ext
    MP3.DestFilename = strTmp                        '-- Set The Destination Filename
    strDelSrcFile = strTmp                           '-- Set Filename For Deleting Source File
  Else
    MP3.SrcFilename = dlgFiles.FileName              '-- Same As Above Just Diff Ext
    lngLen = Len(dlgFiles.FileName)
    strTmp = Left(dlgFiles.FileName, lngLen - 3)
    strTmp = strTmp & "mp3"
    MP3.DestFilename = strTmp
    strDelSrcFile = strTmp                           '-- Set Filename For Deleting Source File
  End If
  
End Sub

Private Sub sldBitrate_Click(Index As Integer)

  Dim strTmp  As String
  Dim strTmp1 As String
  Dim lngBit  As Long
  
  strTmp = "Current Bitrate: "
  
  '-- Set Label Caption Strings
  Select Case sldBitrate(Index).Value
    Case 1: strTmp1 = strTmp & "8 kbits"
    Case 2: strTmp1 = strTmp & "16 kbits"
    Case 3: strTmp1 = strTmp & "24 kbits"
    Case 4: strTmp1 = strTmp & "32 kbits"
    Case 5: strTmp1 = strTmp & "40 kbits"
    Case 6: strTmp1 = strTmp & "48 kbits"
    Case 7: strTmp1 = strTmp & "56 kbits"
    Case 8: strTmp1 = strTmp & "64 kbits"
    Case 9: strTmp1 = strTmp & "80 kbits"
    Case 10: strTmp1 = strTmp & "96 kbits"
    Case 11: strTmp1 = strTmp & "112 kbits"
    Case 12: strTmp1 = strTmp & "128 kbits"
    Case 13: strTmp1 = strTmp & "144 kbits"
    Case 14: strTmp1 = strTmp & "160 kbits"
    Case 15: strTmp1 = strTmp & "192 kbits"
    Case 16: strTmp1 = strTmp & "224 kbits"
    Case 17: strTmp1 = strTmp & "256 kbits"
    Case 18: strTmp1 = strTmp & "320 kbits"
    Case Else
  End Select
  
  '-- Set Label Caption.
  Select Case Index
    Case 0
      Label1.Caption = strTmp1
    Case 1
      Label3.Caption = strTmp1
  End Select
  
  '-- Set MP3 Bitrate.
  Select Case Index
    Case 0          '-- Bitrate Slider
      udtMP3.Bitrate = TicksToBitrate(sldBitrate(0).Value)
    Case 1          '-- VBR Bitrate Slider
      udtMP3.VBRBitrate = TicksToBitrate(sldBitrate(1).Value)
  End Select
  
  '-- Clean Up Strings.
  strTmp = ""
  strTmp1 = ""
  
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub txtABRBitrate_Change()

  udtMP3.ABRBitrate = Val(Trim(txtABRBitrate))       '-- Set ABR Bitrate
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub txtHighPassFreq_Change()

  udtMP3.HighPassFreq = Val(Trim(txtHighPassFreq))   '-- Set High Pass Frequency
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub txtHighPassWidth_Change()

  udtMP3.HighPassWidth = Val(Trim(txtHighPassWidth)) '-- Set High Pass Width
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub txtLowPassFreq_Change()

  udtMP3.LowPassFreq = Val(Trim(txtLowPassFreq))     '-- Set Low Pass Frequency
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub txtLowPassWidth_Change()

  udtMP3.LowPassWidth = Val(Trim(txtLowPassWidth))   '-- Set Low Pass Width
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub txtQuality_Change()

  udtMP3.Quality = Val(Trim(txtQuality))             '-- Set Quality
  Call DisplaySwitches                               '-- Display New Switches
  
End Sub

Private Sub InitializeOptions()

  '-- Initializes The MP3 Program Options.
  
  '----------------------------------------------
  '-- Set Genaeral Options.
  '----------------------------------------------
  sldBitrate(0).Value = BitrateToTicks(udtMP3.Bitrate) '-- Set General Bitrate Slider.
  sldBitrate_Click (0)                                 '-- Set New Slider Label Value.
  cboMode.ListIndex = udtMP3.Mode                      '-- Set General Mode ComboBox.
  
  '----------------------------------------------
  '-- Set Advanced Options.
  '----------------------------------------------
  cboOptimize.ListIndex = udtMP3.Optimization          '-- Set Optimization ComboBox.
  chkCRC.Value = udtMP3.IncludeCRC                     '-- Set Options CRC CheckBox.
  chkCopy.Value = udtMP3.Copy                          '-- Set Flag Copy CheckBox.
  chkCopyright.Value = udtMP3.Copyright                '-- Set Flag Copyright CheckBox.
  chkDelSource.Value = udtMP3.Filing                   '-- Set Filing CheckBox.
  
  '----------------------------------------------
  '-- Set VBR Options.
  '----------------------------------------------
  chkEnableVBR.Value = udtMP3.EnableVBR                '-- Set Enable VBR CheckBox.
  sldBitrate(1).Value = BitrateToTicks(udtMP3.VBRBitrate) '-- Set Maximum VBR Bitrate.
  sldBitrate_Click (1)                                 '-- Set New Slider Label Value.
  txtQuality = udtMP3.Quality                          '-- Set Quality TextBox.
  chkDisableVBRTag.Value = udtMP3.DisableVBRTag        '-- Set Disable VBR Tag CheckBox.
  chkEnforceMinBitrate.Value = udtMP3.EnforceBitrate   '-- Set Enforce Minimum Bitrate CheckBox.
  chkUseABR.Value = udtMP3.UseABR                      '-- Set Use ABR CheckBox.
  txtABRBitrate = udtMP3.ABRBitrate                    '-- Set ABR Bitrate TextBox.
  
  '----------------------------------------------
  '-- Set Audio Processing Options.
  '----------------------------------------------
  cboOutputFreq.ListIndex = udtMP3.OutputFrequency     '-- Set Output Sampling Frequency ComboBox.
  chkHighPassFreq.Value = udtMP3.UseHighPassFilter     '-- Set High Pass Frequency CheckBox.
  chkHighPassWidth.Value = udtMP3.UseHighPassWidth     '-- Set High Pass Width CheckBox.
  txtHighPassFreq = udtMP3.HighPassFreq                '-- Set High Pass Frequency TextBox.
  txtHighPassWidth = udtMP3.HighPassWidth              '-- Set High Pass Width TextBox.
  chkLowPassFreq.Value = udtMP3.UseLowPassFilter       '-- Set Low Pass Frequency CheckBox.
  chkLowPassWidth.Value = udtMP3.UseLowPassWidth       '-- Set Low Pass Width CheckBox.
  txtLowPassFreq = udtMP3.LowPassFreq                  '-- Set Low Pass Frequency TextBox.
  txtLowPassWidth = udtMP3.LowPassWidth                '-- Set Low Pass Width TextBox.
  
  Call DisplaySwitches                                 '-- Display New Switches
  
End Sub

Private Sub DisplaySwitches()

  '----------------------------------------------
  '-- Displays The MP3 Switches In The
  '-- StatusBar.
  '----------------------------------------------
  Call BuildMP3String                                  '-- Build The New MP3 Switches String
    
  StatusBar1.Panels(1).Text = ""                       '-- Clear The StatusBar Text
  
  StatusBar1.Panels(1).Text = " " & strCommands        '-- Display New Switches String
  
End Sub
