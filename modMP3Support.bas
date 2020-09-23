Attribute VB_Name = "modMP3Support"
Option Explicit

'------------------------------------------------
'-- Lame MP3 Public Enums, Variables Etc.
'------------------------------------------------
'------------------------------------------------
'-- Public Enum Lame MP3 Error Codes.
'------------------------------------------------
Public Enum MP3Errors
  ERR_NONE = 0                         '-- No Error
  ERR_NO_FORM = 1                      '-- Error No Form Handle
  ERR_NO_SRC_FILE = 2                  '-- Error No Srouce Filename
  ERR_NO_DEST_FILE = 3                 '-- Error No Destination Filename
  ERR_NO_START_APP = 4                 '-- Error Starting Shelled Application
End Enum

'--------------------------------------------------
'-- User Defined Type To Hold The Lame MP3 Options
'--------------------------------------------------
Public Type MP3Options
  '-- General Options
  Bitrate           As Long            '-- The Current Bitrate [ -b 8 - 320 Switch ]
  Mode              As Long            '-- The Mode Stereo Etc [ -m Switch ]
  '-- Advanced Options
  Optimization      As Long            '-- Optimization Speed Etc [ -f Speed | -h Quality Switchs ]
  IncludeCRC        As Long            '-- Include CRC Checksums  [ -p Switch ]
  Copy              As Long            '-- Mark The Encoded File As A Copy [ -o Switch ]
  Copyright         As Long            '-- Mark The Encoded File As Being Copyrighted [ -c Switch ]
  Filing            As Long            '-- Delete Source File After Processing
  '-- VBR Options
  EnableVBR         As Long            '-- Enable Variable Bitrate VBR [ -B Value Switch ]
  VBRBitrate        As Long            '-- Maximum VBR Bitrate [ Used Above With The -B ]
  Quality           As Long            '-- Quality [ -V 0 - 9  Switch ]
  DisableVBRTag     As Long            '-- Disable Writing Of The VBR Tag [ -t Switch ]
  EnforceBitrate    As Long            '-- Strickly Enforce Minimum Bitrate [ -F Switch ]
  UseABR            As Long            '-- Use ABR Instead Of VBR [ --abr Value Switch ]
  ABRBitrate        As Long            '-- Target Bitrate For ABR [ Value For Above 8 - 320 ]
  '-- Expert Options
  BlockTypes        As Long            '-- Allow Block Types To Differ Between Channels [ -d Switch ]
  DisableFiltering  As Long            '-- Disable All Filtering [ -k Switch ]
  DisableBitRes     As Long            '-- Disable Bit reservoir [ --nores Switch ]
  DisableShort      As Long            '-- Disable Short Blocks [ --noshort Switch ]
  ComplyISO         As Long            '-- Comply As Much As Possibe To ISO MPEG Spec [ --strictly-enforce-ISO ]
  ATHControl        As Long            '-- ATH Control [ -athonly --noath --athshort Switches ]
  QLevel            As Long            '-- Quality Level [ -q Value 0 - 9 ]
  CustomOptions     As String          '-- Custom Option Switches
  UseCustom         As Long            '-- USe Custom Option Switches.
  '-- Audo Processing Options
  OutputFrequency   As Double          '-- Output Sampling Frequency [ --resample Value In kHz ]
  UseHighPassFilter As Long            '-- Use High Pass Filter Frequency
  UseHighPassWidth  As Long            '-- Use High Pass Filter Width
  HighPassFreq      As Double          '-- High Pass Filter Frequency In kHz
  HighPassWidth     As Double          '-- High Pass Filter Width In kHz
  UseLowPassFilter  As Long            '-- Use Low Pass Filter Frequency
  UseLowPassWidth   As Long            '-- Use Low Pass Filter Width
  LowPassFreq       As Double          '-- Low Pass Filter Frequency In kHz
  LowPassWidth      As Double          '-- Low Pass Filter Width In kHz
End Type

Public udtMP3        As MP3Options     '-- Make It Global To All
Public blnRetLoad    As Boolean        '-- Flag For Loading Options
Public blnRetSave    As Boolean        '-- Flag For Saving Options
Public strCommands   As String         '-- Command String For MP3 Switches
Public strDelSrcFile As String         '-- Path To Source File For Deleting.

'------------------------------------------------
'-- Sub Main - Main Entry Point To Application.
'------------------------------------------------
Public Sub Main()

  blnRetLoad = LoadOptions
      
  If blnRetLoad = False Then
    MsgBox "Error Loading Program Options!.", vbCritical
  End If
  
  Form1.Show
  
End Sub

'------------------------------------------------
'-- Builds A MP3 String For Lame.exe Switches.
'------------------------------------------------
Public Function BuildMP3String() As String

  Dim strTmp As String                       '-- Temp String Storage
  
  '----------------------------------------------
  '-- General Switches.
  '----------------------------------------------
  
  '-- Set Bitrate Switch Switch[-b Bitrate].
  strTmp = "-b " & udtMP3.Bitrate
  
  '-- Set Mode Switch.
  Select Case udtMP3.Mode
    Case 0    '-- Mode Stereo - Switch[-m s]
      strTmp = strTmp & " -m s"
    Case 1    '-- Mode Joint Stereo - Switch[-m j]
      strTmp = strTmp & " -m j"
    Case 2    '-- Mode Forced Joint Stereo - Switch[-m f]
      strTmp = strTmp & " -m f"
    Case 3    '-- Mode Mono - Switch[-m m]
      strTmp = strTmp & " -m m"
    Case 4    '-- Mode Default - No Switch
  End Select
  
  '----------------------------------------------
  '-- Advanced Switches.
  '----------------------------------------------
  
  '-- Set Optimization Switch.
  Select Case udtMP3.Optimization
    Case 0    '-- None - No Switch
    Case 1    '-- Speed - Switch[-f]
      strTmp = strTmp & " -f"
    Case 2    '-- Quality - Switch[-h]
      strTmp = strTmp & " -h"
  End Select
  
  '-- Set CRC Checksum Switch Switch[-p].
  If udtMP3.IncludeCRC = 1 Then strTmp = strTmp & " -p"
  
  '-- Set Copy Switch Switch[-o].
  If udtMP3.Copy = 1 Then strTmp = strTmp & " -o"
  
  '-- Set Copyright Switch Switch[-c].
  If udtMP3.Copyright = 1 Then strTmp = strTmp & " -c"
  
  '----------------------------------------------
  '-- VBR Switches.
  '----------------------------------------------
  
  '-- Set Enable Variable VBR Bitrate Switchs.
  If udtMP3.EnableVBR = 1 Then
    '-- Set Quality - Switch[-V Quality].
    If udtMP3.UseABR = 0 Then strTmp = strTmp & " -V " & udtMP3.Quality
    '-- Set VBR Bitrate - Switch[-B VBRBitrate].
    strTmp = strTmp & " -B " & udtMP3.VBRBitrate
    '-- Set Disable Writing Of VBR Tag - Switch[-t]
    If udtMP3.DisableVBRTag = 1 Then strTmp = strTmp & " -t"
    '-- Set Enforce Minimum Bitrate - Switch[-F].
    If udtMP3.EnforceBitrate = 1 Then strTmp = strTmp & " -F"
    '-- Set Use ABR Instead Of VBR - Switch[--abr Bitrate].
    If udtMP3.UseABR = 1 Then strTmp = strTmp & " --abr " & udtMP3.ABRBitrate
  End If
  
  '----------------------------------------------
  '-- Audio Processing Switches.
  '----------------------------------------------
  
  '-- Set Output Sampling Frequency - Switch[--resample kHz].
  Select Case udtMP3.OutputFrequency
    Case 0    '-- Default - NO Switch.
    Case 1    '-- 16 kHz.
      strTmp = strTmp & " --resample 16"
    Case 2    '-- 22.05 kHz.
      strTmp = strTmp & " --resample 22.05"
    Case 3    '-- 24 kHz.
      strTmp = strTmp & " --resample 24"
    Case 4    '-- 32 kHz.
      strTmp = strTmp & " --resample 32"
    Case 5    '-- 44.1 kHz.
      strTmp = strTmp & " --resample 44.1"
    Case 6    '-- 48 kHz.
      strTmp = strTmp & " --resample 48"
  End Select
  
  '-- Set High Pass Filtering Frequency - Switch[--highpass Value].
  If udtMP3.UseHighPassFilter = 1 Then strTmp = strTmp & " --highpass " & udtMP3.HighPassFreq
  
  '-- Set High Pass Width Filter - Switch[--highpass-width Value].
  If udtMP3.UseHighPassWidth = 1 And udtMP3.UseHighPassFilter = 1 Then
    strTmp = strTmp & " --highpass-width " & udtMP3.HighPassWidth
  End If
  
  '-- Set Low Pass Filtering Frequency - Switch[--highpass Value].
  If udtMP3.UseLowPassFilter = 1 Then strTmp = strTmp & " --lowpass " & udtMP3.LowPassFreq
  
  '-- Set Low Pass Width Filter - Switch[--highpass-width Value].
  If udtMP3.UseLowPassWidth = 1 And udtMP3.UseLowPassFilter = 1 Then
    strTmp = strTmp & " --lowpass-width " & udtMP3.LowPassWidth
  End If
  
  strCommands = strTmp                       '-- Assign String Switches
  
  strTmp = ""                                '-- Clean Up Strings
  
End Function

'------------------------------------------------
'-- Loads Options To The User Defined Type
'-- MP3Options.
'------------------------------------------------
Public Function LoadOptions() As Boolean

  Dim lngFileNum  As Long                    '-- For File Number Handle
  Dim strFileName As String                  '-- For The Filename
  
  On Error Resume Next                       '-- Let Us Handle ERROR'S
  
  lngFileNum = FreeFile                      '-- Get Free File Handle
  
  strFileName = App.Path & "\OptsMP3.dat"    '-- Assign The Filename
  
  Open strFileName For Input As #lngFileNum  '-- Open The File
    
  '-- Check For A File Open Error 53 File Not Found!.
  If Err.Number = 53 Then                    '-- File Was Not Found
    Call SetDefaultOptions                   '-- Set Options To Default
    blnRetSave = SaveOptions                 '-- Save Default Options
    If Not blnRetSave Then                   '-- Check For Save Options Error
      MsgBox "Error Creating " & strFileName, vbCritical
    End If
  End If
  
  '-- Read In The Program Options.
  Do While Not EOF(lngFileNum)               '-- Loop Until The End Of File.
    Input #lngFileNum, udtMP3.Bitrate, udtMP3.Mode, udtMP3.Optimization, _
    udtMP3.IncludeCRC, udtMP3.Copy, udtMP3.Copyright, udtMP3.Filing, _
    udtMP3.EnableVBR, udtMP3.VBRBitrate, udtMP3.Quality, _
    udtMP3.DisableVBRTag, udtMP3.EnforceBitrate, udtMP3.UseABR, _
    udtMP3.ABRBitrate, udtMP3.BlockTypes, udtMP3.DisableFiltering, _
    udtMP3.DisableBitRes, udtMP3.DisableShort, udtMP3.ComplyISO, _
    udtMP3.ATHControl, udtMP3.QLevel, udtMP3.CustomOptions, _
    udtMP3.UseCustom, udtMP3.OutputFrequency, udtMP3.UseHighPassFilter, _
    udtMP3.UseHighPassWidth, udtMP3.HighPassFreq, udtMP3.HighPassWidth, _
    udtMP3.UseLowPassFilter, udtMP3.UseLowPassWidth, udtMP3.LowPassFreq, _
    udtMP3.LowPassWidth
  Loop
  
  Close #lngFileNum                          '-- Close The File Handle
  
  '-- Return To Caller What Happened.
  If Err.Number <> 0 Then                    '-- An Error Occured
    LoadOptions = False
  Else                                       '-- No Error
    LoadOptions = True
  End If
    
End Function

'-----------------------------------------------
'-- Saves Options From The MP3Options User
'-- Defined Type.
'-----------------------------------------------
Public Function SaveOptions() As Boolean

  Dim lngFileNum  As Long                    '-- For File Number Handle
  Dim strFileName As String                  '-- For The Filename
  
  On Error Resume Next                       '-- Let US Handle Any ERROR'S
  
  lngFileNum = FreeFile                      '-- Assign The File Number
  
  strFileName = App.Path & "\OptsMP3.dat"    '-- Assign The Filename
  
  Open strFileName For Output As #lngFileNum '-- Open The Options File
    
  '-- Write Out All Options
  Write #lngFileNum, udtMP3.Bitrate, udtMP3.Mode, udtMP3.Optimization, _
  udtMP3.IncludeCRC, udtMP3.Copy, udtMP3.Copyright, udtMP3.Filing, _
  udtMP3.EnableVBR, udtMP3.VBRBitrate, udtMP3.Quality, _
  udtMP3.DisableVBRTag, udtMP3.EnforceBitrate, udtMP3.UseABR, _
  udtMP3.ABRBitrate, udtMP3.BlockTypes, udtMP3.DisableFiltering, _
  udtMP3.DisableBitRes, udtMP3.DisableShort, udtMP3.ComplyISO, _
  udtMP3.ATHControl, udtMP3.QLevel, udtMP3.CustomOptions, _
  udtMP3.UseCustom, udtMP3.OutputFrequency, udtMP3.UseHighPassFilter, _
  udtMP3.UseHighPassWidth, udtMP3.HighPassFreq, udtMP3.HighPassWidth, _
  udtMP3.UseLowPassFilter, udtMP3.UseLowPassWidth, udtMP3.LowPassFreq, _
  udtMP3.LowPassWidth
  
  Close #lngFileNum                          '-- Close The File Handle
  
  '-- Return To Caller What Happened.
  If Err.Number <> 0 Then                    '-- An Error Occured
    SaveOptions = False
  Else                                       '-- No Error
    SaveOptions = True
  End If
    
End Function

'----------------------------------------------
'-- Will Delete The Source File If User
'-- Selects To.
'----------------------------------------------
Public Function DeleteSourceFile()
  
  Kill strDelSrcFile                         '-- Delete The Source File
  
End Function
'------------------------------------------------
'-- Sets The MP3Options User Defined Type To
'-- Default Options.
'------------------------------------------------
Public Sub SetDefaultOptions()
  
  With udtMP3
    '-- General Options
    .Bitrate = 128        '-- The Current Bitrate [ -b 8 - 320 Switch ]
    .Mode = 4             '-- The Mode Stereo Etc [ -m Switch ]
    '-- Advanced Options
    .Optimization = 0     '-- Optimization Speed Etc [ -f Speed | -h Quality Switchs ]
    .IncludeCRC = 0       '-- Include CRC Checksums  [ -p Switch ]
    .Copy = 0             '-- Mark The Encoded File As A Copy [ -o Switch ]
    .Copyright = 0        '-- Mark The Encoded File As Being Copyrighted [ -c Switch ]
    .Filing = 0           '-- Delete Source File After Processing
    '-- VBR Options
    .EnableVBR = 0        '-- Enable Variable Bitrate VBR [ -B Value Switch ]
    .VBRBitrate = 128     '-- Maximum VBR Bitrate [ Used Above With The -B ]
    .Quality = 4          '-- Quality [ -V 0 - 9  Switch ]
    .DisableVBRTag = 0    '-- Disable Writing Of The VBR Tag [ -t Switch ]
    .EnforceBitrate = 0   '-- Strickly Enforce Minimum Bitrate [ -F Switch ]
    .UseABR = 0           '-- Use ABR Instead Of VBR [ --abr Value Switch ]
    .ABRBitrate = 128     '-- Target Bitrate For ABR [ Value For Above 8 - 320 ]
    '-- Expert Options
    .BlockTypes = 0       '-- Allow Block Types To Differ Between Channels [ -d Switch ]
    .DisableFiltering = 0 '-- Disable All Filtering [ -k Switch ]
    .DisableBitRes = 0    '-- Disable Bit reservoir [ --nores Switch ]
    .DisableShort = 0     '-- Disable Short Blocks [ --noshort Switch ]
    .ComplyISO = 0        '-- Comply As Much As Possibe To ISO MPEG Spec [ --strictly-enforce-ISO ]
    .ATHControl = 0       '-- ATH Control [ -athonly --noath --athshort Switches ]
    .QLevel = 0           '-- Quality Level [ -q Value 0 - 9 ]
    .CustomOptions = String(128, " ") '-- Custom Option Switches
    .UseCustom = 0        '-- USe Custom Option Switches.
    '-- Audo Processing Options
    .OutputFrequency = 0  '-- Output Sampling Frequency [ --resample Value In kHz ]
    .UseHighPassFilter = 0 '-- Use High Pass Filter Frequency
    .UseHighPassWidth = 0 '-- Use High Pass Filter Width
    .HighPassFreq = 0     '-- High Pass Filter Frequency In kHz
    .HighPassWidth = 0    '-- High Pass Filter Width In kHz
    .UseLowPassFilter = 0 '-- Use Low Pass Filter Frequency
    .UseLowPassWidth = 0  '-- Use Low Pass Filter Width
    .LowPassFreq = 0      '-- Low Pass Filter Frequency In kHz
    .LowPassWidth = 0     '-- Low Pass Filter Width In kHz
  End With
  
End Sub
'----------------------------------------------
'-- Takes The MP3 Bitrate And Converts It To
'-- A Tick Value For Window Slider Controls.
'----------------------------------------------
Public Function BitrateToTicks(ByVal lngBitrate As Long) As Long
  
  Select Case lngBitrate
    Case 8: BitrateToTicks = 1     '--   8 kbits
    Case 16: BitrateToTicks = 2    '--  16 kbits
    Case 24: BitrateToTicks = 3    '--  24 kbits
    Case 32: BitrateToTicks = 4    '--  32 kbits
    Case 40: BitrateToTicks = 5    '--  40 kbits
    Case 48: BitrateToTicks = 6    '--  48 kbits
    Case 56: BitrateToTicks = 7    '--  56 kbits
    Case 64: BitrateToTicks = 8    '--  64 kbits
    Case 80: BitrateToTicks = 9    '--  80 kbits
    Case 96: BitrateToTicks = 10   '--  96 kbits
    Case 112: BitrateToTicks = 11  '-- 112 kbits
    Case 128: BitrateToTicks = 12  '-- 128 kbits
    Case 144: BitrateToTicks = 13  '-- 144 kbits
    Case 160: BitrateToTicks = 14  '-- 160 kbits
    Case 192: BitrateToTicks = 15  '-- 192 kbits
    Case 224: BitrateToTicks = 16  '-- 224 kbits
    Case 256: BitrateToTicks = 17  '-- 256 kbits
    Case 320: BitrateToTicks = 18  '-- 320 kbits
  End Select
    
End Function
'----------------------------------------------
'-- Takes The The Slider Tick Value And
'-- Converts It To A Bitrate Value For
'-- MP3 Options.
'----------------------------------------------
Public Function TicksToBitrate(ByVal lngTick As Long) As Long
  
  Select Case lngTick
    Case 1: TicksToBitrate = 8     '--   8 kbits
    Case 2: TicksToBitrate = 16    '--  16 kbits
    Case 3: TicksToBitrate = 24    '--  24 kbits
    Case 4: TicksToBitrate = 32    '--  32 kbits
    Case 5: TicksToBitrate = 40    '--  40 kbits
    Case 6: TicksToBitrate = 48    '--  48 kbits
    Case 7: TicksToBitrate = 56    '--  56 kbits
    Case 8: TicksToBitrate = 64    '--  64 kbits
    Case 9: TicksToBitrate = 80    '--  80 kbits
    Case 10: TicksToBitrate = 96   '--  96 kbits
    Case 11: TicksToBitrate = 112  '-- 112 kbits
    Case 12: TicksToBitrate = 128  '-- 128 kbits
    Case 13: TicksToBitrate = 144  '-- 144 kbits
    Case 14: TicksToBitrate = 160  '-- 160 kbits
    Case 15: TicksToBitrate = 192  '-- 192 kbits
    Case 26: TicksToBitrate = 224  '-- 224 kbits
    Case 17: TicksToBitrate = 256  '-- 256 kbits
    Case 18: TicksToBitrate = 320  '-- 320 kbits
  End Select
  
End Function

