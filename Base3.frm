VERSION 5.00
Begin VB.Form Base 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Deeth Spectrum Analyzer"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   571
   ScaleMode       =   0  'User
   ScaleWidth      =   358.004
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "NORMAL READING"
      Height          =   615
      Left            =   11040
      TabIndex        =   29
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "PEAK READING"
      Height          =   615
      Left            =   11040
      TabIndex        =   28
      Top             =   7080
      Width           =   975
   End
   Begin VB.TextBox Text12 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   9120
      TabIndex        =   27
      Top             =   7680
      Width           =   255
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   9120
      TabIndex        =   26
      Top             =   7320
      Width           =   255
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   7560
      TabIndex        =   25
      Top             =   8040
      Width           =   255
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H000000FF&
      Height          =   285
      Left            =   7560
      TabIndex        =   24
      Top             =   7680
      Width           =   255
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   7560
      TabIndex        =   23
      Top             =   7320
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      Caption         =   "20 DB"
      Height          =   255
      Left            =   9480
      TabIndex        =   21
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "40 DB"
      Height          =   255
      Left            =   9480
      TabIndex        =   20
      Top             =   7320
      Width           =   855
   End
   Begin VB.CommandButton Command9 
      Caption         =   "60 DB"
      Height          =   255
      Left            =   7920
      TabIndex        =   19
      Top             =   8040
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "80 DB"
      Height          =   255
      Left            =   7920
      TabIndex        =   18
      Top             =   7680
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "100 DB"
      Height          =   255
      Left            =   7920
      TabIndex        =   17
      Top             =   7320
      Width           =   855
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   7080
      TabIndex        =   11
      Top             =   8040
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   7080
      TabIndex        =   10
      Top             =   7680
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H000000FF&
      Height          =   285
      Left            =   7080
      TabIndex        =   9
      Top             =   7320
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "11025"
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "22050"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "44100"
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   1095
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   7320
      Width           =   2895
   End
   Begin VB.Timer QuitTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   6240
   End
   Begin VB.PictureBox Scope 
      BackColor       =   &H80000009&
      ForeColor       =   &H80000002&
      Height          =   6090
      Left            =   840
      ScaleHeight     =   402
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   796
      TabIndex        =   2
      Top             =   480
      Width           =   12000
   End
   Begin VB.CommandButton StopButton 
      Caption         =   "S&top"
      Enabled         =   0   'False
      Height          =   570
      Left            =   120
      TabIndex        =   3
      Top             =   7800
      Width           =   984
   End
   Begin VB.ComboBox DevicesBox 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   3108
   End
   Begin VB.CommandButton StartButton 
      Caption         =   "&Start"
      Height          =   570
      Left            =   120
      TabIndex        =   1
      Top             =   7080
      Width           =   984
   End
   Begin VB.PictureBox ScopeBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   336
      Left            =   7200
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   336
   End
   Begin VB.Label Label7 
      Caption         =   "PEAK READING FREQUENCY"
      Height          =   255
      Left            =   1560
      TabIndex        =   22
      Top             =   6960
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "RANGE"
      Height          =   255
      Left            =   8640
      TabIndex        =   16
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "5 KHZ  BW"
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   8040
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "10 KHZ  BW"
      Height          =   255
      Left            =   4440
      TabIndex        =   14
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "20 KHZ  BW"
      Height          =   255
      Left            =   4440
      TabIndex        =   13
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "SAMPLE RATE"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   6960
      Width           =   1215
   End
End
Attribute VB_Name = "Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------
' Deeth Spectrum Analyzer v1.0
' A simple audio spectrum analyzer
'----------------------------------------------------------------------
' Opens a waveform audio device for 16-bit 11kHz mono input, gets
' chunks of audio, runs the FFT on them, and displays the output in
' a little window.  It's fairly optimized for speed.
' Demonstrates an easy and fairly fast way to do graphics double-
' buffering with a hidden picturebox, audio input, FFT usage, etc.
'----------------------------------------------------------------------
' The input audio data really should be windowed.  Maybe I'll do it
' later.
'----------------------------------------------------------------------
' Murphy McCauley (MurphyMc@Concentric.NET) 08/14/99
' http://www.fullspectrum.com/deeth/
'----------------------------------------------------------------------

Option Explicit

Dim Range As Long
Dim Reference As Long
Dim sr As Long
Dim freqscale As Long
Dim tick As Integer
Dim Factor2(4080) As Variant
Dim logvol3(4080) As Variant
Dim logvol2(4080) As Variant
Dim logvol1(4080) As Variant
Dim dbscale(10) As Variant
Dim peakflag As Integer
Dim n As Integer



Private DevHandle As Long 'Handle of the open audio device

Private Visualizing As Boolean
Private Divisor As Long

Private ScopeHeight As Long 'Saves time because hitting up a Long is faster
                            'than a property.
                            
Private Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    Reserved As Long
End Type

Private Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Private Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Private Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Private Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Private Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Private Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit

Private Const WAVE_FORMAT_PCM = 1

Private Const WHDR_DONE = &H1&              '/* done bit */
Private Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Private Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Private Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Private Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Private Const WIM_OPEN = &H3BE
Private Const WIM_CLOSE = &H3BF
Private Const WIM_DATA = &H3C0

Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal Flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long


Sub InitDevices()
    'Fill the DevicesBox box with all the compatible audio input devices
    'Bail if there are none.
    
    Dim Caps As WaveInCaps, Which As Long
    DevicesBox.Clear
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
        If Caps.Formats And WAVE_FORMAT_1M16 Then '16-bit mono devices
            Call DevicesBox.AddItem(StrConv(Caps.ProductName, vbUnicode), Which)
        End If
    Next
    If DevicesBox.ListCount = 0 Then
        MsgBox "You have no audio input devices!", vbCritical, "Ack!"
        End 'Ewww!  End!  Bad me!
    End If
    DevicesBox.ListIndex = 0
End Sub


Private Sub Command1_Click()  '44100

sr = 44100
Text2.BackColor = vbRed
Text3.BackColor = vbBlue
Text4.BackColor = vbBlue
Update

End Sub





Private Sub Command2_Click()  '22050

sr = 22050
Text2.BackColor = vbBlue
Text3.BackColor = vbRed
Text4.BackColor = vbBlue
Update

End Sub

Private Sub Command3_Click()  '11025

sr = 11025
Text2.BackColor = vbBlue
Text3.BackColor = vbBlue
Text4.BackColor = vbRed
Update

End Sub



Private Sub Command4_Click()  'PEAK READING
 
 peakflag = 1
 For n = 1 To 4080
    logvol2(n) = 0
    Next n
    
End Sub





Private Sub Command5_Click()  'normal reading

peakflag = 0
 For n = 1 To 4080
    logvol2(n) = 0
    Next n
    
End Sub

Private Sub Command7_Click() '100 DB

Divisor = 10000
Text8.BackColor = vbRed
Text9.BackColor = vbBlue
Text10.BackColor = vbBlue
Text11.BackColor = vbBlue
Text12.BackColor = vbBlue
dbscale(1) = "-25"
dbscale(2) = "-50"
dbscale(3) = "-75"
dbscale(4) = "-100"
Update
End Sub

Private Sub Command8_Click()  '80 DB

Divisor = 1000
Text8.BackColor = vbBlue
Text9.BackColor = vbRed
Text10.BackColor = vbBlue
Text11.BackColor = vbBlue
Text12.BackColor = vbBlue
dbscale(1) = "-20"
dbscale(2) = "-40"
dbscale(3) = "-60"
dbscale(4) = "-80"
Update
End Sub

Private Sub Command9_Click()  '60 DB

Divisor = 100
Text8.BackColor = vbBlue
Text9.BackColor = vbBlue
Text10.BackColor = vbRed
Text11.BackColor = vbBlue
Text12.BackColor = vbBlue
dbscale(1) = "-15"
dbscale(2) = "-30"
dbscale(3) = "-45"
dbscale(4) = "-60"
Update
End Sub

Private Sub Command10_Click()  '40 DB

Divisor = 10
Text8.BackColor = vbBlue
Text9.BackColor = vbBlue
Text10.BackColor = vbBlue
Text11.BackColor = vbRed
Text12.BackColor = vbBlue
dbscale(1) = "-10"
dbscale(2) = "-20"
dbscale(3) = "-30"
dbscale(4) = "-40"
Update
End Sub

Private Sub Command11_Click()  '20 db

Divisor = 1
Text8.BackColor = vbBlue
Text9.BackColor = vbBlue
Text10.BackColor = vbBlue
Text11.BackColor = vbBlue
Text12.BackColor = vbRed
dbscale(1) = "-5"
dbscale(2) = "-10"
dbscale(3) = "-15"
dbscale(4) = "-20"
Update
End Sub

Private Sub Form_Load()
    
    Range = 100
    Reference = 0
    sr = 44100
    freqscale = 4
    Divisor = 1000
    dbscale(0) = "  0"
    dbscale(1) = "-20"
    dbscale(2) = "-40"
    dbscale(3) = "-60"
    dbscale(4) = "-80"

    Call InitDevices 'Fill the DevicesBox
    
    Call DoReverse   'Pre-calculate these
    
    
    'Set the double buffer to match the display
    ScopeBuff.Width = Scope.ScaleWidth
    ScopeBuff.Height = Scope.ScaleHeight
    ScopeBuff.BackColor = Scope.BackColor
    
    ScopeHeight = Scope.Height
    
    Update
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If DevHandle <> 0 Then
        Call DoStop
        Cancel = 1
        If Visualizing = True Then
            QuitTimer.Enabled = True
        End If
    End If
End Sub


Private Sub QuitTimer_Timer()
    Unload Me
End Sub


Private Sub StartButton_Click()
    Static WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 1
        .SamplesPerSec = sr '11025  22050  44100
        .BitsPerSample = 16
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    Debug.Print "waveInOpen:"; waveInOpen(DevHandle, DevicesBox.ListIndex, VarPtr(WaveFormat), 0, 0, 0)
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Sub
    End If
    Debug.Print " "; DevHandle
    Call waveInStart(DevHandle)
    
    StopButton.Enabled = True
    StartButton.Enabled = False
    DevicesBox.Enabled = False
    
    Call Visualize
End Sub


Private Sub StopButton_Click()
    Call DoStop
    
    For n = 1 To 4080
    logvol3(n) = logvol2(n)
    logvol2(n) = 0
    Next n
    
End Sub


Private Sub DoStop()
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0
    StopButton.Enabled = False
    StartButton.Enabled = True
    DevicesBox.Enabled = True
End Sub


Private Sub Visualize()
    Visualizing = True
    
    'These are all static just because they can.
    Static X As Long
    Static Wave As WaveHdr
    Static InData(0 To NumSamples - 1) As Integer
    Dim InData2(0 To NumSamples - 1) As Integer
    Static OutData(0 To NumSamples - 1) As Single
    Dim PeakData As Single
    Dim freqpoint As Single
    Dim hamx(0 To NumSamples - 1) As Single
    
    
    For X = 0 To NumSamples - 1
    'hamx(X) = 0.54 - (0.46 * Cos((6.283185 * X) / (NumSamples - 1)))
    'hamx(X) = 0.42 - (0.5 * Cos((6.283185 * X) / (NumSamples - 1))) + (0.08 * Cos(6.283185 * 2 * (X) / (NumSamples - 1)))
    hamx(X) = 1
    'Debug.Print hamx(X)
    Next X
    
    
    
    With ScopeBuff 'Save some time referencing it...
    
        Do
            Wave.lpData = VarPtr(InData(0))
            Wave.dwBufferLength = NumSamples
            Wave.dwFlags = 0
            Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
            Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
            
            Do
                'Just wait for the blocks to be done or the device to close
            Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
            If DevHandle = 0 Then Exit Do 'Cut out if the device is closed
            
            Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
              
           'add windowing function here
            For X = 0 To NumSamples - 1
            InData2(X) = Int(InData(X) * hamx(X))
            Next X
                 
            Call FFTAudio(InData2, OutData)
            
           
            .Cls
            .CurrentX = 0
            .CurrentY = ScopeHeight
        
            For X = 1 To 4080
                '.CurrentY = ScopeHeight
               
                
                


                
                
                
If X < 101 / 4.9 And X > 10 / 4.9 Then 'multiplier freq 10-100 Hz
Factor2(X) = ((2606 / 4.9) * (Log(X / 10))) + 1
End If

If X < 1001 / 4.9 And X > 100 / 4.9 Then 'multiplier freq 100-1000 Hz
Factor2(X) = ((2606 / 4.9) * (Log(X / (100)))) + (6000 / 4.9)
End If

If X < 10001 / 4.9 And X > 1000 / 4.9 Then 'multiplier freq 1000-10000 Hz
Factor2(X) = ((2606 / 4.9) * (Log(X / (1000)))) + (12000 / 4.9)
End If

If X < 20001 / 4.9 And X > 10000 / 4.9 Then 'multiplier freq 10000-20000 Hz
Factor2(X) = ((2606 / 4.9) * (Log(X / (10000)))) + (18000 / 4.9)
End If
 '.CurrentX = ((Factor2(X)) / freqscale)
                'I average two elements here because it gives a smoother appearance.
        
        
        
        
            
            If (((Sqr(Abs(OutData(X * 2) / Divisor)) + Sqr(Abs(OutData((X * 2) + 1) / Divisor))))) > 0 Then
            logvol1(X) = (Log(((Sqr(Abs(OutData(X * 2) / Divisor)) + Sqr(Abs(OutData((X * 2) + 1) / Divisor))))) / 2.3) * 20
            'logvol1(X) = (Log(((Sqr(Abs(OutData(X * 2) / Divisor))))) / 2.3) * 20
            End If
            
            If logvol1(X) > logvol2(X) Then
            logvol2(X) = logvol1(X)
            End If
            
            If peakflag = 0 Then
            ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight - 8 * (logvol1(X)))
            End If
             
             If peakflag = 1 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight - 8 * (logvol2(X)))
             End If
             
     
             Next
            
            Scope.Picture = .Image 'Display the double-buffer
            DoEvents
        
        Loop While DevHandle <> 0
    
    End With
    
    Visualizing = False
    
    PeakData = 0
    
    For X = 1 To 4080
    
    If PeakData < logvol3(X) Then
    PeakData = logvol3(X)
    freqpoint = X
   
    Text1.Text = CStr((X) * sr * 2 / 16384) + " Hertz  " + vbCrLf
    Text1.Text = Text1.Text + CStr(120 - (2 * (Abs(PeakData)))) + vbCrLf
    If (X * 2) < 4080 Then
    Text1.Text = Text1.Text + CStr((X) * 4 * sr / (16384)) + " Hertz  " + CStr(120 - (2 * (logvol3(X * 2)))) + vbCrLf
    End If
    If (X * 3) < 4080 Then
    Text1.Text = Text1.Text + CStr((X) * 6 * sr / 16384) + " Hertz  " + CStr(120 - (2 * (logvol3(X * 3)))) + vbCrLf
    End If
    If (X * 4) < 4080 Then
    Text1.Text = Text1.Text + CStr((X) * 8 * sr / 16384) + " Hertz  " + CStr(120 - (2 * (logvol3(X * 4)))) + vbCrLf
    End If
    If (X * 5) < 4080 Then
    Text1.Text = Text1.Text + CStr((X) * 10 * sr / 16384) + " Hertz  " + CStr(120 - (2 * (logvol3(X * 5)))) + vbCrLf
    End If
    If (X * 6) < 4080 Then
    Text1.Text = Text1.Text + CStr((X) * 12 * sr / 16384) + " Hertz  " + CStr(120 - (2 * (logvol3(X * 6)))) + vbCrLf
    End If
    If (X * 7) < 4080 Then
    Text1.Text = Text1.Text + CStr((X) * 14 * sr / 16384) + " Hertz  " + CStr(120 - (2 * (logvol3(X * 7)))) + vbCrLf
    End If
    
    
    End If
    
    
    
    Next X
    
    
End Sub



Private Function Update()  'scale ht = 571   scale width = 417

Base.Cls



If sr = 44100 Then

    Base.CurrentY = 442
    Base.CurrentX = 20
    Base.Print "20"

    Base.CurrentY = 442
    Base.CurrentX = 90
    Base.Print "200"
    
    Base.CurrentY = 442
    Base.CurrentX = 220
    Base.Print "2000"
    
    Base.CurrentY = 442
    Base.CurrentX = 345
    Base.Print "20000"
    
End If

If sr = 22050 Then

    Base.CurrentY = 442
    Base.CurrentX = 20
    Base.Print "10"
    
    Base.CurrentY = 442
    Base.CurrentX = 90
    Base.Print "100"
    
    Base.CurrentY = 442
    Base.CurrentX = 220
    Base.Print "1000"
    
    Base.CurrentY = 442
    Base.CurrentX = 345
    Base.Print "10000"
    
End If

If sr = 11025 Then

    Base.CurrentY = 442
    Base.CurrentX = 20
    Base.Print "5"

    Base.CurrentY = 442
    Base.CurrentX = 90
    Base.Print "50"
    
    Base.CurrentY = 442
    Base.CurrentX = 220
    Base.Print "500"
    
    Base.CurrentY = 442
    Base.CurrentX = 345
    Base.Print "5000"
    
End If
    
    
    Base.CurrentY = 30
    Base.CurrentX = 10
    Base.Print dbscale(0)

    Base.CurrentY = 130
    Base.CurrentX = 10
    Base.Print dbscale(1)
    
    Base.CurrentY = 230
    Base.CurrentX = 10
    Base.Print dbscale(2)
    
    Base.CurrentY = 330
    Base.CurrentX = 10
    Base.Print dbscale(3)

    Base.CurrentY = 430
    Base.CurrentX = 10
    Base.Print dbscale(4)
    
    For tick = 0 To 20
    Base.CurrentY = 25 + (tick * 40 / 2)
    Base.CurrentX = 20
    Base.Print "__"
    Next tick
    
    
    
End Function

