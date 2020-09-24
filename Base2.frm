VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Base 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "DWL   Spectrum   Analyzer"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   -2535
   ClientWidth     =   11880
   Icon            =   "Base2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   529
   ScaleMode       =   0  'User
   ScaleWidth      =   329.314
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   0
      TabIndex        =   33
      Text            =   "0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text8 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   6960
      TabIndex        =   32
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox Text7 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   6600
      TabIndex        =   31
      Top             =   7440
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   6240
      TabIndex        =   30
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "M"
      Height          =   375
      Left            =   6960
      TabIndex        =   29
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "R"
      Height          =   375
      Left            =   6600
      TabIndex        =   28
      Top             =   6960
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "L"
      Height          =   375
      Left            =   6240
      TabIndex        =   27
      Top             =   6960
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9000
      TabIndex        =   26
      Text            =   "1"
      Top             =   7440
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   9000
      TabIndex        =   25
      Text            =   "1"
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RANGE DOWN"
      Height          =   375
      Left            =   4800
      TabIndex        =   24
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RANGE UP"
      Height          =   375
      Left            =   4800
      TabIndex        =   23
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton Command32 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9480
      TabIndex        =   22
      Top             =   7440
      Width           =   495
   End
   Begin VB.CommandButton Command31 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8520
      TabIndex        =   21
      Top             =   7440
      Width           =   495
   End
   Begin VB.TextBox Text18 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   10200
      TabIndex        =   19
      Top             =   7080
      Width           =   255
   End
   Begin VB.CommandButton Command28 
      Caption         =   "STOP"
      Height          =   375
      Left            =   11040
      TabIndex        =   18
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton Command27 
      Caption         =   "PLAY"
      Height          =   375
      Left            =   10200
      TabIndex        =   17
      Top             =   7440
      Width           =   735
   End
   Begin VB.CommandButton Command26 
      Caption         =   "LOAD WAV"
      Height          =   375
      Left            =   10560
      TabIndex        =   16
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command25 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8520
      TabIndex        =   15
      Top             =   6960
      Width           =   495
   End
   Begin VB.CommandButton Command24 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9480
      TabIndex        =   14
      Top             =   6960
      Width           =   495
   End
   Begin VB.PictureBox nullscope 
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   435
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComDlg.CommonDialog dlgfile 
      Left            =   120
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command18 
      BackColor       =   &H00C0C0C0&
      Caption         =   "REF DN"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   7440
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      BackColor       =   &H00C0C0C0&
      Caption         =   "REF UP"
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox Text14 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Text            =   "SPECTRUM PLOT USING PC SOUNDCARD"
      Top             =   0
      Width           =   8415
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   8160
      TabIndex        =   9
      Top             =   6960
      Width           =   255
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FF0000&
      Height          =   285
      Left            =   8160
      TabIndex        =   8
      Top             =   7440
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "AVE"
      Height          =   375
      Left            =   7440
      TabIndex        =   7
      Top             =   6960
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PEAK"
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   7440
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000009&
      Height          =   975
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   6840
      Width           =   2415
   End
   Begin VB.Timer QuitTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   6240
   End
   Begin VB.PictureBox Scope 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      ForeColor       =   &H00000000&
      Height          =   6090
      Left            =   840
      ScaleHeight     =   402
      ScaleMode       =   0  'User
      ScaleWidth      =   724
      TabIndex        =   2
      Top             =   480
      Width           =   10920
   End
   Begin VB.CommandButton StopButton 
      BackColor       =   &H00C0C0C0&
      Caption         =   "STOP"
      Enabled         =   0   'False
      Height          =   450
      Left            =   120
      TabIndex        =   3
      Top             =   7320
      Width           =   984
   End
   Begin VB.ComboBox DevicesBox 
      BackColor       =   &H80000009&
      Height          =   315
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton StartButton 
      BackColor       =   &H00C0C0C0&
      Caption         =   "START"
      Height          =   450
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   1
      Top             =   6840
      Width           =   984
   End
   Begin VB.PictureBox ScopeBuff 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C000C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000002&
      Height          =   336
      Left            =   11640
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   336
   End
   Begin VB.Label Label1 
      Caption         =   "Cal Offset"
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "DAZYWEB LABS  VB-3000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   0
      Width           =   3855
   End
   Begin VB.Menu menuFile 
      Caption         =   "File"
      Begin VB.Menu menuLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu menuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu menuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu menuPrint 
      Caption         =   "Print"
      Begin VB.Menu menuPlot 
         Caption         =   "Plot to Printer"
      End
      Begin VB.Menu menuPlotwidth 
         Caption         =   "Plotwidth"
         Begin VB.Menu menuPlotwide 
            Caption         =   "Wide"
         End
         Begin VB.Menu menuPlotnarrow 
            Caption         =   "Narrow"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu menuOptions 
      Caption         =   "Options"
      Begin VB.Menu menuSR 
         Caption         =   "Sample Rate"
         Begin VB.Menu menu44100 
            Caption         =   "44100"
            Checked         =   -1  'True
         End
         Begin VB.Menu menu22050 
            Caption         =   "22050"
         End
         Begin VB.Menu menu11025 
            Caption         =   "11025"
         End
      End
      Begin VB.Menu menuWindow 
         Caption         =   "Window"
         Begin VB.Menu menuBlackman 
            Caption         =   "Blackman"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuHamming 
            Caption         =   "Hamming"
         End
         Begin VB.Menu menuNowindow 
            Caption         =   "None"
         End
      End
      Begin VB.Menu menuDatalock 
         Caption         =   "Datalock"
         Begin VB.Menu menuPeakfreq 
            Caption         =   "Peak freq"
            Checked         =   -1  'True
         End
         Begin VB.Menu menu60Hz 
            Caption         =   "60 Hz"
         End
      End
      Begin VB.Menu menuScale 
         Caption         =   "Scale"
         Begin VB.Menu menuLog 
            Caption         =   "Log"
            Checked         =   -1  'True
         End
         Begin VB.Menu menuLin 
            Caption         =   "Lin"
         End
      End
      Begin VB.Menu menuCaloffset 
         Caption         =   "Cal Offset"
         Begin VB.Menu menuCalset 
            Caption         =   "Set"
         End
         Begin VB.Menu menuCaluse 
            Caption         =   "Engage"
         End
      End
   End
   Begin VB.Menu menuClear 
      Caption         =   "Clear"
   End
   Begin VB.Menu menuscope 
      Caption         =   "Scope"
   End
   Begin VB.Menu menuReset 
      Caption         =   "Reset"
   End
   Begin VB.Menu menuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------
'            DWL Spectrum Analyzer v1.0
'                    VB-3000
' An audio spectrum analyzer using the soundcard
'   build 6-02-01 by DazyWeb Laboratories  copyright 2001
'          Special credit to:
' Murphy McCauley (MurphyMc@Concentric.NET) 08/14/99
' http://www.fullspectrum.com/deeth/
' for building the core FFT module and Soundcard access routine.
'----------------------------------------------------------------------

Option Explicit
Dim calsetflag As Integer
Public caloffset As Single
Public caloffset2 As Single
Public calflag As Integer
Dim channels As Integer
Dim numchannels As Integer
Public firsttimeflag As Integer
Dim loadflag As Integer
Dim pwid As Single
Dim phgt As Single
Dim xmid As Single
Dim ymid As Single
Dim factor1 As Long
Dim y As Long
Dim freq(8160) As Single
Dim lplotold As Long
Dim lplotnew As Long
Dim texttoprint As String
    
    Dim X As Long
    Dim Wave As WaveHdr
    Dim InData(0 To (2 * (NumSamples - 1))) As Integer
    Dim InData2(0 To (2 * (NumSamples - 1))) As Integer
    Dim OutData(0 To NumSamples - 1) As Single
    Dim OutData2(0 To NumSamples - 1) As Single
    Dim PeakData As Single
    Dim freqpoint As Single
    Dim hamx(0 To NumSamples - 1) As Single
Dim ret As Variant
Dim I As Integer
Dim filename As String
Dim scalerange As Long
Dim logflag As Integer
Dim meterflag As Integer
Dim Range As Long
Dim Reference As Long
Dim sr As Long
Dim freqscale As Long
Dim tick As Integer
Dim Factor2(100000) As Double
Dim logvol3(8160) As Variant
Dim logvol2(8160) As Variant
Dim logvol1(8160) As Variant
Dim scopedata(19384) As Long
Dim dbscale(10) As Variant
Public peakflag As Integer
Dim aveflag As Integer
Dim avetimes As Long
Dim hamflag As Integer
Dim blackflag As Integer
Dim n As Long
Dim nn As Long
Dim nt As Long
Dim nnn As Long
Dim ss As Single
Dim Twopi As Single
Dim vertline(100) As Double
Dim ref As Long
Dim fnum1 As Variant
Dim loadedsamples As Long
Dim offset As Single
Dim sumlevel As Double
Dim thd As Single
Dim sigtonoise As Single
Dim fundlevel(20) As Double
Dim diflevel As Double
Dim harmlevel As Double
Dim highflag As Integer
Dim fname As String
Dim legend As String
Dim plotwidth As Integer




Private DevHandle As Long 'Handle of the open audio device

Private Visualizing As Boolean
Private Divisor As Long

Private ScopeHeight As Long 'Saves time because hitting up a Long is faster
                            'than a property.
                            
Private Type WaveFormatEx
    FormatTag As Integer
    channels As Integer
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
    channels As Integer
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




Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'''sndPlaySound Constants
Const SND_ALIAS = &H10000
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000
Const SND_LOOP = &H8
Const SND_MEMORY = &H4
Const SND_NODEFAULT = &H2
Const SND_NOSTOP = &H10
Const SND_SYNC = &H0
Const SND_PURGE = &H40

Dim SoundFile As String

Sub InitDevices()
    'Fill the DevicesBox box with all the compatible audio input devices
    'Bail if there are none.
    
    'Dim Caps As WaveInCaps, Which As Long
    'DevicesBox.Clear
    'For Which = 0 To waveInGetNumDevs - 1
    '    Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
    '    If Caps.Formats And WAVE_FORMAT_1M16 Then '16-bit mono devices
    '        Call DevicesBox.AddItem(StrConv(Caps.ProductName, vbUnicode), Which)
    '    End If
    'Next
    'If DevicesBox.ListCount = 0 Then
    '    MsgBox "You have no audio input devices!", vbCritical, "Ack!"
    '    End 'Ewww!  End!  Bad me!
    'End If
    'DevicesBox.ListIndex = 0
End Sub




Private Sub menuCalset_Click()  'set dB scale calibration offset

If calsetflag = 0 Then
Label1.Visible = True
Text9.Visible = True
calsetflag = 1
Text9.Text = CStr(caloffset)
Else
caloffset = Val(Text9.Text)
calsetflag = 0
Label1.Visible = False
Text9.Visible = False
End If


End Sub

Private Sub menuCaluse_Click()  'engage Calibration offset

If calflag = 0 Then
calflag = 1
menuCaluse.Checked = True
caloffset2 = caloffset
Else: calflag = 0
menuCaluse.Checked = False
caloffset2 = 0
End If

End Sub

Private Sub Command4_Click()  'PEAK READING
 
    If peakflag = 0 Then
       peakflag = 1
       Text5.BackColor = vbRed
     For n = 1 To 8160
     logvol2(n) = logvol1(n)
     Next n
    Else
      peakflag = 0
      Text5.BackColor = vbBlue
    End If
    
End Sub

Private Sub Command3_Click() 'Left channel select
firsttimeflag = 1
channels = 1
numchannels = 2
Call StopButton_Click
Text4.BackColor = vbRed
Text7.BackColor = vbBlue
Text8.BackColor = vbBlue
Call StartButton_Click


End Sub

Private Sub Command5_Click() 'Right channel select
firsttimeflag = 1
channels = 0
numchannels = 2
Call StopButton_Click
Text4.BackColor = vbBlue
Text7.BackColor = vbRed
Text8.BackColor = vbBlue
Call StartButton_Click

End Sub

Private Sub Command7_Click()  'mono mode
firsttimeflag = 1
channels = 3
numchannels = 1
Call StopButton_Click
Text4.BackColor = vbBlue
Text7.BackColor = vbBlue
Text8.BackColor = vbRed
Call StartButton_Click

End Sub

Private Sub menu44100_Click() '44100
firsttimeflag = 1
sr = 44100
menu11025.Checked = False
menu22050.Checked = False
menu44100.Checked = True
Update
Call StopButton_Click
Call StartButton_Click
End Sub





Private Sub menuBlackman_Click() 'Blackman window
 firsttimeflag = 1
 blackflag = 1
 menuBlackman.Checked = True
 menuHamming.Checked = False
 menuNowindow.Checked = False
 Call StopButton_Click
Call StartButton_Click
End Sub

Private Sub Command17_Click()  'REF UP

If ref < 80 Then
ref = ref + 20
End If

Update

Scope.Picture = nullscope.Image


replot

End Sub

Private Sub Command18_Click()  'REF DOWN

If ref > -100 Then
ref = ref - 20
End If

Update

Scope.Picture = nullscope.Image


replot

End Sub



Private Sub menu22050_Click()  '22050
firsttimeflag = 1
sr = 22050
menu11025.Checked = False
menu22050.Checked = True
menu44100.Checked = False
Update
Call StopButton_Click
Call StartButton_Click

End Sub


Private Sub menuLog_Click()  'LOG SCALE

menuLog.Checked = True
menuLin.Checked = False
scalerange = 1
logflag = 0
Update
replot

End Sub

Private Sub menuLin_Click()  'LIN SCALE

menuLin.Checked = True
menuLog.Checked = False

logflag = 1
Update
replot

End Sub

Private Sub Command24_Click()  'EXPAND SCALE

If scalerange < 32 Then
scalerange = scalerange * 2
Text2.Text = CStr(scalerange)
End If
Update
replot



End Sub

Private Sub Command25_Click()  'COMPRESS SCALE

If scalerange > 1 Then
scalerange = scalerange / 2
Text2.Text = CStr(scalerange)
End If
Update
replot

End Sub

Private Sub Command26_Click()  'LOAD WAV FILE

dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0

    LoadFile2 dlgfile.filename, dlgfile.FileTitle

End Sub


Sub LoadFile2(fname As String, fTitle As String)

filename = dlgfile.filename


End Sub


Private Sub Command27_Click()  'PLAY WAV FILE

If filename = "" Then
filename = "C:\VB2001files\wavcsetf8.wav"
End If

SoundFile = filename

'Ret = PlaySound(SoundFile, 0, SND_ASYNC Or SND_FILENAME)
ret = PlaySound(SoundFile, 0, SND_ASYNC Or SND_FILENAME Or SND_LOOP)



End Sub

Private Sub Command28_Click()   'STOP WAV FILE

'SoundFile = "c:\VB2001files\stopwav.wav"
'ret = PlaySound(SoundFile, 0, SND_ASYNC Or SND_FILENAME)


ret = PlaySound(0&, 0&, SND_PURGE Or SND_NODEFAULT)



End Sub

Private Sub menuPeakfreq_Click()  'PEAK DATA TO DATA WINDOW

menuPeakfreq.Checked = True
menu60Hz.Checked = False

meterflag = 0

End Sub

Private Sub menu11025_Click()  '11025
firsttimeflag = 1
sr = 11025
menu11025.Checked = True
menu22050.Checked = False
menu44100.Checked = False
Update
Call StopButton_Click
Call StartButton_Click


End Sub



Private Sub menu60Hz_Click()  '60 HZ TO DATA WINDOW
menu60Hz.Checked = True
menuPeakfreq.Checked = False

meterflag = 1

End Sub

Private Sub Command31_Click()  'slide left
If offset > -25 Then
offset = offset - 1
Text3.Text = CStr(offset)
End If
Update
replot

End Sub

Private Sub Command32_Click()  'slide right
If offset < 25 Then
offset = offset + 1
Text3.Text = CStr(offset)
End If
Update
replot

End Sub







Private Sub menuHamming_Click()  'Hamming filter

 hamflag = 1
 menuHamming.Checked = True
 menuBlackman.Checked = False
 menuNowindow.Checked = False
 Call StopButton_Click
Call StartButton_Click
End Sub

Private Sub Command6_Click() 'AVERAGE

If aveflag = 0 Then
aveflag = 1
Text6.BackColor = vbRed
Else
aveflag = 0
avetimes = 0
Text6.BackColor = vbBlue
End If


End Sub

Private Sub Command1_Click()  'RANGE UP

If Divisor < 10000000 Then
Divisor = Divisor * 10
End If

Update
Scope.Picture = nullscope.Image

replot

End Sub

Private Sub Command2_Click() 'RANGE DOWN

If Divisor > 1 Then
Divisor = Divisor / 10
End If

Update
Scope.Picture = nullscope.Image

replot

End Sub



Private Sub Form_Load()
    channels = 1
    ref = 0
    Range = 1000
    Reference = 0
    sr = 44100
    freqscale = 2
    Divisor = 1000
    dbscale(0) = -20
    dbscale(1) = -40
    dbscale(2) = -60
    dbscale(3) = -80
    dbscale(4) = -100
    Twopi = 6.28318530717958
    blackflag = 1
    Base.BackColor = &H80000004
    logflag = 0
    meterflag = 0
    scalerange = 1
    caloffset = 0
    calflag = 0
    'load ini file here
    
fname = "c:/vb3000init"
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Input As #fnum1
On Error Resume Next

Input #fnum1, sr
Input #fnum1, freqscale
Input #fnum1, Divisor
Input #fnum1, logflag
Input #fnum1, logflag
Input #fnum1, meterflag
Input #fnum1, scalerange
Input #fnum1, ref
Input #fnum1, Range
Input #fnum1, peakflag
Input #fnum1, aveflag
Input #fnum1, offset
Input #fnum1, blackflag
Input #fnum1, hamflag
Input #fnum1, plotwidth
Input #fnum1, channels
Input #fnum1, calflag
Input #fnum1, caloffset

Close fnum1


If calflag = 1 Then
menuCaluse.Checked = True
Else
menuCaluse.Checked = False
End If

If channels = 0 Or channels = 1 Then
numchannels = 2
Else
numchannels = 1
End If

If channels = 0 Then
Text4.BackColor = vbBlue
Text7.BackColor = vbRed
Text8.BackColor = vbBlue
End If

If channels = 1 Then
Text4.BackColor = vbRed
Text7.BackColor = vbBlue
Text8.BackColor = vbBlue
End If

If channels = 3 Then
Text4.BackColor = vbBlue
Text7.BackColor = vbBlue
Text8.BackColor = vbRed
End If

If plotwidth = 1 Then
menuPlotwide.Checked = True
menuPlotnarrow.Checked = False
Else
menuPlotwide.Checked = False
menuPlotnarrow.Checked = True
End If

If blackflag = 1 Then
menuBlackman.Checked = True
menuHamming.Checked = False
menuNowindow.Checked = False
End If

If hamflag = 1 Then
menuBlackman.Checked = False
menuHamming.Checked = True
menuNowindow.Checked = False
End If

If blackflag = 0 And hamflag = 0 Then
menuBlackman.Checked = False
menuHamming.Checked = False
menuNowindow.Checked = True
End If

If aveflag = 1 Then
Text6.BackColor = vbRed
End If

If peakflag = 1 Then
Text5.BackColor = vbRed
End If

If logflag = 1 Then
menuLog.Checked = False
menuLin.Checked = True
Else
menuLog.Checked = True
menuLin.Checked = False
End If
    
If meterflag = 1 Then
menuPeakfreq.Checked = False
menu60Hz.Checked = True
Else
menuPeakfreq.Checked = True
menu60Hz.Checked = False
End If
    
If sr = 44100 Then
menu44100.Checked = True
menu22050.Checked = False
menu11025.Checked = False
End If
    
 If sr = 22050 Then
menu44100.Checked = False
menu22050.Checked = True
menu11025.Checked = False
End If

If sr = 11025 Then
menu44100.Checked = False
menu22050.Checked = False
menu11025.Checked = True
End If
    
    Call InitDevices 'Fill the DevicesBox
    
    Call DoReverse   'Pre-calculate these
    
    
    'Set the double buffer to match the display
    ScopeBuff.Width = Scope.ScaleWidth
    ScopeBuff.Height = Scope.ScaleHeight
    ScopeBuff.BackColor = Scope.BackColor
    
    nullscope.Width = Scope.ScaleWidth
    nullscope.Height = Scope.ScaleHeight
    nullscope.BackColor = Scope.BackColor
    
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



Private Sub menuClear_Click() 'clear window

Scope.Cls
Scope.Picture = nullscope.Image
ScopeBuff.Picture = nullscope.Image

Update

End Sub

Private Sub menuExit_Click()

Call DoStop

fname = "c:/vb3000init"
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Output As #fnum1
On Error Resume Next

Write #fnum1, sr
Write #fnum1, freqscale
Write #fnum1, Divisor
Write #fnum1, logflag
Write #fnum1, logflag
Write #fnum1, meterflag
Write #fnum1, scalerange
Write #fnum1, ref
Write #fnum1, Range
Write #fnum1, peakflag
Write #fnum1, aveflag
Write #fnum1, offset
Write #fnum1, blackflag
Write #fnum1, hamflag
Write #fnum1, plotwidth
Write #fnum1, channels
Write #fnum1, calflag
Write #fnum1, caloffset

Close fnum1
Unload Base
End


End Sub

Private Sub menuHelp_Click() 'help file

helpform.Show



End Sub

Private Sub menuLoad_Click()  'load data

dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0

    LoadFile1 dlgfile.filename, dlgfile.FileTitle


End Sub

Sub LoadFile1(fname As String, fTitle As String)

loadflag = 1
    
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Input As #fnum1
On Error Resume Next

Input #fnum1, sr
Input #fnum1, loadedsamples

For n = 1 To 8160
Input #fnum1, OutData(n)    'old logvol1(n)
Next n

For n = 0 To 16383
Input #fnum1, scopedata(n)
Next n

Input #fnum1, aveflag
Input #fnum1, legend
Input #fnum1, channels
Close fnum1

'display data

Scope.ForeColor = vbBlue
ScopeBuff.ForeColor = vbBlue

If aveflag = 0 Then
Text6.BackColor = vbBlue
Else
Text6.BackColor = vbRed
End If


Text14.Text = legend

Update
replot

End Sub

Private Sub menuNowindow_Click() 'no window
blackflag = 0
hamflag = 0
menuNowindow.Checked = True
menuBlackman.Checked = False
menuHamming.Checked = False
Call StopButton_Click
Call StartButton_Click
End Sub

Private Sub menuPlot_Click()  'plot to printer

doplot

End Sub

Private Sub menuPlotnarrow_Click()  'plot narrow on printer

plotwidth = 0
menuPlotwide.Checked = False
menuPlotnarrow.Checked = True

End Sub

Private Sub menuPlotwide_Click()   'plot wide on printer

plotwidth = 1
menuPlotwide.Checked = True
menuPlotnarrow.Checked = False

End Sub

Private Sub menuReset_Click()  'Reset

ref = 0
   
    Range = 1000
    Reference = 0
    sr = 44100
    freqscale = 2
    Divisor = 1000
    dbscale(0) = -20
    dbscale(1) = -40
    dbscale(2) = -60
    dbscale(3) = -80
    dbscale(4) = -100
    Twopi = 6.28318530717958
    blackflag = 1
    Base.BackColor = &H80000004
    logflag = 0
    meterflag = 0
    scalerange = 1



If blackflag = 1 Then
menuBlackman.Checked = True
menuHamming.Checked = False
menuNowindow.Checked = False
End If

If hamflag = 1 Then
menuBlackman.Checked = False
menuHamming.Checked = True
menuNowindow.Checked = False
End If

If blackflag = 0 And hamflag = 0 Then
menuBlackman.Checked = False
menuHamming.Checked = False
menuNowindow.Checked = True
End If

If aveflag = 1 Then
Text6.BackColor = vbRed
End If

If peakflag = 1 Then
Text5.BackColor = vbRed
End If

If logflag = 1 Then
menuLog.Checked = False
menuLin.Checked = True
Else
menuLog.Checked = True
menuLin.Checked = False
End If
    
If meterflag = 1 Then
menuPeakfreq.Checked = False
menu60Hz.Checked = True
Else
menuPeakfreq.Checked = True
menu60Hz.Checked = False
End If
    
If sr = 44100 Then
menu44100.Checked = True
menu22050.Checked = False
menu11025.Checked = False
End If
    
 If sr = 22050 Then
menu44100.Checked = False
menu22050.Checked = True
menu11025.Checked = False
End If

If sr = 11025 Then
menu44100.Checked = False
menu22050.Checked = False
menu11025.Checked = True
End If


Update
replot


End Sub

Private Sub menuSave_Click()  'save data

dlgfile.Flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames
    
    On Error Resume Next
    dlgfile.ShowSave
    If Err.Number = cdlCancel Then Exit Sub
    If Err.Number <> 0 Then
        MsgBox "Error" & Str$(Err.Number) & _
            " selecting file." & vbCrLf & _
            Err.Description
    End If
    On Error GoTo 0

    SaveFile1 dlgfile.filename, dlgfile.FileTitle
    
    




End Sub

Sub SaveFile1(fname As String, fTitle As String)


    
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Output As #fnum1
On Error Resume Next

Write #fnum1, sr
Write #fnum1, NumSamples

For n = 1 To 8160
If aveflag = 1 Then
Write #fnum1, OutData2(n)
Else
Write #fnum1, OutData(n)
End If
Next n

For n = 0 To 16383
Write #fnum1, scopedata(n)
Next n

Write #fnum1, aveflag
Write #fnum1, Text14.Text
Write #fnum1, channels

Close fnum1


End Sub

Private Sub menuscope_Click() 'run scope

fname = "c:/wavscopedata"
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Output As #fnum1
On Error Resume Next

For n = 0 To 16383
Write #fnum1, scopedata(n)
Next n
Write #fnum1, channels
Close fnum1


Scopeform.Show

End Sub

Private Sub QuitTimer_Timer()
    Unload Me
End Sub


Private Sub StartButton_Click()            'START BUTTON
    Static WaveFormat As WaveFormatEx
    With WaveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .channels = numchannels         ' 1 = right only, 2 = stereo
        .SamplesPerSec = sr   '11025  22050  44100
        .BitsPerSample = 16
        .BlockAlign = (.channels * .BitsPerSample) \ 8
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
    
Text1.Text = "S.Rate = " + CStr(sr) + vbCrLf
If channels < 3 Then
Text1.Text = Text1.Text + "Sample size = 8K" + vbCrLf
Else
Text1.Text = Text1.Text + "Sample size = 16K" + vbCrLf
End If

    firsttimeflag = 1
    Call Visualize
End Sub


Private Sub StopButton_Click()          'STOP BUTTON
    Call DoStop
    
    For n = 1 To 8160
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
        
    ScopeBuff.ForeColor = vbBlack
    Scope.ForeColor = vbBlack



    For X = 0 To (NumSamples - 1)
     
     If hamflag = 1 Then
     hamx(X) = 0.54 - (0.54 * (Cos((Twopi * 2 * X) / NumSamples)))
     End If
     If blackflag = 1 Then
     hamx(X) = 0.42 - (0.5 * Cos((Twopi * 2 * X) / NumSamples)) + (0.08 * Cos(Twopi * 4 * X / NumSamples))
     End If
     If hamflag = 0 And blackflag = 0 Then
     hamx(X) = 1
     End If
      
    Next X
    
    
    If calflag = 1 Then
    caloffset2 = caloffset
    Else
    caloffset2 = 0
    End If
    
    
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
           
           'PARSE STEREO DATAFILE HERE
           
         If channels = 0 Or channels = 1 Then
         For X = 0 To (NumSamples - 3) / 2
         InData2(X) = (InData((X * 2) + channels) * hamx(X * 2))   'stereo +0 = left, +1 = right
         Next X
         
         For X = (NumSamples - 3) / 2 To (NumSamples - 3)
         InData2(X) = (InData((X * 2) + channels) * hamx(X / 2))
         Next X
         End If
            
            
         If channels = 3 Then
         For X = 0 To (NumSamples - 1)   'mono
         InData2(X) = (InData(X) * hamx(X))
         Next X
         End If
            
            
            
            
                 
            Call FFTAudio(InData2, OutData)
            
            
            
            
            If aveflag = 1 Then
            
            avetimes = avetimes + 1
            
            If avetimes > 2000000 Then
            avetimes = 1
            End If
            
            If avetimes = 1 Then
            For X = 1 To NumSamples - 1
            OutData2(X) = OutData(X)
            Next X
            End If
            
            For X = 1 To NumSamples - 1
             
            OutData2(X) = (Abs(OutData2(X) * avetimes + 1) + Abs(OutData(X))) / (avetimes + 1)
            OutData(X) = OutData2(X)
            
            Next X
            
            End If
           
             
            .Cls
            .CurrentX = 0
            .CurrentY = ScopeHeight
        
        
        
        
        
            Scope.ForeColor = &HBBBBBB         'draw dB grid lines
            ScopeBuff.ForeColor = &HBBBBBB
            For tick = 0 To 40
            If (tick - 1) / 5 = Int((tick - 1) / 5) Then
            ScopeBuff.ForeColor = &H888888
            Else
            ScopeBuff.ForeColor = &HBBBBBB
            End If
            
            Scope.Line (0, ((tick * 10) - 7))-(724, ((tick * 10) - 7))
            ScopeBuff.Line (0, ((tick * 10) - 7))-(724, ((tick * 10) - 7))
            Next tick
            Scope.ForeColor = vbBlack
            ScopeBuff.ForeColor = vbBlack
        
        
        
        If logflag = 0 Then    'print log freq lines
         
            Scope.ForeColor = &HAAAAAA
            ScopeBuff.ForeColor = &HAAAAAA
        
        For n = 1 To 3
  
  nt = 10 ^ n
  
For nnn = 1 To 10

nn = nnn * nt

        If nn < 101 / 11.966 And nn > 10 / 11.966 Then 'multiplier freq 10-100 Hz
Factor2(nn) = ((2606 / 11.966) * (Log(nn / 10))) + 1
End If

If nn < 1001 / 11.966 And nn > 100 / 11.966 Then 'multiplier freq 100-1000 Hz
Factor2(nn) = ((2606 / 11.966) * (Log(nn / (100)))) + (6000 / 11.966)
End If

If nn < 10001 / 11.966 And nn > 1000 / 11.966 Then 'multiplier freq 1000-10000 Hz
Factor2(nn) = ((2606 / 11.966) * (Log(nn / (1000)))) + (12000 / 11.966)
End If

If nn < 100001 / 11.966 And nn > 10000 / 11.966 Then 'multiplier freq 10000-20000 Hz
Factor2(nn) = ((2606 / 11.966) * (Log(nn / (10000)))) + (18000 / 11.966)
End If

If nn < 1000001 / 11.966 And nn > 100000 / 11.966 Then 'multiplier freq >20000 Hz
Factor2(nn) = ((2606 / 11.966) * (Log(nn / (100000)))) + (24000 / 11.966)
End If

ScopeBuff.Line ((Factor2(nn) / (freqscale * 0.98)) - 121, 0)-((Factor2(nn) / (freqscale * 0.98)) - 121, 402)
Scope.Line ((Factor2(nn) / (freqscale * 0.98)) - 121, 0)-((Factor2(nn) / (freqscale * 0.98)) - 121, 402)


Next nnn

Next n
        
            Scope.ForeColor = vbBlack
            ScopeBuff.ForeColor = vbBlack
        
        End If
        
        
        If logflag = 1 Then   'draw linear freq scale lines

Scope.ForeColor = &HAAAAAA
ScopeBuff.ForeColor = &HAAAAAA

For tick = 1 To 11
ScopeBuff.Line (((Int(tick * 71.9)) - 1), 0)-(((Int(tick * 71.9)) - 1), 402)
Scope.Line (((Int(tick * 71.9)) - 1), 0)-(((Int(tick * 71.9)) - 1), 402)
Next tick

Scope.ForeColor = vbBlack
ScopeBuff.ForeColor = vbBlack
End If
     
        
        
            .CurrentX = 0
            .CurrentY = ScopeHeight
        
     
        
            For X = 1 To 8160           '8160
                '.CurrentY = ScopeHeight
                              
                
                
If X < 101 / 11.966 And X > 10 / 11.966 Then 'multiplier freq 10-100 Hz
Factor2(X) = ((2606 / 11.966) * (Log(X / 10))) + 1
End If

If X < 1001 / 11.966 And X > 100 / 11.966 Then 'multiplier freq 100-1000 Hz
Factor2(X) = ((2606 / 11.966) * (Log(X / (100)))) + (6000 / 11.966)
End If

If X < 10001 / 11.966 And X > 1000 / 11.966 Then 'multiplier freq 1000-10000 Hz
Factor2(X) = ((2606 / 11.966) * (Log(X / (1000)))) + (12000 / 11.966)
End If

If X < 100001 / 11.966 And X > 10000 / 11.966 Then 'multiplier freq 10000-20000 Hz
Factor2(X) = ((2606 / 11.966) * (Log(X / (10000)))) + (18000 / 11.966)
End If


                
 If logflag = 1 Then   'linear scale factor calc
 
 Factor2(X) = (X * 784 * scalerange / 4080) + (784 * offset * (20 / 21.97) / 0.5)
 
 End If
               
               
        If blackflag = 1 Then
        OutData(X) = OutData(X) / 0.38
        End If
        
        If hamflag = 1 Then
        OutData(X) = OutData(X) / 0.5
        End If
        
        
        
        
            'I average two elements here because it gives a smoother appearance.
            
            If (X + 1) < 16384 And (((Sqr(Abs((OutData(X) + OutData(X + 1)) / 1000))))) > 0 Then
            logvol1(X) = ((Log(((Sqr(Abs((OutData(X) + OutData(X + 1)) / 1000))))) / 2.3) * 20) + (caloffset2 / 2.3)
                        
            End If
            
            
            If logvol1(X) > logvol2(X) Then   'peak hold data array
            logvol2(X) = logvol1(X)
            End If
            
           
            
             If peakflag = 0 And Divisor = 10000000 Then  '160 range
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 0 + (ref / 0.4) - (4 * (logvol1(X))))
             End If                                                              '25
             
             If peakflag = 1 And firsttimeflag = 0 And Divisor = 10000000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 0 + (ref / 0.4) - (4 * (logvol2(X))))
             End If
            
             If peakflag = 0 And Divisor = 1000000 Then   '140 range
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 15 + (ref / 0.35) - (5 * (logvol1(X))))
             End If                                                              '37
             
             If peakflag = 1 And firsttimeflag = 0 And Divisor = 1000000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 15 + (ref / 0.35) - (5 * (logvol2(X))))
             End If
            
            
             If peakflag = 0 And Divisor = 100000 Then  '120 range
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 50 + (ref / 0.3) - (6.6666 * (logvol1(X))))
             End If
             
             If peakflag = 1 And firsttimeflag = 0 And Divisor = 100000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 50 + (ref / 0.3) - (6.6666 * (logvol2(X))))
             End If
            
            
            If peakflag = 0 And Divisor = 1000 Then   '80 range
            ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 75 + (ref / 0.2) - (10 * (logvol1(X))))
            End If
             
             If peakflag = 1 And firsttimeflag = 0 And Divisor = 1000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 75 + (ref / 0.2) - (10 * (logvol2(X))))
             End If
             
            If peakflag = 0 And Divisor = 10000 Then   '100 range
            ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 75 + (ref / 0.25) - (8 * (logvol1(X))))
            End If
             
             If peakflag = 1 And firsttimeflag = 0 And Divisor = 10000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 75 + (ref / 0.25) - (8 * (logvol2(X))))
             End If
             
            If peakflag = 0 And Divisor = 100 Then   '60 range
            ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 150 + (ref / 0.15) - (13.333 * (logvol1(X))))
            End If
             
             If peakflag = 1 And firsttimeflag = 0 And Divisor = 100 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 150 + (ref / 0.15) - (13.333 * (logvol2(X))))
             End If
            
            If peakflag = 0 And Divisor = 10 Then   '40 range
            ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 175 + (ref / 0.1) - (20 * (logvol1(X))))
            End If
             
             If peakflag = 1 And firsttimeflag = 0 And Divisor = 10 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 175 + (ref / 0.1) - (20 * (logvol2(X))))
             End If
             
             If peakflag = 0 And Divisor = 1 Then  '20 range
            ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 350 + (ref / 0.05) - (40 * (logvol1(X))))
            End If
             
             If peakflag = 1 And firsttimeflag = 0 And Divisor = 1 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 350 + (ref / 0.05) - (40 * (logvol2(X))))
             End If
             
                       
             Next
             
   
        
        If firsttimeflag = 1 Then
           firsttimeflag = 0
           avetimes = 0           'reset averaging
           For n = 1 To 8160      'reset running peak
           logvol2(n) = logvol1(n)
           
           Next n
               
        End If
        
            
           
            
            Scope.Picture = .Image 'Display the double-buffer
            DoEvents
        
       
        
        Loop While DevHandle <> 0
    
    End With
    
    Visualizing = False
    
    PeakData = -20
    
         
    
    If meterflag = 0 Then
    
    highflag = 1
    
    For X = 1 To 8160
    
    
    If PeakData < logvol1(X) Then
    PeakData = logvol1(X)
    freqpoint = X
    
    
    Text1.Text = Left$(CStr((X) * sr / 16384), 6) + " Hertz  " + vbCrLf
    Text1.Text = Text1.Text + "-" + Left$(CStr(118 - (2 * (Abs(PeakData)))), 6) + " dB" + vbCrLf + vbCrLf
    
    If (X * 2) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 2 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 2)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 2
    End If
    If (X * 3) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 3 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 3)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 3
    End If
    If (X * 4) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 4 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 4)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 4
    End If
    If (X * 5) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 5 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 5)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 5
    End If
    If (X * 6) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 6 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 6)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 6
    End If
    If (X * 7) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 7 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 7)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 7
    End If
    If (X * 8) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 8 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 8)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 8
    End If
    If (X * 9) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 9 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 9)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 9
    End If
    If (X * 10) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 10 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 10)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 10
    End If
    If (X * 11) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 11 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 11)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 11
    End If
    If (X * 12) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 12 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 12)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 12
    End If
    If (X * 13) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 13 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 13)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 13
    End If
    If (X * 14) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 14 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 14)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 14
    End If
    If (X * 15) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 15 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 15)))), 6) + " dB" + vbCrLf + vbCrLf
    highflag = 15
    End If
    
    End If
   
    Next X
    
    
    ' S/N calculation
    
    'sumlevel = 0
    'sigtonoise = 0
    For n = 1 To 20
    fundlevel(n) = 0
    Next n
    'diflevel = 0
    harmlevel = 0
    
    
    'For X = 1 To 8160
    'sumlevel = sumlevel + ((OutData(X)) ^ 2)
    'Next X
    
    'sumlevel = (Sqr(sumlevel)) / 8160
    
    'Text1.Text = Text1.Text + "Total level = " + Left$(CStr(120 - (20 * (Log(sumlevel)) / Log(10))), 5) + "dB" + vbCrLf + vbCrLf
    
    
    For nn = 1 To highflag
       
       fundlevel(nn) = 0
        
    If (nn * freqpoint) < 8160 Then
    fundlevel(nn) = Abs(fundlevel(nn)) + (Abs(OutData(freqpoint * nn)))
    End If
      
    Next nn
    
    
    'diflevel = (Abs(fundlevel(1)) - (Abs(sumlevel)))
    
    'sigtonoise = fundlevel(1) / Abs(diflevel)   'diflevel
    'If sigtonoise > 0 Then
    'sigtonoise = (20 * ((Log(sigtonoise)) / (Log(10)))) '+ 25
    'End If
    'Text1.Text = Text1.Text + "S/N = " + Left$(CStr(sigtonoise), 5) + "dB" + vbCrLf + vbCrLf
    
    'If fundlevel(1) <> 0 Then
    'Text1.Text = Text1.Text + "THD+N = " + CStr(FormatNumber((100 * (sumlevel / fundlevel(1))), 4)) + " %" + vbCrLf
    'End If
    
    'total thd minus noise calc
    
    For n = 2 To highflag
    harmlevel = harmlevel + fundlevel(n)
    Next n
    harmlevel = (harmlevel) / 1
    
    thd = ((harmlevel) / (fundlevel(1))) * 50  ' not 100 due to windowing rolloff and worsened s/n and thd
    
    Text1.Text = "THD = " + CStr(FormatNumber((thd), 4)) + " %" + vbCrLf + Text1.Text
    
    
    
    End If   'for meterflag = 0
    
    If meterflag = 1 Then
    freqpoint = Int(60 / (sr / 16384))
    X = freqpoint
    PeakData = logvol1(X)
    
    Text1.Text = Left$(CStr((X) * sr / 16384), 6) + " Hertz  " + vbCrLf
    Text1.Text = Text1.Text + "-" + Left$(CStr(118 - (2 * (Abs(PeakData)))), 6) + " dB" + vbCrLf + vbCrLf
    
    If (X * 2) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 2 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 2)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 3) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 3 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 3)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 4) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 4 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 4)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 5) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 5 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 5)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 6) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 6 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 6)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 7) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 7 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 7)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 8) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 8 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 8)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 9) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 9 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 9)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 10) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 10 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 10)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 11) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 11 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 11)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 12) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 12 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 12)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 13) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 13 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 13)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 14) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 14 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 14)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    If (X * 15) < 8160 Then
    Text1.Text = Text1.Text + Left$(CStr((X) * 15 * sr / 16384), 6) + " Hertz  " + vbCrLf + "-" + Left$(CStr(118 - (2 * (logvol1(X * 15)))), 6) + " dB" + vbCrLf + vbCrLf
    End If
    
    End If   'for meterflag = 1
    
    
   For n = 0 To 16383
   scopedata(n) = InData(n)
   Next n
   
       
    avetimes = 1
    
End Sub



Private Function Update()  'scale ht = 571   scale width = 417

Base.Cls


If Divisor = 1 Then
dbscale(0) = -80
dbscale(1) = -85
dbscale(2) = -90
dbscale(3) = -95
dbscale(4) = -100
End If

If Divisor = 10 Then
dbscale(0) = -60
dbscale(1) = -70
dbscale(2) = -80
dbscale(3) = -90
dbscale(4) = -100
End If

If Divisor = 100 Then
dbscale(0) = -40
dbscale(1) = -45
dbscale(2) = -60
dbscale(3) = -85
dbscale(4) = -100
End If

If Divisor = 1000 Then
dbscale(0) = -20
dbscale(1) = -40
dbscale(2) = -60
dbscale(3) = -80
dbscale(4) = -100
End If

If Divisor = 10000 Then
dbscale(0) = 0
dbscale(1) = -25
dbscale(2) = -50
dbscale(3) = -75
dbscale(4) = -100
End If

If Divisor = 100000 Then
dbscale(0) = 20
dbscale(1) = -10
dbscale(2) = -40
dbscale(3) = -70
dbscale(4) = -100
End If

If Divisor = 1000000 Then
dbscale(0) = 40
dbscale(1) = 5
dbscale(2) = -30
dbscale(3) = -65
dbscale(4) = -100
End If

If Divisor = 10000000 Then
dbscale(0) = 60
dbscale(1) = 20
dbscale(2) = -20
dbscale(3) = -60
dbscale(4) = -100
End If



If logflag = 1 Then   ' linear freq scale numbering

    'Base.CurrentY = 442
    'Base.CurrentX = 24
    'Base.Print "0"

   For n = 0 To 10
    Base.CurrentY = 442
    Base.CurrentX = 20 + (n * 29.57)
    Base.Print CStr(Left$((n * (sr / 44100) * 2000 / scalerange) - ((offset * (sr / 44100) * 20000 / scalerange)), 8))
    Next n
    
End If



If sr = 44100 And logflag = 0 Then

    Base.CurrentY = 442
    Base.CurrentX = 24
    Base.Print "30"

    Base.CurrentY = 442
    Base.CurrentX = 77
    Base.Print "100"
    
    Base.CurrentY = 442
    Base.CurrentX = 185
    Base.Print "1000"
    
    Base.CurrentY = 442
    Base.CurrentX = 288
    Base.Print "10000"
    
End If

If sr = 22050 And logflag = 0 Then

    Base.CurrentY = 442
    Base.CurrentX = 24
    Base.Print "15"
    
    Base.CurrentY = 442
    Base.CurrentX = 77
    Base.Print "50"
    
    Base.CurrentY = 442
    Base.CurrentX = 185
    Base.Print "500"
    
    Base.CurrentY = 442
    Base.CurrentX = 288
    Base.Print "5000"
    
End If

If sr = 11025 And logflag = 0 Then

    Base.CurrentY = 442
    Base.CurrentX = 24
    Base.Print "7.5"

    Base.CurrentY = 442
    Base.CurrentX = 77
    Base.Print "25"
    
    Base.CurrentY = 442
    Base.CurrentX = 185
    Base.Print "250"
    
    Base.CurrentY = 442
    Base.CurrentX = 288
    Base.Print "2500"
    
End If
    
    
    Base.CurrentY = 30
    Base.CurrentX = 8
    Base.Print CStr(dbscale(0) + ref)

    Base.CurrentY = 130
    Base.CurrentX = 8
    Base.Print CStr(dbscale(1) + ref)
    
    Base.CurrentY = 230
    Base.CurrentX = 8
    Base.Print CStr(dbscale(2) + ref)
    
    Base.CurrentY = 330
    Base.CurrentX = 8
    Base.Print CStr(dbscale(3) + ref)

    Base.CurrentY = 430
    Base.CurrentX = 8
    Base.Print CStr(dbscale(4) + ref)
    
    For tick = 0 To 40
    Base.CurrentY = 25 + (tick * 10)
    Base.CurrentX = 21
    Base.Print "__"
    Next tick
    
     For tick = 0 To 8
    Base.CurrentY = 25 + (tick * 50)
    Base.CurrentX = 18
    Base.Print "___"
    Next tick
    
    
    Scope.ForeColor = &HBBBBBB      'draw dB gridlines
    ScopeBuff.ForeColor = &HBBBBBB
    For tick = 0 To 40
    
            If (tick - 1) / 5 = Int((tick - 1) / 5) Then
            ScopeBuff.ForeColor = &H888888
            Scope.ForeColor = &H888888
            Else
            ScopeBuff.ForeColor = &HBBBBBB
            Scope.ForeColor = &HBBBBBB
            End If
    
    Scope.Line (0, ((tick * 10) - 7))-(724, ((tick * 10) - 7))
    ScopeBuff.Line (0, ((tick * 10) - 7))-(724, ((tick * 10) - 7))
    Next tick
    Scope.ForeColor = vbBlack
    ScopeBuff.ForeColor = vbBlack
    
    
     If logflag = 0 Then    'print log freq lines
         
            Scope.ForeColor = &HAAAAAA
            ScopeBuff.ForeColor = &HAAAAAA
        
        ss = 11.966      ' 11.966
        
        For n = 1 To 3
  
  nt = 10 ^ n
  
For nnn = 1 To 10

nn = nnn * nt

        If nn < 101 / ss And nn > 10 / ss Then 'multiplier freq 10-100 Hz
Factor2(nn) = ((2606 / ss) * (Log(nn / 10))) + 1
End If

If nn < 1001 / ss And nn > 100 / ss Then 'multiplier freq 100-1000 Hz
Factor2(nn) = ((2606 / ss) * (Log(nn / (100)))) + (6000 / ss)
End If

If nn < 10001 / ss And nn > 1000 / ss Then 'multiplier freq 1000-10000 Hz
Factor2(nn) = ((2606 / ss) * (Log(nn / (1000)))) + (12000 / ss)
End If

If nn < 100001 / ss And nn > 10000 / ss Then 'multiplier freq 10000-20000 Hz
Factor2(nn) = ((2606 / ss) * (Log(nn / (10000)))) + (18000 / ss)
End If

If nn < 1000001 / ss And nn > 100000 / ss Then 'multiplier freq >20000 Hz
Factor2(nn) = ((2606 / ss) * (Log(nn / (100000)))) + (24000 / ss)
End If


ScopeBuff.Line ((Factor2(nn) / (freqscale * 0.98)) - 121, 0)-((Factor2(nn) / (freqscale * 0.98)) - 121, 402)
Scope.Line ((Factor2(nn) / (freqscale * 0.98)) - 121, 0)-((Factor2(nn) / (freqscale * 0.98)) - 121, 402)

Next nnn


Next n
        
            Scope.ForeColor = vbBlack
            ScopeBuff.ForeColor = vbBlack
        
        End If
    
    
    
    
 If logflag = 1 Then   'draw linear freq scale lines

Scope.ForeColor = &HAAAAAA
ScopeBuff.ForeColor = &HAAAAAA

For tick = 1 To 11
ScopeBuff.Line (((Int(tick * 71.9)) - 1), 0)-(((Int(tick * 71.9)) - 1), 402)
Scope.Line (((Int(tick * 71.9)) - 1), 0)-(((Int(tick * 71.9)) - 1), 402)
Next tick

Scope.ForeColor = vbBlack
ScopeBuff.ForeColor = vbBlack
End If
    
    
    
    
    
    ss = 9.8  '11.966
    
For n = 1 To 3
  
  nt = 10 ^ n
  
For nnn = 1 To 10

nn = nnn * nt


    If nn < 101 / ss And nn > 10 / ss Then 'multiplier freq 10-100 Hz
Factor2(nn) = ((2606 / ss) * (Log(nn / 10))) + 1
End If

If nn < 1001 / ss And nn > 100 / ss Then 'multiplier freq 100-1000 Hz
Factor2(nn) = ((2606 / ss) * (Log(nn / (100)))) + (6000 / ss)
End If

If nn < 10001 / ss And nn > 1000 / ss Then 'multiplier freq 1000-10000 Hz
Factor2(nn) = ((2606 / ss) * (Log(nn / (1000)))) + (12000 / ss)
End If

If nn < 100001 / ss And nn > 10000 / ss Then 'multiplier freq 10000-20000 Hz
Factor2(nn) = ((2606 / ss) * (Log(nn / (10000)))) + (18000 / ss)
End If

vertline(nnn + ((n - 1) * 10)) = Factor2(nn) / 1.23

Next nnn

Next n




If logflag = 0 Then

 For tick = 3 To 39
 
    Base.CurrentY = 433
    Base.CurrentX = -25 + Int(vertline(tick) / 4.7)
    Base.Print "I"
    Next tick
        
End If

If logflag = 1 Then
For tick = 1 To 11
Base.CurrentY = 433
Base.CurrentX = -6 + Int(tick * 29.9)
Base.Print "I"
Next tick
End If
    
Text3.Text = CStr(offset)
Text2.Text = CStr(scalerange)
    
End Function


Private Function doplot()  'for printer



    ' Get the printer's dimensions in twips.
    pwid = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbTwips)
    phgt = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbTwips)
    
    ' Convert the printer's dimensions into the
    ' object's coordinates.
    pwid = Scope.ScaleX(pwid, vbTwips, Scope.ScaleMode)
    phgt = Scope.ScaleY(phgt, vbTwips, Scope.ScaleMode)
    
    ' Compute the center of the object.
    xmid = Scope.ScaleLeft + Scope.ScaleWidth / 2
    ymid = Scope.ScaleTop + Scope.ScaleHeight / 2
    
    ' Pass the coordinates of the upper left and
    ' lower right corners into the Scale method.
    Printer.Scale _
        (xmid - pwid / 2, ymid - phgt / 2)- _
        (xmid + pwid / 2, ymid + phgt / 2)





If logflag = 0 Then  'if logscale

For X = 1 To 10  'make freq lines 10-100 Hz
If X = 1 Then
Printer.DrawWidth = 2
Else: Printer.DrawWidth = 1
End If
factor1 = Int(((2606 * (Log(X))) + 1) * 724 / 22050)   '884
Printer.Line (factor1, 0)-(factor1, 400)
Next X
For X = 1 To 10 'make freq lines 100-1000 Hz
factor1 = Int(((2606 * (Log(X))) + 6000) * 724 / 22050)
Printer.Line (factor1, 0)-(factor1, 400)
Next X
For X = 1 To 10 'make freq lines 1000-10000 Hz
factor1 = Int(((2606 * (Log(X))) + 12000) * 724 / 22050)
Printer.Line (factor1, 0)-(factor1, 400)
Next X
For X = 1 To 2  'make freq lines 10000-20000 Hz
If X = 2 Then
Printer.DrawWidth = 2
Else: Printer.DrawWidth = 1
End If
factor1 = Int(((2606 * (Log(X))) + 18000) * 724 / 22050)
Printer.Line (factor1, 0)-(factor1, 400)
Next X

End If





If logflag = 1 Then   'linear scaling
For X = 0 To 10
Printer.Line ((X * 72.4 * 10 / 11), 0)-((X * 72.4 * 10 / 11), 400)
Next X
End If



For y = 0 To 40   'make dB lines
If y / 10 = Int(y / 10) Then
Printer.DrawWidth = 2
Printer.Line (0, (y * 10))-(724 * (20000 / 22050), (y * 10))
Else
Printer.DrawWidth = 1
Printer.Line (0, (y * 10))-(724 * (20000 / 22050), (y * 10))
End If
Next y

Printer.DrawWidth = 1





For X = 0 To 4   'print db numbers

Printer.CurrentY = (X * 100) - 5
Printer.CurrentX = -30
texttoprint = CStr(dbscale(X) + ref)
Printer.Print texttoprint
Next X

Printer.CurrentY = -20
Printer.CurrentX = 200
texttoprint = Text14.Text  'print legend
Printer.Print texttoprint
Printer.CurrentY = -50
Printer.CurrentX = 300
texttoprint = Now
Printer.Print texttoprint



If logflag = 0 Then  'logscale numbering

Printer.CurrentY = 405   'print freq numbers
Printer.CurrentX = -15
texttoprint = CStr((sr / 11025) * 2.5)
Printer.Print texttoprint

Printer.CurrentY = 405
Printer.CurrentX = Int((5800 / 22000) * 724)
texttoprint = CStr((sr / 11025) * 25)
Printer.Print texttoprint

Printer.CurrentY = 405
Printer.CurrentX = Int((11800 / 22000) * 724)
texttoprint = CStr((sr / 11025) * 250)
Printer.Print texttoprint

Printer.CurrentY = 405
Printer.CurrentX = Int((17800 / 22000) * 724)
texttoprint = CStr((sr / 11025) * 2500)
Printer.Print texttoprint

Printer.CurrentY = 420
Printer.CurrentX = 350
texttoprint = "Hertz"
Printer.Print texttoprint

End If

'CStr(X * (sr / 44100) * (2000) / scalerange)
'CStr(Left$((n * 2000 / scalerange) - ((offset * 20000 / scalerange)), 8))

If logflag = 1 Then   'linear scale numbering
For X = 0 To 10
Printer.CurrentY = 405
Printer.CurrentX = (Int(X * 72.4 * 10 / 11)) - 5
texttoprint = CStr(Left$((X * (sr / 44100) * 2000 / scalerange) - ((offset * (sr / 44100) * 20000 / scalerange)), 8))
Printer.Print texttoprint
Printer.CurrentY = 420
Printer.CurrentX = 350
texttoprint = "Hertz"
Printer.Print texttoprint
Next X

End If




For X = 4 To 7401

freq(X) = X * (22050 / 8160)

If freq(X) < 101 And freq(X) > 10 Then  'find log factor 10-100 Hz
Factor2(X) = Int(2606 * (Log(freq(X) / 10))) + 1
End If

If freq(X) < 1001 And freq(X) > 100 Then 'find log factor 100-1000 Hz
Factor2(X) = Int(2606 * (Log(freq(X) / 100))) + 6000
End If

If freq(X) < 10001 And freq(X) > 1000 Then 'find log factor 1000-10000 Hz
Factor2(X) = Int(2606 * (Log(freq(X) / 1000))) + 12000
End If

If freq(X) < 20001 And freq(X) > 10000 Then  'find log factor 10000-20000 Hz
Factor2(X) = Int(2606 * (Log(freq(X) / 10000))) + 18000
End If

'ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale),
' ScopeHeight + 75 + (ref / 0.25) - (8 * (logvol2(X))))


Factor2(X) = Int((Factor2(X) / (22050)) * 724)


'Factor2(X) = (X * 784 * scalerange / 4080) + (784 * offset * (20 / 21.97) / 0.5)

 If logflag = 1 Then    'linear plot
 Factor2(X) = (offset * 724 / 1.1) + ((X * scalerange) / 5.189) * (0.459)
 End If




If Factor2(X) < 0 Then
Factor2(X) = 0
End If

If Factor2(X) > 655 Then
Factor2(X) = 655
End If


'ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale),
' ScopeHeight + 75 + (ref / 0.25) - (8 * (logvol2(X))))

If X > 1 Then
Printer.DrawWidth = 1

'magnitude calc



             If peakflag = 0 And Divisor = 10000 Then
             lplotold = 475 - Int(((logvol1(X - 1)) * 8) - (ref / 0.25))
             lplotnew = 475 - Int(((logvol1(X)) * 8) - (ref / 0.25))
             End If
             
             If peakflag = 1 And Divisor = 10000 Then
             lplotold = 475 - Int(((logvol2(X - 1)) * 8) - (ref / 0.25))
             lplotnew = 475 - Int(((logvol2(X)) * 8) - (ref / 0.25))
             End If
            
            If peakflag = 0 And Divisor = 1000 Then
             lplotold = 475 - Int(((logvol1(X - 1)) * 10) - (ref / 0.2))
             lplotnew = 475 - Int(((logvol1(X)) * 10) - (ref / 0.2))
             End If
             
             If peakflag = 1 And Divisor = 1000 Then
             lplotold = 475 - Int(((logvol2(X - 1)) * 10) - (ref / 0.2))
             lplotnew = 475 - Int(((logvol2(X)) * 10) - (ref / 0.2))
             End If
            
           If peakflag = 0 And Divisor = 100000 Then
             lplotold = 450 - Int(((logvol1(X - 1)) * 6.6666) - (ref / 0.3))
             lplotnew = 450 - Int(((logvol1(X)) * 6.6666) - (ref / 0.3))
             End If
             
             If peakflag = 1 And Divisor = 100000 Then
             lplotold = 450 - Int(((logvol2(X - 1)) * 6.6666) - (ref / 0.3))
             lplotnew = 450 - Int(((logvol2(X)) * 6.6666) - (ref / 0.3))
             End If
            
            If peakflag = 0 And Divisor = 100 Then
             lplotold = 550 - Int(((logvol1(X - 1)) * 13.3333) - (ref / 0.15))
             lplotnew = 550 - Int(((logvol1(X)) * 13.3333) - (ref / 0.15))
             End If
             
             If peakflag = 1 And Divisor = 100 Then
             lplotold = 550 - Int(((logvol2(X - 1)) * 13.3333) - (ref / 0.15))
             lplotnew = 550 - Int(((logvol2(X)) * 13.3333) - (ref / 0.15))
             End If
            
            If peakflag = 0 And Divisor = 10 Then
             lplotold = 575 - Int(((logvol1(X - 1)) * 20) - (ref / 0.1))
             lplotnew = 575 - Int(((logvol1(X)) * 20) - (ref / 0.1))
             End If
             
             If peakflag = 1 And Divisor = 10 Then
             lplotold = 575 - Int(((logvol2(X - 1)) * 20) - (ref / 0.1))
             lplotnew = 575 - Int(((logvol2(X)) * 20) - (ref / 0.1))
             End If
            
            If peakflag = 0 And Divisor = 1 Then
             lplotold = 750 - Int(((logvol1(X - 1)) * 40) - (ref / 0.05))
             lplotnew = 750 - Int(((logvol1(X)) * 40) - (ref / 0.05))
             End If
             
             If peakflag = 1 And Divisor = 1 Then
             lplotold = 750 - Int(((logvol2(X - 1)) * 40) - (ref / 0.05))
             lplotnew = 750 - Int(((logvol2(X)) * 40) - (ref / 0.05))
             End If
            
            If peakflag = 0 And Divisor = 1000000 Then
             lplotold = 437 - Int(((logvol1(X - 1)) * 5) - (ref / 0.35))
             lplotnew = 437 - Int(((logvol1(X)) * 5) - (ref / 0.35))
             End If
             
             If peakflag = 1 And Divisor = 1000000 Then
             lplotold = 437 - Int(((logvol2(X - 1)) * 5) - (ref / 0.35))
             lplotnew = 437 - Int(((logvol2(X)) * 5) - (ref / 0.35))
             End If
            
            If peakflag = 0 And Divisor = 10000000 Then
             lplotold = 425 - Int(((logvol1(X - 1)) * 4) - (ref / 0.4))
             lplotnew = 425 - Int(((logvol1(X)) * 4) - (ref / 0.4))
             End If
             
             If peakflag = 1 And Divisor = 10000000 Then
             lplotold = 425 - Int(((logvol2(X - 1)) * 4) - (ref / 0.4))
             lplotnew = 425 - Int(((logvol2(X)) * 4) - (ref / 0.4))
             End If
            
End If


If lplotold > 400 Then  'prevent off graph plotting
lplotold = 400
End If

If lplotnew > 400 Then
lplotnew = 400
End If

If lplotold < 0 Then
lplotold = 0
End If

If lplotnew < 0 Then
lplotnew = 0
End If

If Factor2(X) < 0 Then
Factor2(X) = 0
End If

If Factor2(X) > 724 Then
Factor2(X) = 724
End If

If Factor2(X - 1) > 724 Then
Factor2(X - 1) = 724
End If

If Factor2(X - 1) < 0 Then
Factor2(X - 1) = 0
End If


If plotwidth = 1 Then
Printer.DrawWidth = 2
Else
Printer.DrawWidth = 1
End If

Printer.Line ((Factor2(X) / 1), lplotnew)-(((Factor2(X - 1)) / 1), lplotold), vbBlack

Printer.DrawWidth = 1

Next X




Printer.EndDoc







End Function




Private Function replot()

    'ScopeBuff.ForeColor = vbBlack
    'Scope.ForeColor = vbBlack
    
    
    If calflag = 1 Then
    caloffset2 = caloffset
    Else
    caloffset2 = 0
    End If
    
    
    
    For X = 0 To NumSamples - 1
     
     If hamflag = 1 Then
     hamx(X) = 0.54 - (0.54 * (Cos((Twopi * 2 * X) / NumSamples)))
     End If
     If blackflag = 1 Then
     hamx(X) = 0.42 - (0.5 * Cos((Twopi * 2 * X) / NumSamples)) + (0.08 * Cos(Twopi * 4 * X / NumSamples))
     End If
     If hamflag = 0 And blackflag = 0 Then
     hamx(X) = 1
     End If
     
     
    Next X
    
    If loadflag = 0 Then 'new data capture
            
          
             'PARSE STEREO DATAFILE HERE
           
         If channels = 0 Or channels = 1 Then
         For X = 0 To (NumSamples - 3) / 2
         InData2(X) = (InData((X * 2) + channels) * hamx(X * 2))   'stereo +0 = left, +1 = right
         Next X
         
         For X = (NumSamples - 3) / 2 To (NumSamples - 3)
         InData2(X) = (InData((X * 2) + channels) * hamx(X / 2))
         Next X
         End If
            
            
         If channels = 3 Then
         For X = 0 To (NumSamples - 1)   'mono
         InData2(X) = (InData(X) * hamx(X))
         Next X
         End If
            
         
            
            Call FFTAudio(InData2, OutData)
              
              
            
            If aveflag = 1 Then
            
            avetimes = avetimes + 1
            
            If avetimes > 2000000 Then
            avetimes = 1
            End If
            
            If avetimes = 1 Then
            For X = 1 To NumSamples - 1
            OutData2(X) = OutData(X)
            Next X
            End If
            
            For X = 1 To NumSamples - 1
             
            OutData2(X) = (Abs(OutData2(X) * avetimes + 1) + Abs(OutData(X))) / (avetimes + 1)
            OutData(X) = OutData2(X)
            
            Next X
            
            End If
           
            Scope.Cls
            ScopeBuff.Cls
   End If   'for new data capture
            
            
            
            
            
            
           
            
            
            
            
            Scope.CurrentX = 0
            Scope.CurrentY = ScopeHeight
        
            ScopeBuff.CurrentX = 0
            ScopeBuff.CurrentY = ScopeHeight
        
        
        
        
             Scope.ForeColor = &HAAAAAA               'draw dB gridlines
             ScopeBuff.ForeColor = &HAAAAAA
             For tick = 0 To 40
              
              If (tick - 1) / 5 = Int((tick - 1) / 5) Then
            ScopeBuff.ForeColor = &H888888
            Else
            ScopeBuff.ForeColor = &HBBBBBB
            End If
             
             Scope.Line (0, ((tick * 10) - 7))-(724, ((tick * 10) - 7))
             ScopeBuff.Line (0, ((tick * 10) - 7))-(724, ((tick * 10) - 7))
             Next tick
             
             
             
             If loadflag = 0 Then
             Scope.ForeColor = vbBlack
             ScopeBuff.ForeColor = vbBlack
             Else
             Scope.ForeColor = vbBlue
             ScopeBuff.ForeColor = vbBlue
             End If
        
         If logflag = 0 Then    'print log freq lines
         
            Scope.ForeColor = &HAAAAAA
            ScopeBuff.ForeColor = &HAAAAAA
        
        For n = 1 To 3
  
  nt = 10 ^ n
  
For nnn = 1 To 10

nn = nnn * nt

        If nn < 101 / 11.966 And nn > 10 / 11.966 Then 'multiplier freq 10-100 Hz
Factor2(nn) = ((2606 / 11.966) * (Log(nn / 10))) + 1
End If

If nn < 1001 / 11.966 And nn > 100 / 11.966 Then 'multiplier freq 100-1000 Hz
Factor2(nn) = ((2606 / 11.966) * (Log(nn / (100)))) + (6000 / 11.966)
End If

If nn < 10001 / 11.966 And nn > 1000 / 11.966 Then 'multiplier freq 1000-10000 Hz
Factor2(nn) = ((2606 / 11.966) * (Log(nn / (1000)))) + (12000 / 11.966)
End If

If nn < 100001 / 11.966 And nn > 10000 / 11.966 Then 'multiplier freq 10000-20000 Hz
Factor2(nn) = ((2606 / 11.966) * (Log(nn / (10000)))) + (18000 / 11.966)
End If

If nn < 1000001 / 11.966 And nn > 100000 / 11.966 Then 'multiplier freq >20000 Hz
Factor2(nn) = ((2606 / 11.966) * (Log(nn / (100000)))) + (24000 / 11.966)
End If


ScopeBuff.Line ((Factor2(nn) / (freqscale * 0.98)) - 121, 0)-((Factor2(nn) / (freqscale * 0.98)) - 121, 402)
Scope.Line ((Factor2(nn) / (freqscale * 0.98)) - 121, 0)-((Factor2(nn) / (freqscale * 0.98)) - 121, 402)


Next nnn

Next n
        
            If loadflag = 0 Then
             Scope.ForeColor = vbBlack
             ScopeBuff.ForeColor = vbBlack
             Else
             Scope.ForeColor = vbBlue
             ScopeBuff.ForeColor = vbBlue
             End If
        
        End If
        
        
        
        
         If logflag = 1 Then   'draw linear freq scale lines

Scope.ForeColor = &HAAAAAA
ScopeBuff.ForeColor = &HAAAAAA


For tick = 1 To 11
ScopeBuff.Line (((Int(tick * 71.9)) - 1), 0)-(((Int(tick * 71.9)) - 1), 402)
Scope.Line (((Int(tick * 71.9)) - 1), 0)-(((Int(tick * 71.9)) - 1), 402)
Next tick

             If loadflag = 0 Then
             Scope.ForeColor = vbBlack
             ScopeBuff.ForeColor = vbBlack
             Else
             Scope.ForeColor = vbBlue
             ScopeBuff.ForeColor = vbBlue
             End If
             
             
End If
    
            Scope.CurrentX = 0
            Scope.CurrentY = ScopeHeight
            ScopeBuff.CurrentX = 0
            ScopeBuff.CurrentY = ScopeHeight
        
            For X = 1 To 8160           '4080
                '.CurrentY = ScopeHeight
               
               
                
If X < 101 / 11.966 And X > 10 / 11.966 Then 'multiplier freq 10-100 Hz
Factor2(X) = ((2606 / 11.966) * (Log(X / 10))) + 1
End If

If X < 1001 / 11.966 And X > 100 / 11.966 Then 'multiplier freq 100-1000 Hz
Factor2(X) = ((2606 / 11.966) * (Log(X / (100)))) + (6000 / 11.966)
End If

If X < 10001 / 11.966 And X > 1000 / 11.966 Then 'multiplier freq 1000-10000 Hz
Factor2(X) = ((2606 / 11.966) * (Log(X / (1000)))) + (12000 / 11.966)
End If

If X < 100001 / 11.966 And X > 10000 / 11.966 Then 'multiplier freq 10000-20000 Hz
Factor2(X) = ((2606 / 11.966) * (Log(X / (10000)))) + (18000 / 11.966)
End If


                
 If logflag = 1 Then
 
 
 
 Factor2(X) = (X * 784 * scalerange / 4080) + (784 * offset * (20 / 22.05) / 0.5)
 
 
 End If
               
               
        If blackflag = 1 Then
        OutData(X) = OutData(X) / 0.38
        End If
        
        If hamflag = 1 Then
        OutData(X) = OutData(X) / 0.5
        End If
        
        
        
        
            'I average two elements here because it gives a smoother appearance.
            
            If (((Sqr(Abs(OutData(X) / 1000))))) > 0 Then
            logvol1(X) = (Log(((Sqr(Abs((OutData(X) + OutData(X + 1)) / 1000))))) / 2.3) * 20 + (caloffset2 / 2.3)
            
            End If
            
            If logvol1(X) > logvol2(X) Then
            logvol2(X) = logvol1(X)
            End If
            
           
            
             If peakflag = 0 And Divisor = 10000000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 25 + (ref / 0.4) - (4 * (logvol1(X))))
             End If
             
             If peakflag = 1 And Divisor = 10000000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 25 + (ref / 0.4) - (4 * (logvol2(X))))
             End If
            
             If peakflag = 0 And Divisor = 1000000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 37 + (ref / 0.35) - (5 * (logvol1(X))))
             End If
             
             If peakflag = 1 And Divisor = 1000000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 37 + (ref / 0.35) - (5 * (logvol2(X))))
             End If
            
            
             If peakflag = 0 And Divisor = 100000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 50 + (ref / 0.3) - (6.6666 * (logvol1(X))))
             End If
             
             If peakflag = 1 And Divisor = 100000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 50 + (ref / 0.3) - (6.6666 * (logvol2(X))))
             End If
            
            
            If peakflag = 0 And Divisor = 1000 Then
            ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 75 + (ref / 0.2) - (10 * (logvol1(X))))
            End If
             
             If peakflag = 1 And Divisor = 1000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 75 + (ref / 0.2) - (10 * (logvol2(X))))
             End If
             
            If peakflag = 0 And Divisor = 10000 Then
            ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 75 + (ref / 0.25) - (8 * (logvol1(X))))
            End If
             
             If peakflag = 1 And Divisor = 10000 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 75 + (ref / 0.25) - (8 * (logvol2(X))))
             End If
             
            If peakflag = 0 And Divisor = 100 Then
            ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 150 + (ref / 0.15) - (13.333 * (logvol1(X))))
            End If
             
             If peakflag = 1 And Divisor = 100 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 150 + (ref / 0.15) - (13.333 * (logvol2(X))))
             End If
            
            If peakflag = 0 And Divisor = 10 Then
            ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 175 + (ref / 0.1) - (20 * (logvol1(X))))
            End If
             
             If peakflag = 1 And Divisor = 10 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 175 + (ref / 0.1) - (20 * (logvol2(X))))
             End If
             
             If peakflag = 0 And Divisor = 1 Then
            ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 350 + (ref / 0.05) - (40 * (logvol1(X))))
            End If
             
             If peakflag = 1 And Divisor = 1 Then
             ScopeBuff.Line Step(0, 0)-(((Factor2(X)) / freqscale), ScopeHeight + 350 + (ref / 0.05) - (40 * (logvol2(X))))
             End If
             
             
             
             Next X
            
            Scope.Picture = ScopeBuff.Image 'Display the double-buffer
          
        
 
         ScopeBuff.ForeColor = vbBlack
         Scope.ForeColor = vbBlack
 
 loadflag = 0
 

End Function
