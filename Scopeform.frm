VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Scopeform 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000A&
   Caption         =   "Scope"
   ClientHeight    =   4215
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   1800
      TabIndex        =   18
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   1080
      TabIndex        =   17
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   2760
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "R"
      Height          =   495
      Left            =   1680
      TabIndex        =   15
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "M"
      Height          =   495
      Left            =   960
      TabIndex        =   14
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "L"
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "COLOR"
      Height          =   375
      Left            =   1320
      TabIndex        =   12
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GRID"
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Text            =   "Oscilloscope Plot from Spectrum Analyzer"
      Top             =   0
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   3015
      Left            =   2760
      ScaleHeight     =   2955
      ScaleWidth      =   4155
      TabIndex        =   6
      Top             =   360
      Width           =   4215
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00C0C0C0&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00C0C0C0&
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4560
      TabIndex        =   1
      Text            =   "4"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox Text20 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7080
      TabIndex        =   0
      Text            =   "1"
      Top             =   1560
      Width           =   495
   End
   Begin MSComDlg.CommonDialog dlgfile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
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
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   3840
      Width           =   135
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   7
      Top             =   1320
      Width           =   255
   End
   Begin VB.Menu menuLoad 
      Caption         =   "File"
   End
   Begin VB.Menu menuPrint 
      Caption         =   "Print"
   End
   Begin VB.Menu menuExit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Scopeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  DazyWeb Laboratories VB-3000 Spectrum Analyzer Scope Module  v1.0   build May 30, 2001
'      copyright 2001
'
Option Explicit

Dim fname As String
Dim fnum1 As Integer
Dim cntr1 As Integer
Dim aveflag As Integer
Dim legend As String
Dim displayflag As String
Dim gridflag As Integer
Dim colorflag As Integer
Dim linecolor
Dim v As Single
Dim vr As String
Dim scopedata(19384) As Long
Dim scopedata2(19384) As Long
Dim data_L(10000) As Long
Dim data_R(10000) As Long
Dim scalar As Single
Dim n As Integer
Dim lvlscalar As Single
Dim logvol1(8160) As Variant
Dim sr As Integer
Dim loadedsamples As Long
Dim X As Long
Dim pwid As Single
Dim phgt As Single
Dim xmid As Single
Dim ymid As Single
Dim factor1 As Long
Dim y As Long
Dim lplotold As Long
Dim lplotnew As Long
Dim texttoprint As String
Dim Y1 As Double
Dim Y2 As Double
Dim channels As Integer
Dim channelflag As String



Public Function doscope()


If channels = 1 Or channels = 0 Then  'parse for stereo data
cntr1 = 0

For n = 0 To 16381 / 2
data_R(cntr1) = scopedata(n * 2)
data_L(cntr1) = scopedata((n * 2) + 1)
cntr1 = cntr1 + 1
Next n

End If 'end of stereo data parse



If channels = 1 Then 'left
For n = 0 To (16383 / 2)
scopedata2(n) = data_L(n)
Next n
For n = (16383 / 2) To 16383
scopedata2(n) = scopedata2(n / 2)
Next n
End If


If channels = 0 Then 'right
For n = 0 To (16383 / 2)
scopedata2(n) = data_R(n)
Next n
For n = (16383 / 2) To 16383
scopedata2(n) = scopedata2(n / 2)
Next n
End If


If channels = 2 Then  'force sum left and right in stereo
For n = 0 To (16383 / 2)
scopedata2(n) = data_L(n) + data_R(n)
Next n
For n = (16383 / 2) To 16383
scopedata2(n) = scopedata2(n / 2)
Next n

End If




 Picture1.Cls


If colorflag = 1 Then
Picture1.BackColor = vbBlack
linecolor = vbGreen
Else
Picture1.BackColor = vbWhite
linecolor = vbBlack
End If


If gridflag = 1 Then

Picture1.ForeColor = &HAAAAAA
  
For n = 1 To 9
Picture1.Line (0, (n * 295.5))-(4155, (n * 295.5))
Picture1.Line ((n * 415.5), 0)-((n * 415.5), 2955)
Next n
  
Picture1.ForeColor = vbBlack
End If
  
  
  
    For n = 1 To Int(8190 / scalar)
    
    Picture1.Line ((Int((n * scalar))), (Int(1500 - (lvlscalar * (scopedata2(n) / 3)))))-((Int(((n + 1) * scalar))), (Int(1500 - (lvlscalar * (scopedata2(n + 1) / 3))))), linecolor
      
    Next n


End Function

Private Sub Command1_Click()  'GRID TOGGLE

If gridflag = 0 Then
gridflag = 1
Else
gridflag = 0
End If

doscope

End Sub

Private Sub Command2_Click() 'DISPLAY COLOR

If colorflag = 0 Then
colorflag = 1
Else
colorflag = 0
End If

doscope

End Sub

Private Sub Command3_Click()  'Display left channel

displayflag = "L"
Text2.BackColor = vbRed
Text4.BackColor = vbBlue
Text3.BackColor = vbBlue
channels = 1
doscope
End Sub

Private Sub Command4_Click()  'Display mono channel

displayflag = "M"
Text2.BackColor = vbBlue
Text4.BackColor = vbBlue
Text3.BackColor = vbRed
If channels <> 3 Then
channels = 2
End If
doscope
End Sub

Private Sub Command5_Click()  'Display right channel

displayflag = "R"
Text2.BackColor = vbBlue
Text4.BackColor = vbRed
Text3.BackColor = vbBlue
channels = 0
doscope
End Sub

Private Sub menuExit_Click() 'EXIT

fname = "c:/vb3scopeinit"
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Output As #fnum1
On Error Resume Next

Write #fnum1, scalar
Write #fnum1, lvlscalar
Write #fnum1, gridflag
Write #fnum1, colorflag
Write #fnum1, displayflag

Close fnum1
Unload Scopeform

End Sub

Private Sub Form_Load()  'scope

scalar = 4
lvlscalar = 1

fname = "c:/vb3scopeinit"
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Input As #fnum1
On Error Resume Next

Input #fnum1, scalar
Input #fnum1, lvlscalar
Input #fnum1, gridflag
Input #fnum1, colorflag
Input #fnum1, displayflag

Close fnum1


If displayflag = "L" Then
Text2.BackColor = vbRed
Text4.BackColor = vbBlue
Text3.BackColor = vbBlue
End If


If displayflag = "R" Then
Text2.BackColor = vbBlue
Text4.BackColor = vbRed
Text3.BackColor = vbBlue
End If


If displayflag = "M" Then
Text2.BackColor = vbBlue
Text4.BackColor = vbBlue
Text3.BackColor = vbRed
End If




fname = "c:/wavscopedata"
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Input As #fnum1
On Error Resume Next

For n = 0 To 16383
Input #fnum1, scopedata(n)
Next n
Input #fnum1, channels
Close fnum1

If channels = 3 Then
Text2.BackColor = vbBlue
Text4.BackColor = vbBlue
Text3.BackColor = vbRed
Command3.Enabled = False
Command4.Enabled = True
Command5.Enabled = False
End If

If channels = 0 Then
Text2.BackColor = vbBlue
Text4.BackColor = vbRed
Text3.BackColor = vbBlue
channelflag = "R"
End If

If channels = 1 Then
Text2.BackColor = vbRed
Text4.BackColor = vbBlue
Text3.BackColor = vbBlue
channelflag = "L"
End If



'add scale lines here

For n = 0 To 10
Scopeform.CurrentY = 18 + (n * 20)
Scopeform.CurrentX = 155
v = (5 - n)
If v > 0 Then
vr = "+" + CStr(v)
Else
vr = " " + CStr(v)
End If
Scopeform.Print vr + "   --"
    
Scopeform.CurrentY = 235
Scopeform.CurrentX = 183 + (n * 27.4)
Scopeform.Print CStr(n)
Scopeform.Line ((187 + (n * 27.5)), 230)-((187 + (n * 27.5)), 222)
Next n


doscope


End Sub


Private Sub Command14_Click()  'scalar up

scalar = scalar * 2
If scalar > 128 Then
scalar = 128
End If
doscope
Text19.Text = CStr(scalar)
End Sub

Private Sub Command13_Click() 'scalar down

scalar = scalar / 2
If scalar < 0.5 Then
scalar = 0.5
End If
doscope
Text19.Text = CStr(scalar)
End Sub

Private Sub Command15_Click() 'scope level up

lvlscalar = lvlscalar * 2
If lvlscalar > 256 Then
lvlscalar = 256
End If
doscope
Text20.Text = CStr(lvlscalar)

End Sub

Private Sub Command16_Click()  'scope level down

lvlscalar = lvlscalar / 2
If lvlscalar < 0.0625 Then
lvlscalar = 0.0625
End If
doscope
Text20.Text = CStr(lvlscalar)

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


    
fnum1 = FreeFile
    
On Error Resume Next
Open fname For Input As #fnum1
On Error Resume Next

Input #fnum1, sr
Input #fnum1, loadedsamples

For n = 1 To 8160
Input #fnum1, logvol1(n)
Next n

For n = 0 To 16383
Input #fnum1, scopedata(n)
Next n
Input #fnum1, aveflag
Input #fnum1, legend
Input #fnum1, channels
Close fnum1



If channels = 3 Then
Text2.BackColor = vbBlue
Text4.BackColor = vbBlue
Text3.BackColor = vbRed
Command3.Enabled = False
Command4.Enabled = True
Command5.Enabled = False
End If

If channels = 0 Then
Text2.BackColor = vbBlue
Text4.BackColor = vbRed
Text3.BackColor = vbBlue
channelflag = "R"
End If

If channels = 1 Then
Text2.BackColor = vbRed
Text4.BackColor = vbBlue
Text3.BackColor = vbBlue
channelflag = "L"
End If

doscope

End Sub




Private Sub menuPrint_Click()



If channels = 1 Or channels = 0 Then  'parse for stereo data
cntr1 = 0

For n = 0 To 16381 / 2
data_R(cntr1) = scopedata(n * 2)
data_L(cntr1) = scopedata((n * 2) + 1)
cntr1 = cntr1 + 1
Next n

End If 'end of stereo data parse



If channels = 1 Then 'left
For n = 0 To (16383 / 2)
scopedata2(n) = data_L(n)
Next n
For n = (16383 / 2) To 16383
scopedata2(n) = scopedata2(n / 2)
Next n
End If


If channels = 0 Then 'right
For n = 0 To (16383 / 2)
scopedata2(n) = data_R(n)
Next n
For n = (16383 / 2) To 16383
scopedata2(n) = scopedata2(n / 2)
Next n
End If


If channels = 2 Then 'force sum left and right in stereo
For n = 0 To (16383 / 2)
scopedata2(n) = data_L(n) + data_R(n)
Next n
For n = (16383 / 2) To 16383
scopedata2(n) = scopedata2(n / 2)
Next n

End If



doplot


End Sub



Private Function doplot()  'for printer



    ' Get the printer's dimensions in twips.
    pwid = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbTwips)
    phgt = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbTwips)
    
    ' Convert the printer's dimensions into the
    ' object's coordinates.
    pwid = Picture1.ScaleX(pwid, vbTwips, Picture1.ScaleMode)
    phgt = Picture1.ScaleY(phgt, vbTwips, Picture1.ScaleMode)
    
    ' Compute the center of the object.
    xmid = Picture1.ScaleLeft + Picture1.ScaleWidth / 2
    ymid = Picture1.ScaleTop + Picture1.ScaleHeight / 2
    
    ' Pass the coordinates of the upper left and
    ' lower right corners into the Scale method.
    Printer.Scale _
        (xmid - pwid / 2, ymid - phgt / 2)- _
        (xmid + pwid / 2, ymid + phgt / 2)



  'print scopedata





    For n = 1 To Int(4095 / scalar)
    
   Y1 = (Int(1500 - (lvlscalar * (scopedata2(n) / 3))))
   
   If Y1 < 0 Then
   Y1 = 0
   End If
   
   If Y1 > 3000 Then
   Y1 = 3000
   End If
   
   Y2 = (Int(1500 - (lvlscalar * (scopedata2(n + 1) / 3))))
  
   If Y2 < 0 Then
   Y2 = 0
   End If
   
   If Y2 > 3000 Then
   Y2 = 3000
   End If
       
    Printer.Line ((Int((n * scalar))), Y1)-((Int(((n + 1) * scalar))), Y2), vbBlack
      
    Next n
    




    
Printer.DrawWidth = 2
Printer.Line (0, 0)-(0, 3000)
Printer.Line (4155, 0)-(4155, 3000)
Printer.Line (0, 0)-(4155, 0)
Printer.Line (0, 3000)-(4155, 3000)
Printer.DrawWidth = 1



For y = 0 To 10   'make gain lines
If y / 5 = Int(y / 5) Then
Printer.DrawWidth = 4
Printer.Line (4155, (y * 300))-(4305, (y * 300))
Printer.Line (-150, (y * 300))-(0, (y * 300))
Else
Printer.DrawWidth = 2
Printer.Line (4155, (y * 300))-(4255, (y * 300))
Printer.Line (-100, (y * 300))-(0, (y * 300))
End If
Next y
Printer.DrawWidth = 1


For y = 0 To 10   'make timebase lines
If y / 5 = Int(y / 5) Then
Printer.DrawWidth = 4
Printer.Line ((y * 415.5), 3000)-((y * 415.5), 3150)
Else
Printer.DrawWidth = 2
Printer.Line ((y * 415.5), 3000)-((y * 415.5), 3100)
End If
Next y

Printer.DrawWidth = 1



For X = 0 To 10   'print gain numbers

Printer.CurrentY = (X * 300) - 100
Printer.CurrentX = -400
texttoprint = CStr(Abs((X * 2) - 10))
Printer.Print texttoprint
Printer.CurrentX = 4455
Printer.CurrentY = (X * 300) - 100
Printer.Print texttoprint
Next X


For X = 0 To 10   'print timebase numbers

Printer.CurrentY = (3200)
Printer.CurrentX = (X * 415.5) - 75
texttoprint = CStr(X)
Printer.Print texttoprint
Next X



Printer.CurrentY = -500
Printer.CurrentX = 700
texttoprint = Text1.Text  'print legend
Printer.Print texttoprint

Printer.CurrentY = -700
Printer.CurrentX = 1525
texttoprint = Now     'print time and date
Printer.Print texttoprint


Printer.EndDoc


End Function



