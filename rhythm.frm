VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rhythm Generator by Lemniscatus"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7440
   Icon            =   "rhythm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   7440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "?"
      Height          =   285
      Left            =   2850
      TabIndex        =   23
      Top             =   120
      Width           =   285
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "&Data Files"
      Height          =   375
      Left            =   6120
      TabIndex        =   16
      Top             =   6120
      Width           =   975
   End
   Begin VB.TextBox txtOffset 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   7
      ToolTipText     =   "Number of whole notes to generate...  "
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate Rhythm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   8
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtCount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   5
      ToolTipText     =   "Number of whole notes to generate...  "
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtResolution 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   960
      TabIndex        =   3
      ToolTipText     =   "Resolution where 1 is whole note, 2 is half note, 4 is quarter note, etc.  "
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtSeeds 
      BackColor       =   &H00C0E0FF&
      Height          =   285
      Left            =   960
      TabIndex        =   1
      ToolTipText     =   "Rhythmic seeds of the form s1.s2.s3... where s is a positive integer  "
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame frameFile 
      Caption         =   "File"
      Enabled         =   0   'False
      Height          =   4695
      Left            =   480
      TabIndex        =   20
      Top             =   1800
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdSave 
         Caption         =   "S&ave"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5640
         TabIndex        =   15
         Top             =   3840
         Width           =   615
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         TabIndex        =   14
         Top             =   3840
         Width           =   615
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   3015
      End
      Begin VB.DirListBox Dir1 
         Height          =   3015
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   3015
      End
      Begin VB.FileListBox File1 
         Height          =   3405
         Left            =   3480
         TabIndex        =   11
         Top             =   240
         Width           =   2775
      End
      Begin VB.TextBox txtFile 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   3885
         Width           =   3735
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "Error: Invalid Parameters!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   4320
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "File Name"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   3930
         Width           =   705
      End
   End
   Begin VB.TextBox txtRhythm 
      Enabled         =   0   'False
      Height          =   5175
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "rhythm.frx":0ECA
      Top             =   1560
      Width           =   7215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Offset"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1245
      Width           =   420
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.01 [One - Alpha] by Lemniscatus"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   3787
      TabIndex        =   19
      Top             =   840
      Width           =   3000
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rhythm Generator"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   3532
      TabIndex        =   18
      Top             =   360
      Width           =   3510
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1455
      Left            =   3240
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Philippine Copyright ©2017 by Noel C. Posicion.  All Rights Reserved"
      Enabled         =   0   'False
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   6840
      Width           =   4875
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Measure"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   885
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Resolution"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   525
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Seeds"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Rhythm generator using interference
Rem Philippine Copyright (c)2017 by Noel C. Posicion
Rem All Rights Reserved

Option Explicit

Const whole = 3840
Dim ErrorCode As Integer

Sub ShowFile(ByVal Y As Boolean)
    frameFile.Enabled = Y
    frameFile.Visible = Y
    cmdFile.Enabled = Not Y
    cmdFile.Visible = Not Y
End Sub

Function MyMax(ByVal a As Integer, ByVal b As Integer)
    MyMax = IIf(a > b, a, b)
End Function

Sub ErrHandler()
    Dim msg As String
    Dim s As String
    Dim t As String
    Select Case ErrorCode
        Case 1
            s = txtResolution.Text
            If Trim(s) = "" Then
                t = "empty string!"
            Else
                t = Trim(s) + " is an invalid integer!"
            End If
            msg = "Check resolution value -- " + t
        Case 2
            s = txtCount.Text
            If Trim(s) = "" Then
                t = "empty string!"
            Else
                t = Trim(s) + " is an invalid integer!"
            End If
            msg = "Check measure value -- " + t
        Case 3
            s = txtOffset.Text
            If Trim(s) = "" Then
                t = "empty string!"
            Else
                t = Trim(s) + " is an invalid integer!"
            End If
            msg = "Check offset value -- " + t
        Case Else
            s = txtSeeds.Text
            If Trim(s) = "" Then
                t = "empty string!"
            Else
                t = Trim(s) + " contains an invalid integer!"
            End If
            msg = "Check seeds value -- " + t
    End Select
    txtRhythm.Text = msg
    
End Sub

Private Sub cmdFile_Click()
    ShowFile True
End Sub

Private Sub cmdGenerate_Click()
    Dim s
    Dim seeds() As String
    Dim rhythm As String
    Dim values As String
    Dim bar As Long
    Dim i As Long
    Dim w As Long
    Dim last As Long
    Dim start As Long
    Dim count As Long
    Dim resolution As Long
    On Error GoTo ErrorLabel
    ErrorCode = 1
    resolution = txtResolution.Text
    ErrorCode = 2
    count = txtCount.Text
    ErrorCode = 3
    start = txtOffset.Text
    ErrorCode = 4
    seeds = Split(txtSeeds.Text, ".")
    If resolution = 0 Then
        resolution = 8
    End If
    rhythm = ""
    values = ""
    last = 0
    For i = resolution * start + 1 To resolution * count
        Dim ofs As Integer
        ofs = i Mod resolution
        bar = i - ofs
        For Each s In seeds
            Dim n As Long
            n = Int(s)
            If i Mod n = 0 Then
                If ofs = 0 Then
                    w = i - MyMax(bar - resolution, last)
                Else
                    w = i - MyMax(bar, last)
                End If
                If w > 0 Then
                    rhythm = rhythm + Format(w, "#")
                    If i > bar Then
                        rhythm = rhythm + "."
                    End If
                End If
                values = values + Format(w * whole / resolution, "#") + " "
                last = i
            End If
        Next
        If ofs = 0 Then
            w = i - MyMax(bar - resolution, last)
            If w > 0 Then
                rhythm = rhythm + Format(w, "#") + "-"
            End If
            rhythm = rhythm + " |"
            If i Mod (4 * resolution) = 0 Then
                rhythm = rhythm + "|"
            End If
            rhythm = rhythm + " "
            If w > 0 Then
                rhythm = rhythm + "-"
            Else
                values = values + " |"
                If i Mod (4 * resolution) = 0 Then
                    values = values + "|"
                End If
                values = values + " "
            End If
        End If
    Next i
    Dim info As String
    info = "N.B.: 1 is equivalent to 1/" + Format(resolution, "#") + " note. i.e. "
    info = info + Format(resolution, "#") + " is a whole note."
    info = info + vbCrLf + vbCrLf
    txtRhythm.Text = info + rhythm + vbCrLf + vbCrLf + values
    Exit Sub
ErrorLabel:
    ErrHandler
End Sub

Private Sub cmdLoad_Click()

    Dim fname As String
    Dim f1 As Integer
    
    AddExtension
    
    fname = Dir1.Path + "\" + txtFile.Text
    
    If Dir(fname) = "" Then
    
        lblStatus.Caption = Trim(txtFile.Text) + " does not exist..."
        lblStatus.Visible = True
        
        Exit Sub
        
    Else
    
        Dim dta As String
    
        On Error GoTo LoadError
        
        f1 = FreeFile
        Open fname For Input As #f1
        Input #f1, dta
        Close #f1
        
        Dim pars() As String
        Dim chk As String
        
        pars = Split(dta, ",")
        
        chk = pars(0)
        
        If Trim(chk) <> "RDF v.:1.0" Then
    
            lblStatus.Caption = Trim(txtFile.Text) + " is not a rhythm data file..."
            lblStatus.Visible = True
        
            Exit Sub
            
        Else
        
            txtSeeds.Text = pars(1)
            txtResolution.Text = pars(2)
            txtCount.Text = pars(3)
            txtOffset.Text = pars(4)
            
            cmdGenerate_Click
            
            ShowFile False
    
        End If
        
    End If
    
    Exit Sub
        
LoadError:

    lblStatus.Caption = "Error loading data..."
    lblStatus.Visible = True
    
End Sub

Sub AddExtension()
    If InStr(txtFile.Text, ".") = 0 And Right(UCase(txtFile.Text), 4) <> ".RDF" Then
        txtFile.Text = txtFile.Text + ".rdf"
    End If
End Sub

Private Sub cmdSave_Click()

    Dim rhythm As String
    Dim fname As String
    Dim dta As String
    Dim seeds As String
    Dim resolution As String
    Dim measure As String
    Dim offset As String
    Dim curr As String
    Dim result As String
    Dim f1 As Integer
    
    On Error GoTo saveError
    
    rhythm = txtRhythm.Text
    
    If Mid(rhythm, 1, 27) <> "N.B.: 1 is equivalent to 1/" Then
        lblStatus.Caption = "No rhythm data to save..."
        lblStatus.Visible = True
        Exit Sub
    End If
    
    AddExtension
    
    fname = Dir1.Path + "\" + txtFile.Text
    
    If Dir(fname) <> "" Then
    
        f1 = FreeFile
        Open fname For Input As #f1
        Input #f1, dta
        Close #f1
        
        If Mid(dta, 1, 10) <> "RDF v.:1.0" Then
            lblStatus.Caption = Trim(txtFile.Text) + " is not a rhythm file!"
            lblStatus.Visible = True
            Exit Sub
        Else
            Dim msg As String
            msg = txtFile.Text + " exists!" + vbCrLf + "Do you want to overwrite it?"
            If MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2 + vbMsgBoxSetForeground, "Rhythm Generator") <> vbYes Then
                Exit Sub
            End If
        End If
        
    End If
 
    seeds = txtSeeds.Text
    resolution = txtResolution.Text
    measure = txtCount.Text
    offset = txtOffset.Text
    
    dta = "RDF v.:1.0" + "," + seeds + "," + resolution + "," + measure + "," + offset
    
    f1 = FreeFile
    Open fname For Output As #f1
    Write #f1, dta
    Close #f1
    
    File1.Refresh
    
    ShowFile False
    
    Exit Sub
    
saveError:
    
    lblStatus.Caption = "Error saving data to file!"
    lblStatus.Visible = True
    
End Sub

Sub Inform()
    Dim info As String
    info = info + "Rhythm is generated using naive interference algorithm as demonstrated by "
    info = info + "Joseph Schillinger et al. It makes use of a simple mathematical algorithm "
    info = info + "to generate 'slices' of values which could be treated as musical "
    info = info + "rhythms." + vbCrLf + vbCrLf
    info = info + "Seeds are divisors that slice time to create a sequence of discreet values that we treat as rhythms" + vbCrLf
    info = info + "Resolution indicates 1 measure; the reciprocal of which is the lowest 'resolution' in the sequence" + vbCrLf
    info = info + "Measure is the total number of cummulative whole notes to generate" + vbCrLf
    info = info + "Offset is the beginning of our sequence of rhythmic values" + vbCrLf + vbCrLf
    info = info + "Example: A resolution of 8 means the value of 1 is 1/8 note, the value of 2 is 1/4 note "
    info = info + "and the value of 4 is 1/2 note. etc. Thus, the value of 8 is a whole note." + vbCrLf + vbCrLf
    info = info + "_____" + vbCrLf + vbCrLf
    info = info + "Ref.: A Guide to Schillinger’s Theory of Rhythm, Second Edition by Frans Absil, 2015"
    ShowFile False
    txtRhythm.Text = info

End Sub

Private Sub Command1_Click()
    Inform
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    lblStatus.Visible = False
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
    lblStatus.Visible = False
End Sub

Private Sub File1_Click()
    txtFile.Text = File1.FileName
    lblStatus.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 0 And KeyCode = 27 Then
        If frameFile.Visible Then
            ShowFile False
        End If
    End If
End Sub

Private Sub Form_Load()
    Inform
    Rem SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
End Sub

Sub Clear()
    ShowFile False
    txtRhythm.Text = "Editing . . ."
End Sub

Private Sub txtCount_Change()
    Clear
End Sub

Private Sub txtFile_Change()
    cmdLoad.Enabled = True
    cmdSave.Enabled = True
    lblStatus.Visible = False
End Sub

Private Sub txtOffset_Change()
    Clear
End Sub

Private Sub txtResolution_Change()
    Clear
End Sub

Private Sub txtSeeds_Change()
    Clear
End Sub
