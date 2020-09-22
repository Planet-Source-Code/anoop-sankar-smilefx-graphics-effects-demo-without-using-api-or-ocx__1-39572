VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Smile FX"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   542
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Save Image"
      Height          =   375
      Left            =   5760
      TabIndex        =   19
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Effects"
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   5520
      Width           =   5415
      Begin VB.OptionButton Option1 
         Caption         =   "Tri-Color"
         Height          =   255
         Index           =   11
         Left            =   3840
         TabIndex        =   16
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Old Photo"
         Height          =   255
         Index           =   10
         Left            =   3840
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Gray Scale"
         Height          =   255
         Index           =   9
         Left            =   3840
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Negative"
         Height          =   255
         Index           =   8
         Left            =   3840
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sketch"
         Height          =   255
         Index           =   7
         Left            =   2160
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Talcom Powder"
         Height          =   255
         Index           =   6
         Left            =   2160
         TabIndex        =   11
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tarnish"
         Height          =   255
         Index           =   5
         Left            =   2160
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Mosaic"
         Height          =   255
         Index           =   4
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Color Blend"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Double Interleave"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Horizontal Interleave"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Vertical Interleave"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Choose an effect and click apply."
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2880
         TabIndex        =   9
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   2
      Top             =   5640
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H0099CC00&
      ForeColor       =   &H00FF8080&
      Height          =   5145
      Left            =   4080
      ScaleHeight     =   339
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   1
      Top             =   240
      Width           =   3915
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      ForeColor       =   &H00FF8080&
      Height          =   5145
      Left            =   120
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   339
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   0
      Top             =   240
      Width           =   3915
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "www.smilehouse.cjb.net"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   5880
      MouseIcon       =   "frmMain.frx":3DAD
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   6960
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Programmed by Anoop Sankar"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   6960
      Width           =   2175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------
'       Copyright 2002, Anoop Sankar
'You may freely use, modify and distribute this source
'code, provided that you do not remove this message.
'But, you are NOT allowed to distribute the compiled
'version (.EXE,.DLL,.OCX etc etc.) of this program
'or any program which uses the below code without my
'consent.
'
'If you modified something, put your name below..
'
'Orginal Code : Anoop Sankar (anoops@gmx.net)
'Modified by  : No one so far
'
'Last Update : Oct 3,2002
'Visit www.smilehouse.cjb.net for more source code
'-------------------------------------------------------
'
'The ShiftColor function was taken from gonchuki's
'Chameleon buttons source.
'The rest of the stuff was done by me.
'
'And oh.. the photograph was downloaded from the net.
'The beautiful face is that of Aishwarya Rai, a popular
'Indian actress, who was also Miss World in 1993.
'
'-------------------------------------------------------

Private Sub Form_Load()
    
    'For this demo Picture1 will be the source
    'and Picture2 will be the destination

    'set the unit to pixels
    Me.ScaleMode = vbPixels
    Picture1.ScaleMode = vbPixels
    Picture2.ScaleMode = vbPixels
    
    'we want the destination picture box to be stable
    Picture2.AutoRedraw = True

End Sub

Private Sub Command1_Click()
    
    'applying the effects
    
    Dim Effect As Integer

    'to find which effect user has selected
    For i = 0 To 11
        If Option1(i).Value Then Effect = i
    Next i
    
    Select Case Effect
        
        Case 0
            Interleave Picture1, Picture2, 2, Picture2.BackColor, 1
        Case 1
            Interleave Picture1, Picture2, 2, Picture2.BackColor, 2
        Case 2
            Interleave Picture1, Picture2, 4, Picture2.BackColor, 3
        Case 3
            Interleave Picture1, Picture2, 2, vbRed, 4
        Case 4
            Mosaic Picture1, Picture2, 10
        Case 5
            Tarnish Picture1, Picture2, 2
        Case 6
            Churn Picture1, Picture2, 250
        Case 7
            'wow!.. seems to be a multipurpose effect..
            'change the last argument to 50,100,500 and
            'you'll see what i mean!
            PencilDraw Picture1, Picture2, 8
        Case 8
            Negative Picture1, Picture2
        Case 9
            GrayScale Picture1, Picture2
        Case 10
            'old photo effect
            GrayScale Picture1, Picture2, , True
        Case 11
            'it is possible to have both destination
            'and source the same as shown below
            TriColor Picture1, Picture1
            TriColor Picture1, Picture2
    End Select
   
End Sub

Private Sub Command2_Click()
    
    'save pic routine
    Dim FileNaam As String
    On Local Error GoTo errTrap
    
    FileNaam = InputBox("Enter path and filename to save :", "Save Image", App.Path & "\ashmod.bmp")
    SavePicture Picture2.Image, FileNaam
    
    MsgBox "Done saving image as " & FileNaam, vbInformation, "Done it!"
    Exit Sub
    
errTrap:
    Dim msg As String
    
    msg = "Could not save file due to below error" & vbCrLf & vbCrLf
    msg = msg & "Error #" & Err.Number & ":"
    msg = msg & Err.Description

    MsgBox msg, vbCritical, "Error!!!"
End Sub

Private Sub Label3_Click(Index As Integer)
    
    'website link
    If Index = 1 Then
        'wowed not to use API, this is the only way out
        Shell "start.exe http://www.smilehouse.cjb.net/"
    End If

End Sub
