VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000080&
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1620
      Width           =   3375
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   315
      Left            =   360
      TabIndex        =   4
      Top             =   1620
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   556
      _Version        =   393216
      Alignment       =   0
      BuddyControl    =   "Text1"
      BuddyDispid     =   196611
      OrigLeft        =   720
      OrigTop         =   960
      OrigRight       =   915
      OrigBottom      =   1755
      Max             =   4
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Switch Direction"
      Height          =   375
      Left            =   2940
      TabIndex        =   3
      Top             =   540
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Left            =   555
      MaxLength       =   1
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form1.frx":0000
      Top             =   1620
      Width           =   300
   End
   Begin Project1.jMarquee jMarquee1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   450
      ForeColor       =   65280
      Caption         =   "jMarquee v1.0 - <jay@alphamedia.net>"
      InsideTrackColor=   0
      OutsideTrackColor=   65535
      Speed           =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   7
      Top             =   1380
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Speed (0-4)"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   1380
      Width           =   825
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Command1.Caption = "Start" Then
        Command1.Caption = "Stop"
        jMarquee1.StartScroll
    Else
        Command1.Caption = "Start"
        jMarquee1.EndScroll
    End If
End Sub

Private Sub Command2_Click()
    If jMarquee1.Direction = dLeft Then
        jMarquee1.Direction = dRight
    Else
        jMarquee1.Direction = dLeft
    End If
End Sub

Private Sub Form_Load()
    Text2 = jMarquee1.Caption
    Text1 = "2"
End Sub

Private Sub Form_Unload(Cancel As Integer)
       End
End Sub

Private Sub Text1_Change()
    
    jMarquee1.Speed = Val(Text1)
End Sub

Private Sub Text2_Change()
    If jMarquee1.isScrolling Then Command1_Click
    jMarquee1.Caption = Text2
    
End Sub
