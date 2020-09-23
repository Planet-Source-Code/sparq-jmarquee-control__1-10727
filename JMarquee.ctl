VERSION 5.00
Begin VB.UserControl jMarquee 
   BackColor       =   &H000000FF&
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4620
   ScaleHeight     =   300
   ScaleWidth      =   4620
   Begin VB.PictureBox Back 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      Begin VB.Label lblMarquee 
         Caption         =   "jMarquee"
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   660
      End
   End
End
Attribute VB_Name = "jMarquee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Dims
Dim mDirection As jDirection
Dim mSpeed As jSpeed
Dim MinHeight As Integer
Dim Scrolling As Boolean

'Enums
Public Enum jDirection
    dLeft        '0
    dRight       '1
End Enum

Public Enum jSpeed
    VerySlow    '0
    Slow        '1
    Medium      '2
    Fast        '3
    VaryFast    '4
End Enum

'Events

Event ScrollBegin()
Event ScrollEnd()
Event CaptionChange()

Private Sub lblMarquee_Click(Index As Integer)
    lblMarquee(Index).Width = Width
End Sub

Private Sub UserControl_Initialize()
    MinHeight = lblMarquee(0).Height * lblMarquee.Count
    Scrolling = False
End Sub

Private Sub UserControl_Resize()
    If Height < MinHeight Then Height = MinHeight
    With lblMarquee(0)
        .Width = Width
        .Left = 0
        .Top = (Height / 2) - (.Height / 2)
    End With

    With Back
        .Width = Width
        .Height = Height
        .Left = 0
        .Top = 0
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDirection = PropBag.ReadProperty("Direction", 1)
    lblMarquee(0).ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
    lblMarquee(0).Caption = PropBag.ReadProperty("Caption", "jMarquee")
    lblMarquee(0).BackColor = PropBag.ReadProperty("InsideTrackColor", vbWhite)
    Back.BackColor = PropBag.ReadProperty("OutsideTrackColor", vbRed)
    mSpeed = PropBag.ReadProperty("Speed", 2)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Direction", mDirection, 1)
    Call PropBag.WriteProperty("ForeColor", lblMarquee(0).ForeColor, vbBlack)
    Call PropBag.WriteProperty("Caption", lblMarquee(0).Caption, "jMarquee")
    Call PropBag.WriteProperty("InsideTrackColor", lblMarquee(0).BackColor, vbWhite)
    Call PropBag.WriteProperty("OutsideTrackColor", Back.BackColor, vbRed)
    Call PropBag.WriteProperty("Speed", mSpeed, 2)
End Sub

Public Property Get Direction() As jDirection
    Direction = mDirection
End Property

Public Property Let Direction(ByVal New_Direction As jDirection)
    mDirection = New_Direction
    PropertyChanged "Direction"
End Property

Public Property Get Caption() As String
    Caption = lblMarquee(0).Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblMarquee(0).Caption = New_Caption
    PropertyChanged "Caption"
    lblMarquee(0).Width = Width
    RaiseEvent CaptionChange
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = lblMarquee(0).ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblMarquee(0).ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get InsideTrackColor() As OLE_COLOR
    InsideTrackColor = lblMarquee(0).BackColor
End Property

Public Property Let InsideTrackColor(ByVal New_InsideTrackColor As OLE_COLOR)
    lblMarquee(0).BackColor = New_InsideTrackColor
    PropertyChanged "InsideTrackColor"
End Property

Public Property Get OutSideTrackColor() As OLE_COLOR
    OutSideTrackColor = Back.BackColor
End Property

Public Property Let OutSideTrackColor(ByVal New_OutSideTrackColor As OLE_COLOR)
    Back.BackColor = New_OutSideTrackColor
    PropertyChanged "OutSideTrackColor"
End Property


Public Property Get Speed() As jSpeed
    Speed = mSpeed
End Property

Public Property Let Speed(ByVal New_Speed As jSpeed)
    mSpeed = New_Speed
    PropertyChanged "Speed"
End Property

Public Property Get isScrolling() As Boolean
    isScrolling = Scrolling
End Property

Public Function StartScroll()
  Dim X As Integer
  Dim String1 As String
  Dim String2 As String
  Dim Pause As Integer
    
    RaiseEvent ScrollBegin
    
    Do Until TextWidth(lblMarquee(0)) >= lblMarquee(0).Width
        lblMarquee(0) = lblMarquee(0) & " "
    Loop
    Scrolling = True
    Do
       If mDirection = dLeft Then
            String1 = Right$(lblMarquee(0).Caption, Len(lblMarquee(0).Caption) - 1)
            String2 = Left$(lblMarquee(0), 1)
            lblMarquee(0).Caption = String1 & String2
       Else
            String2 = Right$(lblMarquee(0).Caption, 1)
            String1 = Left$(lblMarquee(0).Caption, Len(lblMarquee(0).Caption) - 1)
            lblMarquee(0).Caption = String2 & String1
       End If
              
       If mSpeed = VerySlow Then
            Pause = 10000
       ElseIf mSpeed = Slow Then
            Pause = 5000
       ElseIf mSpeed = Medium Then
            Pause = 2500
       ElseIf mSpeed = Fast Then
            Pause = 1000
       ElseIf mSpeed = VaryFast Then
            Pause = 500
       End If
       
       X = 0
       Do Until X = Pause
            X = X + 1
            DoEvents
            If Scrolling = False Then Exit Function
       Loop
    Loop
End Function

Public Function EndScroll()
    Scrolling = False
    RaiseEvent ScrollEnd
End Function
