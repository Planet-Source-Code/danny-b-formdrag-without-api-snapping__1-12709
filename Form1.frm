VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   1320
      Index           =   1
      Left            =   2160
      ScaleHeight     =   1290
      ScaleWidth      =   1890
      TabIndex        =   2
      Top             =   1905
      Width           =   1920
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   4695
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Quit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3720
         TabIndex        =   6
         Top             =   450
         Width           =   450
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3720
         TabIndex        =   5
         Top             =   255
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "This is a simple demo on how forms and objects can be moved without api."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   285
         Width           =   2580
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Move Mania"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   180
         TabIndex        =   3
         Top             =   60
         Width           =   1155
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   4680
         Y1              =   930
         Y2              =   930
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   1080
      Index           =   0
      Left            =   375
      ScaleHeight     =   1050
      ScaleWidth      =   1245
      TabIndex        =   0
      Top             =   1140
      Width           =   1275
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   4680
      Y1              =   945
      Y2              =   945
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FirstX, FirstY As Integer
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    FirstX = X
    FirstY = Y
End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Me.Left = Me.Left + (X - FirstX)
    Me.Top = Me.Top + (Y - FirstY)
End If

End Sub

Private Sub Label3_Click()
MsgBox "Created by Danny B" & Chr(13) & _
"danny@lambdastudios.com" & Chr(13) & Chr(13) & _
"Try moving the form by clicking and draging on the " & _
"orange background. And move the two pictureboxes, " & _
"note the snapping if you release thoose near an edge.", vbInformation, "Move Mania"
End Sub

Private Sub Label4_Click()
Unload Me
End
End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    FirstX = X
    FirstY = Y
End If

End Sub
Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbLeftButton Then
    Picture1(Index).Left = Picture1(Index).Left + (X - FirstX)
    Picture1(Index).Top = Picture1(Index).Top + (Y - FirstY)
End If

End Sub
Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Picture1(Index).Top < Picture2.Height + 315 Then Picture1(Index).Top = Picture2.Height
If Picture1(Index).Left < 300 Then Picture1(Index).Left = 0
If Picture1(Index).Top > Me.Height - (300 + Picture1(Index).Height) Then Picture1(Index).Top = Me.Height - Picture1(Index).Height
If Picture1(Index).Left > Me.Width - (300 + Picture1(Index).Width) Then Picture1(Index).Left = Me.Width - Picture1(Index).Width

End Sub
