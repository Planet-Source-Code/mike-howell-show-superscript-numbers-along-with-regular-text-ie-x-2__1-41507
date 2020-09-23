VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Defferentiate Expression"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ForeColor       =   &H00000000&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   3255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Mathematical Notation"
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3015
      Begin VB.CommandButton Sin 
         Caption         =   "Sin"
         BeginProperty Font 
            Name            =   "Bede"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton cos 
         Caption         =   "Cos"
         BeginProperty Font 
            Name            =   "Bede"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton tan 
         Caption         =   "Tan"
         BeginProperty Font 
            Name            =   "Bede"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   6
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton PowerOfN 
         Caption         =   "x^n"
         BeginProperty Font 
            Name            =   "Bede"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Cubed 
         Caption         =   "x^3"
         BeginProperty Font 
            Name            =   "Bede"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton Squared 
         Caption         =   "x^2"
         BeginProperty Font 
            Name            =   "Bede"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
   End
   Begin RichTextLib.RichTextBox expression 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bede"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Differentiate"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Expression:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Credit goes to Kevin, for Font file



Private Sub Cubed_Click()
expression.Text = expression.Text & Cubed.Caption
End Sub

Private Sub Form_Load()
MsgBox "If this does not work then copy the font file to your window file" _
, vbCritical, "..."    'If this program works then delete the Msgbox line

Squared.Caption = "xÂ"
Cubed.Caption = "xÃ"
PowerOfN.Caption = "x×"
Sin.Caption = "Sinù"
cos.Caption = "Cosù"
tan.Caption = "Tanù"
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &H8000000F
Label2.ForeColor = &HFF0000
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HFF0000
Label2.ForeColor = &HFFFFFF
End Sub

Private Sub PowerOfN_Click()
Dim NValue As String
Dim Power As String
NValue = InputBox("Enter the value of n", "Power of n")


For i = 0 To (Len(NValue) - 1)
    Select Case Mid(NValue, i + 1, 1)
    Case "0"
    Power = Power & PowerOfNone
    Case "1"
    Power = Power & PowerOfOne
    Case "2"
    Power = Power & PowerOfTwo
    Case "3"
    Power = Power & PowerOfThree
    Case "4"
    Power = Power & PowerOfFour
    Case "5"
    Power = Power & PowerOfFive
    Case "6"
    Power = Power & PowerOfSix
    Case "7"
    Power = Power & PowerOfSeven
    Case "8"
    Power = Power & PowerOfEight
    Case "9"
    Power = Power & PowerOfNine
    End Select
Next i

expression.Text = expression.Text & "x" & Power

End Sub

Private Sub Squared_Click()
expression.Text = expression.Text & Squared.Caption
End Sub
