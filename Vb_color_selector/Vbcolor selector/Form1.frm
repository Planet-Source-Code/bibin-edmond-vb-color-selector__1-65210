VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vb colour Mixer"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5835
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "About Me"
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2925
      Width           =   2655
   End
   Begin VB.TextBox txtblue 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtgreen 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtred 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.HScrollBar hsbblue 
      Height          =   375
      Left            =   1320
      Max             =   255
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.HScrollBar hsbgreen 
      Height          =   375
      Left            =   1320
      Max             =   255
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.HScrollBar hsbred 
      Height          =   375
      Left            =   1320
      Max             =   255
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Make your favourite colour and just copy paste it on your code ! pls vote for me"
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   3960
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "BLUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "GREEN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "RED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected colour :"
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
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'                                                           *
'Name :vb colour Mixer                                      *
'                                                           *
'Author:  Bibin edmond                                      *
'                                                           *
'Description : This is a small utility which can be used to *
'generate codes for colours in vb !                         *
'                                                           *
'************************************************************
'

Private Sub Command1_Click()
Form2.Show

End Sub

Private Sub Form_Load()
' Resetting the values

    hsbred.Value = 255
    hsbgreen.Value = 255
    hsbblue.Value = 255
End Sub


Private Sub hsbBlue_Change()


    Form1.BackColor = RGB(hsbred.Value, hsbgreen.Value, hsbblue.Value)
        txtblue.Text = hsbblue.Value
   Text1.Text = "RGB(" + txtblue.Text + "," + txtgreen.Text + "," + txtred.Text + ")"
    End Sub


Private Sub hsbGreen_Change()


    Form1.BackColor = RGB(hsbred.Value, hsbgreen.Value, hsbblue.Value)
        txtgreen.Text = hsbgreen.Value
   Text1.Text = "RGB(" + txtblue.Text + "," + txtgreen.Text + "," + txtred.Text + ")"
  End Sub


Private Sub hsbRed_Change()


    Form1.BackColor = RGB(hsbred.Value, hsbgreen.Value, hsbblue.Value)
        txtred.Text = hsbred.Value
   Text1.Text = "RGB(" + txtblue.Text + "," + txtgreen.Text + "," + txtred.Text + ")"

       
        
        
    End Sub


Private Sub txtBlue_Change()
    hsbblue.Value = txtblue.Text
End Sub


Private Sub txtGreen_Change()
    hsbgreen.Value = txtgreen.Text
End Sub


Private Sub txtRed_Change()
    hsbred.Value = txtred.Text
End Sub
