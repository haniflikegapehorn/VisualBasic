VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "ComboBox (Hanif Project).frx":0000
      Left            =   2520
      List            =   "ComboBox (Hanif Project).frx":0010
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Hasil 
      Caption         =   "Hasil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Nama3 
      Caption         =   "Luas ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Nama2 
      Caption         =   "Wide ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   2040
      Width           =   2535
   End
   Begin VB.Label Nama1 
      Caption         =   "Long ="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "PICK A NUMBER PLEASE!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
X = Combo1.Text
Select Case X
Case "Persegi Panjang":
Nama1.Caption = "Panjang ="
Nama2.Caption = "Lebar ="
Nama3.Caption = "Luas ="
Hasil.Caption = "..."
Case "Segitiga":
Name1.Caption = "Alas ="
Nama2.Caption = "Tinggi ="
Nama3.Caption = "Luas ="
Hasil.Caption = "..."
Case "Tabung":
Nama1.Caption = "Jari-Jari Alas ="
Nama2.Caption = "Tinggi ="
Nama3.Caption = "Volume ="
Hasil.Caption = "..."
End Select
End Sub

Private Sub Command1_Click()
X = Combo1.Text
Select Case X
Case "Persegi Panjang":
LV = Text1.Text * Text2.Text
Hasil.Caption = "..."
Case "Segitiga":
LV = Text1.Text * Text2.Text / 2
Case "Tabung":
LV = 3.14 * Text1.Text ^ 2 * Text2.Text
Case "Kerucut"
LV = 3.14 * Text1.Text ^ 2 * Text2.Text / 3
End Select
Hasil.Caption = LV
End Sub

