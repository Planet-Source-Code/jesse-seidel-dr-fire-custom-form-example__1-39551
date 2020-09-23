VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   3555
   ClientTop       =   4410
   ClientWidth     =   8205
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   6375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "By: SpitFire"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7080
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "  Cancel"
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
      Left            =   4080
      TabIndex        =   4
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image Image7 
      Height          =   525
      Left            =   3960
      Picture         =   "Pass.frx":0000
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "      OK"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.Image Image6 
      Height          =   525
      Left            =   2520
      Picture         =   "Pass.frx":00E8
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the case sensitive password and click OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Crack The Password Punk"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   7680
      Picture         =   "Pass.frx":01D0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   435
   End
   Begin VB.Image Image4 
      Height          =   5715
      Left            =   8160
      Picture         =   "Pass.frx":06F7
      Stretch         =   -1  'True
      Top             =   360
      Width           =   45
   End
   Begin VB.Image Image3 
      Height          =   45
      Left            =   0
      Picture         =   "Pass.frx":0743
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   8220
   End
   Begin VB.Image Image2 
      Height          =   5715
      Left            =   0
      Picture         =   "Pass.frx":079B
      Stretch         =   -1  'True
      Top             =   360
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   540
      Left            =   0
      Picture         =   "Pass.frx":07E7
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8205
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is for the custom layout of the form, you dont need to change this
    If Button = vbLeftButton Then
        Dim lngRetVal As Long
        lngRetVal = ReleaseCapture()
        lngRetVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    Else
        Exit Sub
    End If
    
SetWindowPos Form1.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub

Private Sub Image5_Click()
End
End Sub

Private Sub Image6_Click()
If Text1.Text = "fr00tl00pz" Then MsgBox ("YOU GOT IT! NOW GO CRACK MICROSOFT!!!"), vbOKOnly
If Text1.Text = "fr00tl00pz" Then Form1.Hide
End Sub

Private Sub Image7_Click()
End
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is for the custom layout of the form, you dont need to change this
    If Button = vbLeftButton Then
        Dim lngRetVal As Long
        lngRetVal = ReleaseCapture()
        lngRetVal = SendMessage(Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    Else
        Exit Sub
    End If
    
SetWindowPos Form1.hwnd, -1, 0, 0, 0, 0, &H1 Or &H2
End Sub

Private Sub Label3_Click()
If Text1.Text = "fr00tl00pz" Then MsgBox ("YOU GOT IT! NOW GO CRACK MICROSOFT!!!"), vbOKOnly
If Text1.Text = "fr00tl00pz" Then Form1.Hide
End Sub

Private Sub Label4_Click()
End
End Sub
