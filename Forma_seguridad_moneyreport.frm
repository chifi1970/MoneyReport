VERSION 5.00
Begin VB.Form Forma_seguridad 
   BackColor       =   &H80000010&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Security"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MoneyReport.lvButtons_H btnborra 
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   873
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mode            =   0
      Value           =   0   'False
      Image           =   "Forma_seguridad_moneyreport.frx":0000
      cBack           =   14737632
   End
   Begin VB.TextBox txtpassword 
      BackColor       =   &H8000000A&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
   Begin VB.Image img_candado_up 
      Height          =   495
      Left            =   3120
      Picture         =   "Forma_seguridad_moneyreport.frx":0962
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image img_candado_down 
      Height          =   495
      Left            =   3600
      Picture         =   "Forma_seguridad_moneyreport.frx":40A0
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image btn_ok 
      Height          =   1695
      Left            =   4200
      Picture         =   "Forma_seguridad_moneyreport.frx":6ADA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type the password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Forma_seguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_ok_Click()
On Error Resume Next
transfiere$ = txtpassword.Text
Unload Me

End Sub

Private Sub btn_ok_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btn_ok.Picture = img_candado_down.Picture
End Sub


Private Sub btn_ok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btn_ok.Picture = img_candado_up.Picture

End Sub


Private Sub btnborra_Click()
On Error Resume Next
txtpassword.Text = ""
txtpassword.SetFocus
End Sub

Private Sub btnok_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
Top = 4700  '(Screen.Height - Height) / 2
Left = ((Screen.Width - Width) / 2) + 800


End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then
  btn_ok_Click
  Exit Sub
End If
End Sub


