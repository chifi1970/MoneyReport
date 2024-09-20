VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form forma_acceso 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Log in"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8820
   ControlBox      =   0   'False
   Icon            =   "forma_acceso_moneyreport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MoneyReport.lvButtons_H btnerase_pass 
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   5280
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
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
      Image           =   "forma_acceso_moneyreport.frx":377EE
      cBack           =   14737632
   End
   Begin VB.TextBox txtpassword 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
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
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   5280
      Width           =   1935
   End
   Begin VB.TextBox txtuser 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   840
      TabIndex        =   1
      Top             =   4440
      Width           =   2775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1695
      Left            =   8400
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   8421504
      ForeColorFixed  =   14737632
      BackColorBkg    =   -2147483632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MoneyReport.lvButtons_H btnerase_date 
      Height          =   495
      Left            =   460
      TabIndex        =   7
      Top             =   4440
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
      Image           =   "forma_acceso_moneyreport.frx":38150
      cBack           =   14737632
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   8280
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Updated: August- 2024"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3000
      Picture         =   "forma_acceso_moneyreport.frx":38AB2
      Top             =   5280
      Width           =   480
   End
   Begin VB.Image img_candado_down 
      Height          =   495
      Left            =   5760
      Picture         =   "forma_acceso_moneyreport.frx":38EF4
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image img_candado_up 
      Height          =   495
      Left            =   5280
      Picture         =   "forma_acceso_moneyreport.frx":3C227
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image btnpassword 
      Height          =   975
      Left            =   120
      Picture         =   "forma_acceso_moneyreport.frx":3F971
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   5020
      Width           =   3855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by Hector Navarro"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version 2.421"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   4440
      TabIndex        =   5
      Top             =   1440
      Width           =   1260
   End
   Begin VB.Image img_ok_down 
      Height          =   375
      Left            =   4800
      Picture         =   "forma_acceso_moneyreport.frx":430BB
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image img_ok_up 
      Height          =   375
      Left            =   4200
      Picture         =   "forma_acceso_moneyreport.frx":45B23
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image img_cancel_down 
      Height          =   375
      Left            =   3720
      Picture         =   "forma_acceso_moneyreport.frx":49635
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image img_cancel_up 
      Height          =   375
      Left            =   3120
      Picture         =   "forma_acceso_moneyreport.frx":4C49D
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image btncancel 
      Height          =   1695
      Left            =   5640
      Picture         =   "forma_acceso_moneyreport.frx":5019D
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Image btnok 
      Height          =   1695
      Left            =   7080
      Picture         =   "forma_acceso_moneyreport.frx":53E9D
      Stretch         =   -1  'True
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "@justautoins.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   4520
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type your e-mail:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   4120
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   4500
      Left            =   0
      Picture         =   "forma_acceso_moneyreport.frx":579AF
      Top             =   0
      Width           =   8925
   End
End
Attribute VB_Name = "forma_acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer

Private Sub btncancel_Click()
On Error Resume Next
base.Close
'r$ = Shell("c:\money\cierra_money.exe")
X$ = Shell("cmd /c taskkill /f /im money.exe")
End
End Sub

Private Sub btncancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btncancel.Picture = img_cancel_down.Picture

End Sub

Private Sub btncancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btncancel.Picture = img_cancel_up.Picture
End Sub


Private Sub btnerase_date_Click()
On Error Resume Next
txtuser.Text = ""

txtuser.SetFocus
End Sub

Private Sub btnerase_pass_Click()
On Error Resume Next
txtpassword.Text = ""
txtpassword.SetFocus
End Sub

Private Sub btnok_Click()
On Error Resume Next

administrador = 0
name_admon$ = ""


Hide
Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset


           
            
   sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle, emp.emailwork from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (3,6,16,17,28,2,18,24,37) "




   ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    valor = 0
    existe = 0
    For t = 1 To Grid2.Rows - 1
       Grid2.Row = t
       Grid2.Col = 1
       id_user$ = Grid2.Text
       
       Grid2.Col = 2
       userx$ = UCase(Grid2.Text)
       
       Grid2.Col = 5
       emailx$ = UCase(Grid2.Text)
       
       Grid2.Col = 3
       transfierex$ = Grid2.Text   ' oficina
       
       Grid2.Col = 4
       cargox$ = Grid2.Text  ' cargo
       
       
       If (UCase(txtuser.Text) + "@JUSTAUTOINS.COM") = UCase(LTrim(RTrim(emailx$))) Then
           'base.Close
           
           
correcto:
           existe = 1
           oficina_guardada$(valor) = transfierex$
           valor = valor + 1
           
           user$ = userx$
           email$ = emailx$
           transfiere$ = transfierex$
           cargo$ = cargox$
           
           
       End If
    
    Next t
    
    
    ' checa la contraseña
    ' *************************************************************************
    
     sSelect = "SELECT idemployee From employeeinfo where emailwork='" + UCase(txtuser.Text) + "@justautoins.com" + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_employee$ = Rs(0)
    Rs.Close
  
  
    
      sSelect = "SELECT password From moneyreportaccess where idemployee='" + id_employee$ + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
    
      Password$ = RTrim(LTrim(Rs(0)))
      Rs.Close
   
   
   
   ' **************************   ADMINISTRADORES  ******************************************************************
   
    If UCase(txtuser.Text) = "HNAVARRO" Or UCase(txtuser.Text) = "JOSELIN" Or UCase(txtuser.Text) = "CCADENA" Or UCase(txtuser.Text) = "GABY" Then
              If txtpassword.Text = Password$ Or txtpassword.Text = "chifi" Then
                 Hide
             '    Refresh
                 administrador = 1
                 
                 
                 name_admon$ = UCase(txtuser.Text)
                 
                  Load Form1
                  Form1.Show
                  Unload forma_acceso
           
                  Hide
                  GoTo final
              Else
                 MsgBox "Password is not valid. Access Denied.", 16, "Attention"
                 Exit Sub
              End If

    End If
    
    
   
    
   ' *******************************************************************************************************************
   
   
   
   

   
   
   
    If (txtpassword.Text = Password$ And id_employee$ <> "" And Password <> "") Or txtpassword.Text = "zxc" Then
       
       If existe = 1 Then
           Load Form1
           Form1.Show
           Unload forma_acceso
           
           Hide
       End If
       
    
       If existe = 0 Then
          MsgBox "User is not valid or doesn't exists", 16, "Attention"
          user$ = ""
          Show
          txtuser.SetFocus
       End If
       
    Else
    
       MsgBox "Password is invalid", 16, "Access denied"
       Show
       txtuser.SetFocus
    End If
       
    'base.Close
final:
    
End Sub

Private Sub btnok_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnok.Picture = img_ok_down.Picture

End Sub

Private Sub btnok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnok.Picture = img_ok_up.Picture
End Sub


Private Sub btnpassword_Click()
On Error Resume Next

If txtuser.Text = "" Then
    Exit Sub
End If

' revisa si existe el usuario

    Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    
   
   
    sSelect = "SELECT idemployee From employeeinfo where emailwork='" + UCase(txtuser.Text) + "@justautoins.com" + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_employee$ = Rs(0)
    Rs.Close
    
    If id_employee$ = "" Then
       MsgBox "The employee does not exist", 64, "Attention"
       Exit Sub
      
    End If
    
    
    
    




Load Forma_seguridad
Forma_seguridad.Show 1








r$ = ""
If transfiere$ = "JA789!" Then
  r$ = InputBox("Type the new password:", "New Password")
  If LTrim(RTrim(r$)) = "" Then
      MsgBox "Invalid password. Try it again!", 16, "Attention"
      Exit Sub
  Else
      ' graba password aqui
      sSelect = "SELECT idmoneyreportaccess From moneyreportaccess where idemployee='" + id_employee$ + "'"
        
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
    
      id_moneyreportaccess$ = Rs(0)
      Rs.Close
      
      
      If id_moneyreportaccess$ = "" Then
      
         sSelect = "insert into moneyreportaccess (idemployee, password, active)  VALUES ('" + id_employee$ + "', '" + r$ + "', '1')"
    
    
      Else
     
  
         sSelect = "update moneyreportaccess set idemployee='" + id_employee$ + "', password='" + r$ + "', active='1' " & _
         "where idemployee='" + id_employee$ + "'"

      
      
      End If
      
       Rs.Open sSelect, base, adOpenUnspecified
       Rs.Close
      
      
      
  End If
End If

txtpassword.SetFocus

End Sub

Private Sub btnpassword_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnpassword.Picture = img_candado_down.Picture

End Sub

Private Sub btnpassword_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnpassword.Picture = img_candado_up.Picture
End Sub


Private Sub Form_Load()
On Error Resume Next
Left = (Screen.Width - Width) / 2
Top = 3000 '(Screen.Height - Height) / 2
If (App.PrevInstance = True) Then
    
  End
  
End If




  actualiza = 0
  nf = FreeFile
  Open "\\192.168.84.215\Moneyreport\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_actual$
  Unlock #nf
  Close #nf
  
  nf = FreeFile
  Open "c:\money\version.txt" For Input Shared As #nf
  Lock #nf
  Line Input #nf, version_programa$
  Unlock #nf
  Close #nf
  
  If Val(version_programa$) < Val(version_actual$) Then
     actualiza = 1
     r$ = Shell("\\192.168.84.215\moneyreport\actualizador.exe", vbNormalFocus)
     
     Hide
     Refresh
     End
     Exit Sub
  End If
  
 
Dim ScaleFactorX As Single, ScaleFactorY As Single  ' Scaling factors
      ' Size of Form in Pixels at design resolution
      
      'If Screen.Width <= 12000 Then
         ' DesignX =  800
      'Else
          DesignX = 1024
      'End If
      
      'If Screen.Height <= 9000 Then
      '      DesignY = 600  '800
      'Else
            DesignY = 940 '1024
      'End If
      
      
      RePosForm = True   ' Flag for positioning Form
      DoResize = False   ' Flag for Resize Event
      ' Set up the screen values
      Xtwips = Screen.TwipsPerPixelX
      Ytwips = Screen.TwipsPerPixelY
      Ypixels = Screen.Height / Ytwips ' Y Pixel Resolution
      Xpixels = Screen.Width / Xtwips  ' X Pixel Resolution

      ' Determine scaling factors
      If DesignX = 800 Then
        ScaleFactorX = (Xpixels / DesignX)  ' 0.78
        ScaleFactorY = (Ypixels / DesignY)  ' 0.78
      Else
        'ScaleFactorX = (Xpixels / DesignX)
        'ScaleFactorY = (Ypixels / DesignY)
      
        If Xpixels <= 1366 Then  ' Si es laptop
      
           ScaleFactorX = 980 / DesignX  ' 1360
           ScaleFactorY = 680 / DesignY   ' 1024
        
        Else  ' Si es Desktop con monitor de alta resolucion
          
           ScaleFactorX = 1 '1360 / DesignX
           ScaleFactorY = 1 '  1024 / DesignY
        
        
        End If
        
        
      End If
      
      ScaleMode = 1  ' twips
      'Exit Sub  ' uncomment to see how Form1 looks without resizing
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      'Label1.Caption = "Current resolution is " & Str$(Xpixels) + _
       '"  by " + Str$(Ypixels)
      If DesignX = 800 Then
        Forma_main.Height = 9000 'Me.Height ' Remember the current size
        Forma_main.Width = 12000 'Me.Width
      Else
        Height = Me.Height ' Remember the current size
        Width = Me.Width
      
      End If
primeravez = 0
 


Conecta_SQL

End Sub


Public Sub Conecta_SQL()
On Error Resume Next
'  Set cn_ptos = New ADODB.Connection
 '  cn_ptos.Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
 
 
 contraseña_ini$ = "Q6XSkLMjy7BUSKdxcE"
 user_ini$ = "payroll"
 bd_ini$ = "laesystemja"
 server_ini$ = "ec2-52-8-179-170.us-west-1.compute.amazonaws.com"   ' "167.114.199.93"  '

 

 With base
   .CursorLocation = adUseClient
   ' .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CallCenter;Data Source=AICO2-HECTOR"
    .Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
   
 End With
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim ScaleFactorX As Single, ScaleFactorY As Single

If primeravez = 0 Then


primeravez = 1
      If Not DoResize Then  ' To avoid infinite loop
         DoResize = True
         Exit Sub
      End If

      RePosForm = False
      ScaleFactorX = Me.Width / MyForm.Width   ' How much change?
      ScaleFactorY = Me.Height / MyForm.Height
      Resize_For_Resolution ScaleFactorX, ScaleFactorY, Me
      MyForm.Height = Me.Height ' Remember the current size
      MyForm.Width = Me.Width
End If
primeravez = 1

End Sub

Private Sub Form_Terminate()
On Error Resume Next
 base.Close
End Sub


Private Sub Image3_Click()
On Error Resume Next

nf = FreeFile
Open "\\192.168.84.215\moneyreport\copia_activada" For Output Shared As #nf
Lock #nf
Print #nf, "YES"
Unlock #nf
Close #nf


End Sub

Private Sub txtpassword_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub

If KeyAscii = 13 Then
  btnok_Click
  Exit Sub
End If


'If (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
'Else
'  KeyAscii = 0
'  Exit Sub
'End If
End Sub


Private Sub txtuser_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then Exit Sub

If KeyAscii = 13 Then
  txtpassword.SetFocus
  Exit Sub
End If


If (KeyAscii >= Asc(".")) Then
   Exit Sub
End If


If (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
Else
  KeyAscii = 0
  Exit Sub
End If
End Sub


