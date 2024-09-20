VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form forma_reportes_pendientes 
   BackColor       =   &H8000000A&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pending unclosed reports"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4920
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   360
      TabIndex        =   3
      Top             =   540
      Width           =   2535
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   5280
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   3255
      Left            =   7800
      TabIndex        =   1
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5741
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
   Begin VB.Image img_ok_down 
      Height          =   375
      Left            =   4560
      Picture         =   "forma_reportes_pendientes (moneyreport).frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image img_ok_up 
      Height          =   375
      Left            =   4080
      Picture         =   "forma_reportes_pendientes (moneyreport).frx":2A68
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image btnok 
      Height          =   1695
      Left            =   3120
      Picture         =   "forma_reportes_pendientes (moneyreport).frx":657A
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reports:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "forma_reportes_pendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btnok_Click()
On Error Resume Next
If List3.ListCount = 0 Then
  Form1.Timer1.Enabled = False
  Form1.Picture2.Visible = False
End If

If List3.ListCount > 2 Then
   bloqueado = 1
Else
   bloqueado = 0
End If


Unload Me
End Sub

Private Sub btnok_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnok.Picture = img_ok_down.Picture
End Sub

Private Sub btnok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnok.Picture = img_ok_up.Picture

End Sub



Public Sub carga_reportes()
On Error Resume Next

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
           
  
  id_employee = Val(Right(Form1.cbo_agentes.List(Form1.cbo_agentes.ListIndex), 20))
  ID_manager = Val(Right(Form1.cbo_managers.List(Form1.cbo_managers.ListIndex), 20))
  lae_office$ = RTrim(LTrim(Right(Form1.cbo_oficina.List(Form1.cbo_oficina.ListIndex), 25)))
 
  
  
' carga las fechas de trabajo sin domingos
' ----------------------------------------------------------------------------
  List2.Clear
  List3.Clear

mes = Val(Format(Now, "mm"))
ano = Val(Format(Now, "yyyy"))

For Y = 1 To mes

 Select Case Y
 Case 1, 3, 5, 7, 8, 10, 12
   dias = 31
 Case 2
   r = Val(Format(Now, "yyyy")) / 4
   r2 = Int(Val(Format(Now, "yyyy")))
   residuo = r - r2
   If residuo = 0 Then
     dias = 29
   Else
     dias = 28
   End If
 Case 4, 6, 9, 11
   dias = 30
 End Select
  
 If Y = mes Then
    dias = Val(Format(Now, "dd"))
 End If
 
 If ano <= 2022 And mes = 1 Then
   inicial = 10
 Else
   inicial = 1
 End If
 
 For t = inicial To dias
   
  If Y = 1 And t <= 10 Then
  
  Else
 
   f$ = Format(Y, "00") + "/" + Format(t, "00") + "/" + Format(ano, "0000")
   a$ = Format(f$, "ddd")
   
   If UCase(a$) <> "SUN" Then
     List2.AddItem f$
   End If
  End If
  
 Next t
   
Next Y


List2.RemoveItem List2.ListCount - 1


' ------------------------------------------------------------------------------


  
  
  
  
  sSelect = "select idoffice from officescatalog where office='" + lae_office$ + "'"
  Rs.Open sSelect, base, adOpenUnspecified
  id_oficina$ = Rs(0)
  Rs.Close
  
  
  sSelect = "select datecreated from employeeinfo where idemployee='" + Format(id_employee, "###0") + "'"
  Rs.Open sSelect, base, adOpenUnspecified
  fecha_contratacion$ = Format(Rs(0), "mm/dd/yyyy")
  Rs.Close
  
  
     
 ' Grid2.Visible = False
   
   
  Grid2.Clear
  List3.Clear
    
    
  ' verifica fecha de contratacion
  ano_contratacion$ = Right(fecha_contratacion$, 4)
  mes_contratacion$ = Left(fecha_contratacion$, 2)
  dia_contratacion$ = Mid$(fecha_contratacion$, 4, 2)
    
  fecha_actual$ = Format(Now, "mm/dd/yyyy")
  ano_actual$ = Right(fecha_actual$, 4)
  mes_actual$ = Left(fecha_actual$, 2)
  dia_actual$ = Mid$(fecha_actual$, 4, 2)
  
    
  id_moneyreport$ = ""
  If Val(ano_contratacion$) = Val(ano_actual$) Then
     
    sSelect = "select datereport from moneyreport where idemployee='" + Format(id_employee, "###0") + "'  and datereport>'" + fecha_contratacion$ + "'  and (dayoff='1' or submitted='1') order by datereport "
  Else
    sSelect = "select datereport from moneyreport where idemployee='" + Format(id_employee, "###0") + "'  and datereport>'01-09-2022' and (dayoff='1' or submitted='1') order by datereport "
  End If
    
   
    
    
    
    
    
    
           
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
       
       year_contrata$ = Format(fecha_contratacion$, "yyyy")
       year_actual$ = Format(Now, "yyyy")
       
       If Grid2.Rows <= 1 Then
          GoTo salida
       End If
       
       
       If year_contrata$ = year_actual$ Then
    
         x0 = Format(fecha_contratacion$, "y")
       
         For Y = 0 To List2.ListCount - 1
           f1$ = List2.List(Y)
           
           
           existe = 0
           For t = 1 To Grid2.Rows - 1
              Grid2.Row = t
              Grid2.Col = 1
              f2$ = Format(Grid2.Text, "mm/dd/yyyy")
              X2 = Format(f2$, "y")
              
              If Right(f1$, 4) = Right(f2$, 4) And X2 > x0 Then
                 existe = 1
                 Exit For
              End If
           Next t
       
           X1 = Format(f1$, "y")
           If existe = 0 And X1 > x0 Then
               List3.AddItem f1$
           End If
       
         Next Y
         
       Else
       
         For Y = 0 To List2.ListCount - 1
           f1$ = List2.List(Y)
           
           
           existe = 0
           For t = 1 To Grid2.Rows - 1
              Grid2.Row = t
              Grid2.Col = 1
              f2$ = Format(Grid2.Text, "mm/dd/yyyy")
              
              
              If f1$ = f2$ Then
                 existe = 1
                 Exit For
              End If
           Next t
       
           
           If existe = 0 Then
               List3.AddItem f1$
           End If
       
         Next Y
       
       
    
       End If
    
    
salida:
    
   If valido1 = 999 Then
    transfiere$ = List3.ListCount
    Unload Me
   End If
   
   
   
    
    
    
End Sub

Private Sub Form_Load()
On Error Resume Next
Top = 0
Left = (Screen.Width - Width) / 2
transfiere$ = ""

carga_reportes


End Sub


Private Sub List1_Click()
On Error Resume Next
If List1.ListIndex = -1 Then Exit Sub

transfiere$ = List1.List(List1.ListIndex)

End Sub


Private Sub List3_Click()
On Error Resume Next
transfiere$ = List3.List(List3.ListIndex)

End Sub


