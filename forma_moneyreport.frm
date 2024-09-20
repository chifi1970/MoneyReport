VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form Form1 
   BackColor       =   &H00006400&
   Caption         =   "Money Report"
   ClientHeight    =   15465
   ClientLeft      =   16320
   ClientTop       =   465
   ClientWidth     =   28560
   Icon            =   "forma_moneyreport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   15465
   ScaleWidth      =   28560
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid4 
      Height          =   1455
      Left            =   19680
      TabIndex        =   329
      Top             =   2040
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2566
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MoneyReport.lvButtons_H btnadd_agent 
      Height          =   375
      Left            =   7200
      TabIndex        =   328
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Caption         =   "Add"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.TextBox txtagente 
      BackColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   5280
      TabIndex        =   327
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      Height          =   15495
      Left            =   21600
      ScaleHeight     =   15435
      ScaleWidth      =   6915
      TabIndex        =   326
      Top             =   0
      Width           =   6975
   End
   Begin VB.PictureBox msg 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   9000
      ScaleHeight     =   1425
      ScaleWidth      =   5865
      TabIndex        =   34
      Top             =   4800
      Visible         =   0   'False
      Width           =   5895
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wait a moment please..."
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   36
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label lblmsg1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loading the information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   735
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Width           =   5295
      End
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3975
      Left            =   9240
      TabIndex        =   136
      Top             =   1440
      Visible         =   0   'False
      Width           =   5055
      _Version        =   524288
      _ExtentX        =   8916
      _ExtentY        =   7011
      _StockProps     =   1
      BackColor       =   3329337
      Year            =   2021
      Month           =   9
      Day             =   7
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   0   'False
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MoneyReport.lvButtons_H btnlock 
      Height          =   315
      Left            =   3120
      TabIndex        =   324
      Top             =   840
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      Caption         =   "Lock"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin VB.PictureBox msgdescanso 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   480
      ScaleHeight     =   1305
      ScaleWidth      =   5655
      TabIndex        =   256
      Top             =   1440
      Visible         =   0   'False
      Width           =   5685
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Day Off"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   1695
         Left            =   360
         TabIndex        =   257
         Top             =   45
         Width           =   4935
      End
      Begin VB.Shape Shape18 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         Height          =   1095
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5415
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   11880
      Top             =   0
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   14400
      ScaleHeight     =   465
      ScaleWidth      =   6465
      TabIndex        =   248
      Top             =   0
      Visible         =   0   'False
      Width           =   6495
      Begin MoneyReport.lvButtons_H btncargar_reportes 
         Height          =   375
         Left            =   5520
         TabIndex        =   250
         Top             =   45
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "See them"
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   32896
      End
      Begin VB.Label mensaje 
         BackStyle       =   0  'Transparent
         Caption         =   "A T T E N T I O N ... You have pending unclosed reports"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         TabIndex        =   249
         Top             =   100
         Width           =   5535
      End
   End
   Begin MoneyReport.lvButtons_H btnrevisar 
      Height          =   375
      Left            =   18000
      TabIndex        =   252
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      Caption         =   "See pending unclosed"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   12720
      Top             =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   3255
      Left            =   22320
      TabIndex        =   14
      Top             =   2280
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
   Begin VB.PictureBox marco_revisado 
      Appearance      =   0  'Flat
      BackColor       =   &H00006400&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3360
      ScaleHeight     =   1185
      ScaleWidth      =   3465
      TabIndex        =   110
      Top             =   12120
      Visible         =   0   'False
      Width           =   3495
      Begin MoneyReport.lvButtons_H btndesbloquear 
         Height          =   300
         Left            =   2760
         TabIndex        =   253
         Top             =   240
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         CapAlign        =   2
         BackStyle       =   2
         Shape           =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":16B92
         ImgSize         =   48
         cBack           =   -2147483633
      End
      Begin VB.CheckBox chk_revisado 
         BackColor       =   &H00006400&
         Caption         =   "Reviewed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   255
         Left            =   120
         TabIndex        =   111
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblnum 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   2760
         TabIndex        =   254
         Top             =   600
         Width           =   615
      End
      Begin VB.Image btnok 
         Height          =   855
         Left            =   1800
         Picture         =   "forma_moneyreport.frx":19E7A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   360
      ScaleHeight     =   60
      ScaleWidth      =   21135
      TabIndex        =   104
      Top             =   1420
      Width           =   21135
   End
   Begin VB.ComboBox cbo_year 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   43
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   21840
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkagentes 
      BackColor       =   &H00006400&
      Caption         =   "Inactive agents"
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
      Height          =   255
      Index           =   1
      Left            =   6840
      TabIndex        =   38
      Top             =   960
      Width           =   1455
   End
   Begin VB.CheckBox chkagentes 
      BackColor       =   &H00006400&
      Caption         =   "Active agents"
      Enabled         =   0   'False
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
      Height          =   255
      Index           =   0
      Left            =   5280
      TabIndex        =   37
      Top             =   960
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.ComboBox cboimpre 
      BackColor       =   &H80000010&
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
      Height          =   315
      Left            =   12240
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   12840
      Width           =   3615
   End
   Begin VB.ComboBox cbo_managers 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   10320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   360
      Width           =   2895
   End
   Begin VB.ComboBox cbo_agentes 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   5280
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   360
      Width           =   3615
   End
   Begin VB.ComboBox cbo_oficina 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   360
      Width           =   3135
   End
   Begin VB.PictureBox Picture3 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   45
      Top             =   0
      Width           =   0
   End
   Begin VB.PictureBox Barra_localizacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00006400&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   360
      ScaleHeight     =   1335
      ScaleWidth      =   21255
      TabIndex        =   127
      Top             =   1440
      Width           =   21255
      Begin VB.CheckBox chk_dayoff 
         BackColor       =   &H00000000&
         Caption         =   "Day Off"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   255
         Left            =   6120
         TabIndex        =   251
         Top             =   920
         Width           =   1215
      End
      Begin VB.Label lbloficina_agente 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "San Bernardino"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   16200
         TabIndex        =   135
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label lbl_iniciales_agente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "xxx"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   13200
         TabIndex        =   134
         Top             =   540
         Width           =   3780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Initials"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Index           =   2
         Left            =   13200
         TabIndex        =   133
         Top             =   200
         Width           =   1245
      End
      Begin VB.Label lblname_agent 
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "xxx"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   375
         Left            =   9360
         TabIndex        =   132
         Top             =   540
         Width           =   3495
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   405
         Left            =   9240
         Shape           =   4  'Rounded Rectangle
         Top             =   520
         Width           =   3735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Index           =   1
         Left            =   9240
         TabIndex        =   131
         Top             =   200
         Width           =   960
      End
      Begin VB.Image btncalendar 
         Height          =   1095
         Left            =   7920
         Picture         =   "forma_moneyreport.frx":1DB34
         Stretch         =   -1  'True
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lbldate_agente 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "dd/mm/yyyy"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   6120
         TabIndex        =   130
         Top             =   420
         Width           =   2100
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   270
         Index           =   0
         Left            =   6195
         TabIndex        =   129
         Top             =   200
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   1180
         Index           =   1
         Left            =   5880
         Shape           =   4  'Rounded Rectangle
         Top             =   40
         Width           =   15015
      End
      Begin VB.Label lblleyenda 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Report Form"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   128
         Top             =   240
         Width           =   5415
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   1180
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   40
         Width           =   5415
      End
   End
   Begin MoneyReport.lvButtons_H btnlock2 
      Height          =   315
      Left            =   12120
      TabIndex        =   325
      Top             =   750
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      Caption         =   "Lock"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   12632256
   End
   Begin VB.PictureBox hoja1 
      Appearance      =   0  'Flat
      BackColor       =   &H00006400&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10575
      Left            =   360
      ScaleHeight     =   10575
      ScaleWidth      =   21255
      TabIndex        =   46
      Top             =   1560
      Width           =   21255
      Begin MoneyReport.lvButtons_H btnlimpiacash 
         Height          =   495
         Index           =   0
         Left            =   3600
         TabIndex        =   214
         Top             =   3360
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":217C3
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiacash 
         Height          =   495
         Index           =   5
         Left            =   3600
         TabIndex        =   219
         Top             =   4200
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":22125
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiacash 
         Height          =   495
         Index           =   1
         Left            =   3600
         TabIndex        =   215
         Top             =   4920
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":22A87
         ImgSize         =   40
         cBack           =   32768
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00006400&
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   6920
         TabIndex        =   300
         Top             =   2460
         Width           =   1695
         Begin MoneyReport.lvButtons_H op_cards 
            Height          =   375
            Index           =   0
            Left            =   315
            TabIndex        =   301
            Top             =   0
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   661
            Caption         =   "1-10"
            CapAlign        =   2
            BackStyle       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   -1  'True
            cBack           =   12632256
         End
         Begin MoneyReport.lvButtons_H op_cards 
            Height          =   375
            Index           =   1
            Left            =   840
            TabIndex        =   302
            Top             =   0
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            Caption         =   "11-20"
            CapAlign        =   2
            BackStyle       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   2
            Value           =   0   'False
            cBack           =   12632256
         End
      End
      Begin MoneyReport.lvButtons_H btnfix 
         Height          =   375
         Left            =   19320
         TabIndex        =   323
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   6
         Shape           =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":233E9
         ImgSize         =   40
         cBack           =   32768
      End
      Begin VB.PictureBox tabcard 
         Appearance      =   0  'Flat
         BackColor       =   &H00006400&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Index           =   1
         Left            =   6720
         ScaleHeight     =   3735
         ScaleWidth      =   2190
         TabIndex        =   279
         Top             =   2880
         Visible         =   0   'False
         Width           =   2190
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC70&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   19
            Left            =   480
            TabIndex        =   299
            Top             =   3240
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC8F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   18
            Left            =   480
            TabIndex        =   298
            Top             =   2880
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC70&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   17
            Left            =   480
            TabIndex        =   297
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC8F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   16
            Left            =   480
            TabIndex        =   296
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC70&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   15
            Left            =   480
            TabIndex        =   295
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC8F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   14
            Left            =   480
            TabIndex        =   294
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC70&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   13
            Left            =   480
            TabIndex        =   293
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC8F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   12
            Left            =   480
            TabIndex        =   292
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC70&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   11
            Left            =   480
            TabIndex        =   291
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC8F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   10
            Left            =   480
            TabIndex        =   290
            Top             =   0
            Width           =   1215
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   10
            Left            =   240
            TabIndex        =   280
            Top             =   0
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":23833
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   11
            Left            =   240
            TabIndex        =   281
            Top             =   360
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":24195
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   12
            Left            =   240
            TabIndex        =   282
            Top             =   720
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":24AF7
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   13
            Left            =   240
            TabIndex        =   283
            Top             =   1080
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":25459
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   14
            Left            =   240
            TabIndex        =   284
            Top             =   1440
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":25DBB
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   15
            Left            =   240
            TabIndex        =   285
            Top             =   1800
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":2671D
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   16
            Left            =   240
            TabIndex        =   286
            Top             =   2160
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":2707F
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   17
            Left            =   240
            TabIndex        =   287
            Top             =   2520
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":279E1
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   18
            Left            =   240
            TabIndex        =   288
            Top             =   2880
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":28343
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   19
            Left            =   240
            TabIndex        =   289
            Top             =   3240
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":28CA5
            ImgSize         =   40
            cBack           =   32768
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "11"
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
            Height          =   255
            Index           =   19
            Left            =   40
            TabIndex        =   322
            Top             =   120
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "20"
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
            Height          =   255
            Index           =   18
            Left            =   40
            TabIndex        =   321
            Top             =   3360
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "19"
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
            Height          =   255
            Index           =   17
            Left            =   40
            TabIndex        =   320
            Top             =   3000
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "18"
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
            Height          =   255
            Index           =   16
            Left            =   40
            TabIndex        =   319
            Top             =   2640
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "17"
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
            Height          =   255
            Index           =   15
            Left            =   40
            TabIndex        =   318
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "16"
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
            Height          =   255
            Index           =   14
            Left            =   40
            TabIndex        =   317
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "15"
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
            Height          =   255
            Index           =   13
            Left            =   40
            TabIndex        =   316
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "14"
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
            Height          =   255
            Index           =   2
            Left            =   40
            TabIndex        =   305
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "13"
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
            Height          =   255
            Index           =   1
            Left            =   40
            TabIndex        =   304
            Top             =   840
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "12"
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
            Height          =   255
            Index           =   0
            Left            =   40
            TabIndex        =   303
            Top             =   480
            Width           =   255
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   19
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":29607
            Stretch         =   -1  'True
            Top             =   3240
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   18
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":2AEE8
            Stretch         =   -1  'True
            Top             =   2880
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   17
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":2C7C9
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   16
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":2E0AA
            Stretch         =   -1  'True
            Top             =   2160
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   15
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":2F98B
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   14
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":3126C
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   13
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":32B4D
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   12
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":3442E
            Stretch         =   -1  'True
            Top             =   720
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   11
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":35D0F
            Stretch         =   -1  'True
            Top             =   360
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   10
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":375F0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   450
         End
      End
      Begin VB.PictureBox tabcard 
         Appearance      =   0  'Flat
         BackColor       =   &H00006400&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3735
         Index           =   0
         Left            =   6720
         ScaleHeight     =   3735
         ScaleWidth      =   2175
         TabIndex        =   258
         Top             =   2880
         Width           =   2175
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   9
            Left            =   240
            TabIndex        =   260
            Top             =   3240
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":38ED1
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   8
            Left            =   240
            TabIndex        =   262
            Top             =   2880
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":39833
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   7
            Left            =   240
            TabIndex        =   264
            Top             =   2520
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":3A195
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   6
            Left            =   240
            TabIndex        =   266
            Top             =   2160
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":3AAF7
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   5
            Left            =   240
            TabIndex        =   267
            Top             =   1800
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":3B459
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   4
            Left            =   240
            TabIndex        =   268
            Top             =   1440
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":3BDBB
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   3
            Left            =   240
            TabIndex        =   269
            Top             =   1080
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":3C71D
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   2
            Left            =   240
            TabIndex        =   270
            Top             =   720
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":3D07F
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   1
            Left            =   240
            TabIndex        =   271
            Top             =   360
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":3D9E1
            ImgSize         =   40
            cBack           =   32768
         End
         Begin MoneyReport.lvButtons_H btnlimpiacredit 
            Height          =   405
            Index           =   0
            Left            =   240
            TabIndex        =   272
            Top             =   0
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   714
            CapAlign        =   2
            BackStyle       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            Image           =   "forma_moneyreport.frx":3E343
            ImgSize         =   40
            cBack           =   32768
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC8F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   0
            Left            =   480
            TabIndex        =   259
            Top             =   0
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC70&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   1
            Left            =   480
            TabIndex        =   278
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC8F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   2
            Left            =   480
            TabIndex        =   277
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC70&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   3
            Left            =   480
            TabIndex        =   276
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC8F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   4
            Left            =   480
            TabIndex        =   275
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC70&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   5
            Left            =   480
            TabIndex        =   274
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC8F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   6
            Left            =   480
            TabIndex        =   273
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC70&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   7
            Left            =   480
            TabIndex        =   265
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC8F&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   8
            Left            =   480
            TabIndex        =   263
            Top             =   2880
            Width           =   1215
         End
         Begin VB.TextBox txtcredit_agente 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008FBC70&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   9
            Left            =   480
            TabIndex        =   261
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "1"
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
            Height          =   255
            Index           =   12
            Left            =   100
            TabIndex        =   315
            Top             =   120
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "10"
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
            Height          =   255
            Index           =   11
            Left            =   40
            TabIndex        =   314
            Top             =   3360
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "9"
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
            Height          =   255
            Index           =   10
            Left            =   100
            TabIndex        =   313
            Top             =   3000
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "8"
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
            Height          =   255
            Index           =   9
            Left            =   100
            TabIndex        =   312
            Top             =   2640
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "7"
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
            Height          =   255
            Index           =   8
            Left            =   100
            TabIndex        =   311
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "6"
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
            Height          =   255
            Index           =   7
            Left            =   100
            TabIndex        =   310
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "5"
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
            Height          =   255
            Index           =   6
            Left            =   100
            TabIndex        =   309
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "4"
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
            Height          =   255
            Index           =   5
            Left            =   100
            TabIndex        =   308
            Top             =   1200
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "3"
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
            Height          =   255
            Index           =   4
            Left            =   100
            TabIndex        =   307
            Top             =   840
            Width           =   255
         End
         Begin VB.Label lblnumero 
            BackStyle       =   0  'Transparent
            Caption         =   "2"
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
            Height          =   255
            Index           =   3
            Left            =   100
            TabIndex        =   306
            Top             =   480
            Width           =   255
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   0
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":3ECA5
            Stretch         =   -1  'True
            Top             =   0
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   1
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":40586
            Stretch         =   -1  'True
            Top             =   360
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   2
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":41E67
            Stretch         =   -1  'True
            Top             =   720
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   3
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":43748
            Stretch         =   -1  'True
            Top             =   1080
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   4
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":45029
            Stretch         =   -1  'True
            Top             =   1440
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   5
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":4690A
            Stretch         =   -1  'True
            Top             =   1800
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   6
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":481EB
            Stretch         =   -1  'True
            Top             =   2160
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   7
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":49ACC
            Stretch         =   -1  'True
            Top             =   2520
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   8
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":4B3AD
            Stretch         =   -1  'True
            Top             =   2880
            Width           =   450
         End
         Begin VB.Image img_credit 
            Height          =   405
            Index           =   9
            Left            =   1680
            Picture         =   "forma_moneyreport.frx":4CC8E
            Stretch         =   -1  'True
            Top             =   3240
            Width           =   450
         End
      End
      Begin VB.PictureBox visualizador 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   4815
         Left            =   14520
         ScaleHeight     =   4785
         ScaleWidth      =   5145
         TabIndex        =   117
         Top             =   5760
         Visible         =   0   'False
         Width           =   5175
         Begin MoneyReport.lvButtons_H btnclose_viewer 
            Height          =   375
            Left            =   4320
            TabIndex        =   119
            Top             =   4200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            Caption         =   "Close"
            CapAlign        =   2
            BackStyle       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16777215
            cFHover         =   16777215
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   8421504
         End
         Begin AcroPDFLibCtl.AcroPDF pdf1 
            Height          =   3615
            Left            =   240
            TabIndex        =   118
            Top             =   240
            Visible         =   0   'False
            Width           =   4575
            _cx             =   5080
            _cy             =   5080
         End
         Begin VB.Image img1 
            Height          =   3615
            Left            =   240
            Stretch         =   -1  'True
            Top             =   360
            Visible         =   0   'False
            Width           =   4575
         End
      End
      Begin VB.CheckBox chk_firma_agente 
         BackColor       =   &H00006400&
         Caption         =   "Signed by:"
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
         Left            =   14880
         TabIndex        =   238
         Top             =   9960
         Width           =   1095
      End
      Begin MoneyReport.lvButtons_H btnrefresh_total_lae_agente 
         Height          =   435
         Left            =   11520
         TabIndex        =   237
         Top             =   5040
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   767
         Caption         =   "Update"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFore           =   16777215
         cFHover         =   16777215
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   6908265
      End
      Begin MoneyReport.lvButtons_H btneraser_recibo 
         Height          =   375
         Index           =   1
         Left            =   11880
         TabIndex        =   235
         Top             =   3120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":4E56F
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btneraser_recibo 
         Height          =   375
         Index           =   0
         Left            =   11880
         TabIndex        =   234
         Top             =   2760
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":4EED1
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnclear_cust 
         Height          =   375
         Index           =   1
         Left            =   10320
         TabIndex        =   233
         Top             =   3120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":4F833
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnclear_cust 
         Height          =   375
         Index           =   0
         Left            =   10320
         TabIndex        =   232
         Top             =   2760
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":50195
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnclear_void 
         Height          =   375
         Index           =   1
         Left            =   13200
         TabIndex        =   231
         Top             =   3120
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":50AF7
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnclear_void 
         Height          =   375
         Index           =   0
         Left            =   13200
         TabIndex        =   230
         Top             =   2760
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":51459
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiadebit 
         Height          =   405
         Index           =   9
         Left            =   4800
         TabIndex        =   229
         Top             =   6180
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   714
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":51DBB
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiadebit 
         Height          =   405
         Index           =   8
         Left            =   4800
         TabIndex        =   228
         Top             =   5760
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   714
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":5271D
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiadebit 
         Height          =   405
         Index           =   7
         Left            =   4800
         TabIndex        =   227
         Top             =   5340
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   714
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":5307F
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiadebit 
         Height          =   405
         Index           =   6
         Left            =   4800
         TabIndex        =   226
         Top             =   4940
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   714
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":539E1
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiadebit 
         Height          =   405
         Index           =   5
         Left            =   4800
         TabIndex        =   225
         Top             =   4530
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   714
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":54343
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiadebit 
         Height          =   405
         Index           =   4
         Left            =   4800
         TabIndex        =   224
         Top             =   4120
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   714
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":54CA5
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiadebit 
         Height          =   405
         Index           =   3
         Left            =   4800
         TabIndex        =   223
         Top             =   3720
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   714
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":55607
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiadebit 
         Height          =   405
         Index           =   2
         Left            =   4800
         TabIndex        =   222
         Top             =   3320
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   714
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":55F69
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiadebit 
         Height          =   405
         Index           =   1
         Left            =   4800
         TabIndex        =   221
         Top             =   2890
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   714
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":568CB
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiadebit 
         Height          =   405
         Index           =   0
         Left            =   4800
         TabIndex        =   220
         Top             =   2490
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   714
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":5722D
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiacash 
         Height          =   495
         Index           =   4
         Left            =   0
         TabIndex        =   218
         Top             =   6960
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":57B8F
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiacash 
         Height          =   495
         Index           =   3
         Left            =   0
         TabIndex        =   217
         Top             =   6480
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":584F1
         ImgSize         =   40
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnlimpiacash 
         Height          =   495
         Index           =   2
         Left            =   0
         TabIndex        =   216
         Top             =   6000
         Visible         =   0   'False
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":58E53
         ImgSize         =   40
         cBack           =   32768
      End
      Begin VB.TextBox txtdebit_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   9
         Left            =   5040
         TabIndex        =   31
         Top             =   6165
         Width           =   1215
      End
      Begin VB.TextBox txtdebit_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   8
         Left            =   5040
         TabIndex        =   30
         Top             =   5760
         Width           =   1215
      End
      Begin VB.FileListBox File1 
         Height          =   480
         Left            =   12600
         Pattern         =   "*.pdf;*.bmp;*jpg;*.png;*jpeg;*'gif"
         TabIndex        =   116
         Top             =   3960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtoutput 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   14880
         TabIndex        =   108
         Top             =   6000
         Width           =   3135
      End
      Begin VB.TextBox txttotal_LAE_agente 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   9480
         TabIndex        =   48
         Top             =   4440
         Width           =   2775
      End
      Begin VB.TextBox txtamount_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   12240
         TabIndex        =   11
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox txtamount_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   12240
         TabIndex        =   8
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtcustomer_agente 
         Alignment       =   2  'Center
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   9480
         TabIndex        =   9
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtcustomer_agente 
         Alignment       =   2  'Center
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   9480
         TabIndex        =   6
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtrecibos_agente 
         Alignment       =   2  'Center
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   10680
         TabIndex        =   10
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtrecibos_agente 
         Alignment       =   2  'Center
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   10680
         TabIndex        =   7
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtnotas_agente 
         BackColor       =   &H00006400&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Top             =   8640
         Width           =   4095
      End
      Begin VB.TextBox txtcash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
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
         Index           =   5
         Left            =   1920
         TabIndex        =   5
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtcash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
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
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Top             =   6960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtdebit_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   5040
         TabIndex        =   29
         Top             =   5355
         Width           =   1215
      End
      Begin VB.TextBox txtdebit_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   5040
         TabIndex        =   28
         Top             =   4950
         Width           =   1215
      End
      Begin VB.TextBox txtdebit_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   5040
         TabIndex        =   27
         Top             =   4545
         Width           =   1215
      End
      Begin VB.TextBox txtdebit_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   5040
         TabIndex        =   26
         Top             =   4140
         Width           =   1215
      End
      Begin VB.TextBox txtdebit_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   5040
         TabIndex        =   25
         Top             =   3735
         Width           =   1215
      End
      Begin VB.TextBox txtdebit_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   5040
         TabIndex        =   24
         Top             =   3330
         Width           =   1215
      End
      Begin VB.TextBox txtdebit_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   5040
         TabIndex        =   23
         Top             =   2925
         Width           =   1215
      End
      Begin VB.TextBox txtdebit_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   5040
         TabIndex        =   22
         Top             =   2520
         Width           =   1215
      End
      Begin VB.TextBox txtcash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
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
         Index           =   3
         Left            =   240
         TabIndex        =   3
         Top             =   6480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtcash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   240
         TabIndex        =   2
         Top             =   6000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtcash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
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
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   4920
         Width           =   1695
      End
      Begin VB.TextBox txtcash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
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
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   3360
         Width           =   1695
      End
      Begin MoneyReport.lvButtons_H btn7 
         Height          =   615
         Left            =   14880
         TabIndex        =   49
         Top             =   7080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "7"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   2775
         Left            =   14640
         TabIndex        =   50
         Top             =   2520
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4895
         View            =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   64154
         BorderStyle     =   1
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   0
      End
      Begin MoneyReport.lvButtons_H btn8 
         Height          =   615
         Left            =   15480
         TabIndex        =   51
         Top             =   7080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "8"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btn9 
         Height          =   615
         Left            =   16080
         TabIndex        =   52
         Top             =   7080
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "9"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btn4 
         Height          =   615
         Left            =   14880
         TabIndex        =   53
         Top             =   7680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "4"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btn5 
         Height          =   615
         Left            =   15480
         TabIndex        =   54
         Top             =   7680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "5"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btn6 
         Height          =   615
         Left            =   16080
         TabIndex        =   55
         Top             =   7680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "6"
         CapAlign        =   2
         BackStyle       =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btn1 
         Height          =   615
         Left            =   14880
         TabIndex        =   56
         Top             =   8280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "1"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btn2 
         Height          =   615
         Left            =   15480
         TabIndex        =   57
         Top             =   8280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "2"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btn3 
         Height          =   615
         Left            =   16080
         TabIndex        =   58
         Top             =   8280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "3"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btn0 
         Height          =   615
         Left            =   14880
         TabIndex        =   59
         Top             =   8880
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         Caption         =   "0"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btnpunto 
         Height          =   615
         Left            =   16080
         TabIndex        =   60
         Top             =   8880
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "."
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btnc 
         Height          =   615
         Left            =   16800
         TabIndex        =   61
         Top             =   7080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         Caption         =   "C"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btnmul 
         Height          =   615
         Left            =   16800
         TabIndex        =   62
         Top             =   7680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "*"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btndiv 
         Height          =   615
         Left            =   17400
         TabIndex        =   63
         Top             =   7680
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "/"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btnminus 
         Height          =   615
         Left            =   16800
         TabIndex        =   64
         Top             =   8280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "-"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btnmas 
         Height          =   615
         Left            =   17400
         TabIndex        =   65
         Top             =   8280
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1085
         Caption         =   "+"
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin MoneyReport.lvButtons_H btnigual 
         Height          =   615
         Left            =   16800
         TabIndex        =   66
         Top             =   8880
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         Caption         =   "="
         CapAlign        =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         CapStyle        =   1
         Mode            =   0
         Value           =   0   'False
         cBack           =   14737632
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Electronic Payments"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   1215
         Index           =   0
         Left            =   1800
         TabIndex        =   71
         Top             =   1920
         Width           =   2295
      End
      Begin VB.Image Image10 
         Height          =   735
         Left            =   240
         Picture         =   "forma_moneyreport.frx":597B5
         Top             =   3960
         Width           =   3690
      End
      Begin VB.Label lbltotal_cash_agente 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   73
         Top             =   6000
         Width           =   2535
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   480
         TabIndex        =   72
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00008080&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   5
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   4095
      End
      Begin VB.Shape Shape19 
         BorderColor     =   &H00FFFFFF&
         Height          =   495
         Left            =   16320
         Shape           =   4  'Rounded Rectangle
         Top             =   6480
         Width           =   1695
      End
      Begin VB.Label signo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   17640
         TabIndex        =   255
         Top             =   6600
         Width           =   375
      End
      Begin VB.Label Firma_agente 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Freestyle Script"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   16200
         TabIndex        =   239
         Top             =   9915
         Width           =   3255
      End
      Begin VB.Shape Shape9 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   40
         Left            =   0
         Top             =   10520
         Width           =   21135
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fill out this section with your vouchers in hand"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   255
         Left            =   4800
         TabIndex        =   123
         Top             =   7800
         Width           =   3975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   """Send it to the corp. office"""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   122
         Top             =   5240
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   """Send it to the corporate office"""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Index           =   1
         Left            =   840
         TabIndex        =   121
         Top             =   6480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   """Send it to the corp. office"""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   120
         Top             =   4520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Image img_debito 
         Height          =   405
         Index           =   9
         Left            =   6285
         Picture         =   "forma_moneyreport.frx":5B6F7
         Stretch         =   -1  'True
         Top             =   6165
         Width           =   450
      End
      Begin VB.Image img_debito 
         Height          =   405
         Index           =   8
         Left            =   6285
         Picture         =   "forma_moneyreport.frx":5CC41
         Stretch         =   -1  'True
         Top             =   5760
         Width           =   450
      End
      Begin VB.Image img_debito 
         Height          =   405
         Index           =   7
         Left            =   6285
         Picture         =   "forma_moneyreport.frx":5E18B
         Stretch         =   -1  'True
         Top             =   5355
         Width           =   450
      End
      Begin VB.Image img_debito 
         Height          =   405
         Index           =   6
         Left            =   6285
         Picture         =   "forma_moneyreport.frx":5F6D5
         Stretch         =   -1  'True
         Top             =   4950
         Width           =   450
      End
      Begin VB.Image img_debito 
         Height          =   405
         Index           =   5
         Left            =   6285
         Picture         =   "forma_moneyreport.frx":60C1F
         Stretch         =   -1  'True
         Top             =   4545
         Width           =   450
      End
      Begin VB.Image img_debito 
         Height          =   405
         Index           =   4
         Left            =   6285
         Picture         =   "forma_moneyreport.frx":62169
         Stretch         =   -1  'True
         Top             =   4140
         Width           =   450
      End
      Begin VB.Image img_debito 
         Height          =   405
         Index           =   3
         Left            =   6285
         Picture         =   "forma_moneyreport.frx":636B3
         Stretch         =   -1  'True
         Top             =   3735
         Width           =   450
      End
      Begin VB.Image img_debito 
         Height          =   405
         Index           =   2
         Left            =   6285
         Picture         =   "forma_moneyreport.frx":64BFD
         Stretch         =   -1  'True
         Top             =   3330
         Width           =   450
      End
      Begin VB.Image img_debito 
         Height          =   405
         Index           =   1
         Left            =   6285
         Picture         =   "forma_moneyreport.frx":66147
         Stretch         =   -1  'True
         Top             =   2925
         Width           =   450
      End
      Begin VB.Image img_debito 
         Height          =   400
         Index           =   0
         Left            =   6285
         Picture         =   "forma_moneyreport.frx":67691
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   450
      End
      Begin VB.Image btnerase_archivo 
         Height          =   1695
         Left            =   19560
         Picture         =   "forma_moneyreport.frx":68BDB
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00696969&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   0
         Left            =   9240
         Shape           =   4  'Rounded Rectangle
         Top             =   4200
         Width           =   3255
      End
      Begin VB.Label mem 
         Alignment       =   2  'Center
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   16440
         TabIndex        =   103
         Top             =   6675
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   3735
         Left            =   14640
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   3615
      End
      Begin VB.Image Image6 
         Height          =   1575
         Left            =   11640
         Picture         =   "forma_moneyreport.frx":6BB49
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Image Image5 
         Height          =   1095
         Left            =   6120
         Picture         =   "forma_moneyreport.frx":6DF41
         Stretch         =   -1  'True
         Top             =   1275
         Width           =   1215
      End
      Begin VB.Image Image4 
         Height          =   2295
         Left            =   120
         Picture         =   "forma_moneyreport.frx":767A1
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "You can drag and drop the files here..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   15360
         TabIndex        =   102
         Top             =   5400
         Width           =   3975
      End
      Begin VB.Image btnopen1 
         Height          =   1695
         Left            =   19560
         Picture         =   "forma_moneyreport.frx":7885B
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   1455
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attached files and / or images"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   330
         Index           =   4
         Left            =   14760
         TabIndex        =   101
         Top             =   1875
         Width           =   4785
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   12
         Left            =   14640
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label lbltotal_over_short_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   12240
         TabIndex        =   100
         Top             =   9480
         Width           =   1815
      End
      Begin VB.Label lblgrantotal_cash_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   12240
         TabIndex        =   99
         Top             =   8640
         Width           =   1815
      End
      Begin VB.Label lblgrantotal_debitcredit_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   12240
         TabIndex        =   98
         Top             =   8160
         Width           =   1815
      End
      Begin VB.Label lbltotal_reportado_menos_voids_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   12240
         TabIndex        =   97
         Top             =   7320
         Width           =   1815
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbltotal_void_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   12240
         TabIndex        =   96
         Top             =   6600
         Width           =   1815
      End
      Begin VB.Label lbltotal_reported_agent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   12240
         TabIndex        =   95
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "OVER (SHORT)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   9360
         TabIndex        =   94
         Top             =   9600
         Width           =   2655
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cash"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   9360
         TabIndex        =   93
         Top             =   8760
         Width           =   2655
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Debit and Credit Sales"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   9240
         TabIndex        =   92
         Top             =   8280
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Totals"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   9360
         TabIndex        =   91
         Top             =   7380
         Width           =   2655
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Pending Void Receipts"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   9240
         TabIndex        =   90
         Top             =   6720
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Reported"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   9600
         TabIndex        =   89
         Top             =   6120
         Width           =   2445
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T O T A L"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9480
         TabIndex        =   88
         Top             =   3480
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbltotal_recibos_void_agente 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11640
         TabIndex        =   87
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002E8B57&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amount"
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
         Height          =   255
         Index           =   2
         Left            =   12000
         TabIndex        =   86
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002E8B57&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Receipt #"
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
         Height          =   255
         Index           =   1
         Left            =   10560
         TabIndex        =   85
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipts/Pending"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   330
         Index           =   3
         Left            =   9720
         TabIndex        =   84
         Top             =   1920
         Width           =   2445
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002E8B57&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Customer ID"
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
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   83
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   11
         Left            =   9480
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   3855
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Send suspense to Receipts/Corrections for any corrections or adjustments."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   360
         TabIndex        =   82
         Top             =   8160
         Width           =   3780
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "If you are over explain where this money came from.  If you are short explain who has this money and why?"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   420
         Index           =   0
         Left            =   360
         TabIndex        =   81
         Top             =   7680
         Width           =   3825
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   80
         Top             =   7340
         Width           =   705
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   3615
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   7200
         Width           =   4335
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   5040
         TabIndex        =   79
         Top             =   8640
         Width           =   1095
      End
      Begin VB.Label lbltotal_debit_credit_agent 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6240
         TabIndex        =   78
         Top             =   8520
         Width           =   2535
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00008080&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   10
         Left            =   4680
         Shape           =   4  'Rounded Rectangle
         Top             =   8280
         Width           =   4335
      End
      Begin VB.Label lbltotal_credit_agent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H002E8B57&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6960
         TabIndex        =   77
         Top             =   6960
         Width           =   1935
      End
      Begin VB.Label lbltotal_debit_agent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H002E8B57&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   76
         Top             =   6960
         Width           =   1935
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00004000&
         BorderWidth     =   2
         FillColor       =   &H00006400&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   8
         Left            =   4680
         Shape           =   4  'Rounded Rectangle
         Top             =   6720
         Width           =   4335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   480
         Index           =   2
         Left            =   7320
         TabIndex        =   75
         Top             =   1800
         Width           =   1245
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   480
         Index           =   1
         Left            =   5020
         TabIndex        =   74
         Top             =   1800
         Width           =   1065
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   9
         Left            =   6840
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   6
         Left            =   4680
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         Height          =   4695
         Index           =   7
         Left            =   4680
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   1455
         Index           =   0
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         Height          =   735
         Index           =   4
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   6360
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         Height          =   735
         Index           =   3
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   4800
         Width           =   4095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Coins"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   70
         Top             =   4220
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Checks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   69
         Top             =   6360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         Height          =   735
         Index           =   2
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   4080
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Money Order"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   68
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         Height          =   1575
         Index           =   1
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3240
         Width           =   4095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Slip Cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   67
         Top             =   3480
         Width           =   975
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0FFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   0
         Top             =   1320
         Width           =   21135
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   5775
         Left            =   9240
         Shape           =   4  'Rounded Rectangle
         Top             =   5760
         Width           =   5055
      End
      Begin VB.Image Image1 
         Height          =   1455
         Left            =   12420
         Picture         =   "forma_moneyreport.frx":7C4D2
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Shape Shape17 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   16080
         Shape           =   4  'Rounded Rectangle
         Top             =   9840
         Width           =   3495
      End
   End
   Begin VB.PictureBox Hoja2 
      Appearance      =   0  'Flat
      BackColor       =   &H00006400&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9255
      Left            =   240
      ScaleHeight     =   9255
      ScaleWidth      =   21375
      TabIndex        =   107
      Top             =   2760
      Visible         =   0   'False
      Width           =   21375
      Begin VB.Frame panel 
         BackColor       =   &H00006400&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2640
         TabIndex        =   244
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
         Begin VB.OptionButton op_view 
            BackColor       =   &H00006400&
            Caption         =   "Show all"
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
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   246
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton op_view 
            BackColor       =   &H00006400&
            Caption         =   "Personal only"
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   245
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid3 
         Height          =   1695
         Left            =   120
         TabIndex        =   112
         Top             =   6240
         Width           =   21135
         _ExtentX        =   37280
         _ExtentY        =   2990
         _Version        =   393216
         BackColor       =   14737632
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorBkg    =   25600
         BorderStyle     =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid1 
         Height          =   4935
         Left            =   120
         TabIndex        =   109
         Top             =   600
         Width           =   21255
         _ExtentX        =   37491
         _ExtentY        =   8705
         _Version        =   393216
         BackColor       =   9419919
         BackColorFixed  =   32768
         ForeColorFixed  =   16777215
         BackColorBkg    =   25600
         BorderStyle     =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VOID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   2
         Left            =   360
         TabIndex        =   115
         Top             =   5760
         Width           =   840
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipts:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   114
         Top             =   5880
         Width           =   960
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipts:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   113
         Top             =   140
         Width           =   1275
      End
      Begin VB.Image btncarga_datos_agente 
         Height          =   1215
         Left            =   19560
         Picture         =   "forma_moneyreport.frx":80C33
         Stretch         =   -1  'True
         Top             =   8040
         Width           =   1095
      End
   End
   Begin VB.PictureBox hoja3 
      Appearance      =   0  'Flat
      BackColor       =   &H00006400&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   240
      ScaleHeight     =   9375
      ScaleWidth      =   21495
      TabIndex        =   124
      Top             =   2640
      Width           =   21495
      Begin VB.PictureBox visualizador2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         ForeColor       =   &H80000008&
         Height          =   5295
         Left            =   16320
         ScaleHeight     =   5265
         ScaleWidth      =   5145
         TabIndex        =   241
         Top             =   3960
         Visible         =   0   'False
         Width           =   5175
         Begin MoneyReport.lvButtons_H btncierra_visualizador2 
            Height          =   375
            Left            =   4320
            TabIndex        =   242
            Top             =   4200
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            Caption         =   "Close"
            CapAlign        =   2
            BackStyle       =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFore           =   16777215
            cFHover         =   16777215
            cGradient       =   0
            Mode            =   0
            Value           =   0   'False
            cBack           =   8421504
         End
         Begin AcroPDFLibCtl.AcroPDF pdf2 
            Height          =   3615
            Left            =   240
            TabIndex        =   243
            Top             =   240
            Visible         =   0   'False
            Width           =   4575
            _cx             =   5080
            _cy             =   5080
         End
         Begin VB.Image img2 
            Height          =   3615
            Left            =   240
            Stretch         =   -1  'True
            Top             =   360
            Visible         =   0   'False
            Width           =   4575
         End
      End
      Begin MoneyReport.lvButtons_H btnerasecredit 
         Height          =   495
         Index           =   1
         Left            =   4440
         TabIndex        =   213
         Top             =   6960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":83BF7
         ImgSize         =   24
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnerasecredit 
         Height          =   495
         Index           =   0
         Left            =   2480
         TabIndex        =   212
         Top             =   6960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":84559
         ImgSize         =   24
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnborracash 
         Height          =   495
         Index           =   2
         Left            =   4440
         TabIndex        =   211
         Top             =   3240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":84EBB
         ImgSize         =   24
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnborracash 
         Height          =   495
         Index           =   1
         Left            =   4440
         TabIndex        =   210
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":8581D
         ImgSize         =   24
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnborracash 
         Height          =   495
         Index           =   0
         Left            =   4440
         TabIndex        =   209
         Top             =   1800
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   873
         CapAlign        =   2
         BackStyle       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         Image           =   "forma_moneyreport.frx":8617F
         ImgSize         =   24
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnerase_combo 
         Height          =   375
         Index           =   5
         Left            =   21000
         TabIndex        =   208
         Top             =   5880
         Width           =   320
         _ExtentX        =   556
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
         Image           =   "forma_moneyreport.frx":86AE1
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnerase_combo 
         Height          =   375
         Index           =   4
         Left            =   21000
         TabIndex        =   207
         Top             =   5520
         Width           =   320
         _ExtentX        =   556
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
         Image           =   "forma_moneyreport.frx":87443
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnerase_combo 
         Height          =   375
         Index           =   3
         Left            =   21000
         TabIndex        =   206
         Top             =   5160
         Width           =   320
         _ExtentX        =   556
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
         Image           =   "forma_moneyreport.frx":87DA5
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnerase_combo 
         Height          =   375
         Index           =   2
         Left            =   20880
         TabIndex        =   205
         Top             =   2400
         Width           =   320
         _ExtentX        =   556
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
         Image           =   "forma_moneyreport.frx":88707
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnerase_combo 
         Height          =   375
         Index           =   1
         Left            =   20880
         TabIndex        =   204
         Top             =   2040
         Width           =   320
         _ExtentX        =   556
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
         Image           =   "forma_moneyreport.frx":89069
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnerase_combo 
         Height          =   375
         Index           =   0
         Left            =   20880
         TabIndex        =   203
         Top             =   1680
         Width           =   320
         _ExtentX        =   556
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
         Image           =   "forma_moneyreport.frx":899CB
         cBack           =   32768
      End
      Begin MoneyReport.lvButtons_H btnupdate_LAE 
         Height          =   375
         Left            =   8400
         TabIndex        =   202
         Top             =   1320
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "Update"
         CapAlign        =   2
         BackStyle       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cGradient       =   0
         Mode            =   0
         Value           =   0   'False
         cBack           =   8421504
      End
      Begin VB.ComboBox cbo_employees 
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   16440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   160
         Top             =   5940
         Width           =   1695
      End
      Begin VB.ComboBox cbo_employees 
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   16440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   157
         Top             =   5550
         Width           =   1695
      End
      Begin VB.ComboBox cbo_employees 
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   16440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   154
         Top             =   5160
         Width           =   1695
      End
      Begin VB.ComboBox cbo_employees 
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   16440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   151
         Top             =   2480
         Width           =   1695
      End
      Begin VB.ComboBox cbo_employees 
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   16440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   148
         Top             =   2080
         Width           =   1695
      End
      Begin VB.ComboBox cbo_employees 
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   16440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   145
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ComboBox cbooficina1 
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   18120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   161
         Top             =   5940
         Width           =   1935
      End
      Begin VB.ComboBox cbooficina1 
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   18120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   158
         Top             =   5550
         Width           =   1935
      End
      Begin VB.ComboBox cbooficina1 
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   18120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   155
         Top             =   5160
         Width           =   1935
      End
      Begin VB.ComboBox cbooficina1 
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   18120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   152
         Top             =   2480
         Width           =   1815
      End
      Begin VB.ComboBox cbooficina1 
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   18120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   149
         Top             =   2080
         Width           =   1815
      End
      Begin VB.ComboBox cbooficina1 
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   18120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   146
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox chkmanager 
         BackColor       =   &H00006400&
         Caption         =   "Reviewed by Manager"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   17640
         TabIndex        =   201
         Top             =   8880
         Width           =   2655
      End
      Begin VB.TextBox txtcant_venida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
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
         Index           =   2
         Left            =   20040
         TabIndex        =   163
         Top             =   5910
         Width           =   975
      End
      Begin VB.TextBox txtcant_venida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
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
         Index           =   1
         Left            =   20040
         TabIndex        =   159
         Top             =   5530
         Width           =   975
      End
      Begin VB.TextBox txtcant_venida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
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
         Index           =   0
         Left            =   20040
         TabIndex        =   156
         Top             =   5160
         Width           =   975
      End
      Begin VB.TextBox txtcant_ida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
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
         Index           =   2
         Left            =   19920
         TabIndex        =   153
         Top             =   2430
         Width           =   975
      End
      Begin VB.TextBox txtcant_ida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
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
         Index           =   1
         Left            =   19920
         TabIndex        =   150
         Top             =   2060
         Width           =   975
      End
      Begin VB.TextBox txtcant_ida 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
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
         Index           =   0
         Left            =   19920
         TabIndex        =   147
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtnotas_manager 
         Appearance      =   0  'Flat
         BackColor       =   &H00006400&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   11520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   186
         Top             =   2160
         Width           =   4455
      End
      Begin VB.TextBox txttotal_venta_manager 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6360
         TabIndex        =   172
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtcredit_manager 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3040
         TabIndex        =   144
         Top             =   6960
         Width           =   1400
      End
      Begin VB.TextBox txtdebit_manager 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   143
         Top             =   6960
         Width           =   1400
      End
      Begin VB.TextBox txtdinero 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
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
         Index           =   2
         Left            =   2760
         TabIndex        =   142
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox txtdinero 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC70&
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
         Index           =   1
         Left            =   2760
         TabIndex        =   141
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtdinero 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008FBC8F&
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
         Index           =   0
         Left            =   2760
         TabIndex        =   140
         Top             =   1800
         Width           =   1695
      End
      Begin ComctlLib.ListView ListView2 
         Height          =   2175
         Left            =   11400
         TabIndex        =   240
         Top             =   5760
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3836
         View            =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   64154
         BorderStyle     =   1
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   0
      End
      Begin VB.Image Image11 
         Height          =   735
         Left            =   960
         Picture         =   "forma_moneyreport.frx":8A32D
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Attached files and / or images"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   495
         Left            =   11520
         TabIndex        =   247
         Top             =   5160
         Width           =   4455
      End
      Begin VB.Image btnborrar_archivo2 
         Height          =   1455
         Left            =   13800
         Picture         =   "forma_moneyreport.frx":8C26F
         Stretch         =   -1  'True
         Top             =   7860
         Width           =   1215
      End
      Begin VB.Image btnopen2 
         Height          =   1455
         Left            =   14880
         Picture         =   "forma_moneyreport.frx":8F1DD
         Stretch         =   -1  'True
         Top             =   7860
         Width           =   1215
      End
      Begin VB.Image arrow 
         Height          =   255
         Index           =   1
         Left            =   7440
         Picture         =   "forma_moneyreport.frx":92E54
         Stretch         =   -1  'True
         Top             =   6600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image arrow 
         Height          =   240
         Index           =   0
         Left            =   6600
         Picture         =   "forma_moneyreport.frx":93296
         Stretch         =   -1  'True
         Top             =   6600
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label firma 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Freestyle Script"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   17760
         TabIndex        =   236
         Top             =   7800
         Width           =   3015
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H0000C000&
         BorderWidth     =   5
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1335
         Left            =   17520
         Shape           =   4  'Rounded Rectangle
         Top             =   7440
         Width           =   3495
      End
      Begin VB.Label lblover_short_oficina 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FF80FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   8280
         TabIndex        =   182
         Top             =   6720
         Width           =   2535
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Over / Short:"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   345
         Index           =   4
         Left            =   6360
         TabIndex        =   183
         Top             =   6840
         Width           =   1860
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H00000000&
         BorderStyle     =   0  'Transparent
         BorderWidth     =   2
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   6000
         Top             =   6600
         Width           =   5025
      End
      Begin VB.Label lbltotal_agentes_que_vinieron 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   19320
         TabIndex        =   200
         Top             =   6480
         Width           =   1575
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   18360
         TabIndex        =   199
         Top             =   6555
         Width           =   855
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00008080&
         FillStyle       =   0  'Solid
         Height          =   855
         Index           =   22
         Left            =   18360
         Shape           =   4  'Rounded Rectangle
         Top             =   6135
         Width           =   2655
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   20040
         TabIndex        =   198
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Office"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   18240
         TabIndex        =   197
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Agent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   16800
         TabIndex        =   196
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Agent came to help us:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   495
         Index           =   1
         Left            =   16680
         TabIndex        =   195
         Top             =   4200
         Width           =   4335
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   16680
         Shape           =   4  'Rounded Rectangle
         Top             =   4080
         Width           =   4335
      End
      Begin VB.Label lbltotal_agentes_idos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   19200
         TabIndex        =   194
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   18360
         TabIndex        =   193
         Top             =   3080
         Width           =   735
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   19920
         TabIndex        =   192
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Office"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   18120
         TabIndex        =   191
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Agent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   16800
         TabIndex        =   190
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "My Agent went to:"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   615
         Index           =   0
         Left            =   16680
         TabIndex        =   189
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N O T E S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11640
         TabIndex        =   188
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Attach VOID RECEIPT FORM and set suspense to UW for any adjustments/corrections."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   11760
         TabIndex        =   187
         Top             =   1080
         Width           =   3930
         WordWrap        =   -1  'True
      End
      Begin VB.Label lbltotal_needed_oficina 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   8280
         TabIndex        =   185
         Top             =   4440
         Width           =   2535
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cash:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   6360
         TabIndex        =   184
         Top             =   4560
         Width           =   600
      End
      Begin VB.Label lbltotal_dejado_por_agentes_que_vinieron 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   8280
         TabIndex        =   181
         Top             =   5880
         Width           =   2535
      End
      Begin VB.Label lbltotal_dejado_por_agentes_de_oficina 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   8280
         TabIndex        =   180
         Top             =   5160
         Width           =   2535
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Came:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   6360
         TabIndex        =   179
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "My agent left:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6360
         TabIndex        =   178
         Top             =   5280
         Width           =   1455
      End
      Begin VB.Label lbltotal_debit_and_credit_oficina 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   8280
         TabIndex        =   177
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit and Credit:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   6360
         TabIndex        =   176
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label lbltotal_LAE_oficina 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   8280
         TabIndex        =   175
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total LAE System:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   6360
         TabIndex        =   174
         Top             =   3120
         Width           =   1965
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "TOTALS"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   6960
         TabIndex        =   173
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   6000
         Shape           =   4  'Rounded Rectangle
         Top             =   2040
         Width           =   5055
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H0000C000&
         BorderWidth     =   3
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   5655
         Left            =   6000
         Shape           =   4  'Rounded Rectangle
         Top             =   2160
         Width           =   5055
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00696969&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   1
         Left            =   6120
         Shape           =   4  'Rounded Rectangle
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label lbltotal_debito_credito_manager 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   171
         Top             =   8040
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   1320
         TabIndex        =   170
         Top             =   8160
         Width           =   1095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         Height          =   735
         Index           =   20
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   6840
         Width           =   3975
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00008080&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   19
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   7800
         Width           =   4095
      End
      Begin VB.Image Image8 
         Height          =   1095
         Left            =   2280
         Picture         =   "forma_moneyreport.frx":936D8
         Stretch         =   -1  'True
         Top             =   5600
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   480
         Index           =   7
         Left            =   3480
         TabIndex        =   169
         Top             =   6120
         Width           =   1245
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   1200
         TabIndex        =   168
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label lbltotal_cash_manager 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   167
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00008080&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   18
         Left            =   840
         Shape           =   4  'Rounded Rectangle
         Top             =   4080
         Width           =   4095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   """Send it to the corporate office"""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Index           =   3
         Left            =   960
         TabIndex        =   166
         Top             =   2880
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   """Send it to the corporate office"""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Index           =   4
         Left            =   960
         TabIndex        =   165
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Money Order:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   960
         TabIndex        =   164
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Coins:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   960
         TabIndex        =   162
         Top             =   2520
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         Height          =   735
         Index           =   17
         Left            =   840
         Shape           =   4  'Rounded Rectangle
         Top             =   3120
         Width           =   4095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         Height          =   735
         Index           =   16
         Left            =   840
         Shape           =   4  'Rounded Rectangle
         Top             =   2400
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00000000&
         Height          =   1455
         Index           =   15
         Left            =   840
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Slip Cash:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   6
         Left            =   960
         TabIndex        =   139
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Debit"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   480
         Index           =   6
         Left            =   1200
         TabIndex        =   138
         Top             =   6120
         Width           =   1065
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   14
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   6000
         Width           =   3975
      End
      Begin VB.Image Image7 
         Height          =   2295
         Left            =   960
         Picture         =   "forma_moneyreport.frx":9BF38
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Electronic Payments"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   1215
         Index           =   5
         Left            =   2640
         TabIndex        =   137
         Top             =   360
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   1455
         Index           =   13
         Left            =   840
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   4095
      End
      Begin VB.Shape Shape12 
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   3975
         Left            =   11280
         Shape           =   4  'Rounded Rectangle
         Top             =   720
         Width           =   4935
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   16560
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   4335
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00004000&
         BorderWidth     =   3
         FillColor       =   &H00008080&
         FillStyle       =   0  'Solid
         Height          =   975
         Index           =   21
         Left            =   18240
         Shape           =   4  'Rounded Rectangle
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Image Image9 
         Height          =   1935
         Left            =   9120
         Picture         =   "forma_moneyreport.frx":9DFF2
         Stretch         =   -1  'True
         Top             =   600
         Width           =   2175
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   3
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   735
         Index           =   23
         Left            =   11400
         Shape           =   4  'Rounded Rectangle
         Top             =   4980
         Width           =   4695
      End
   End
   Begin VB.Image Image12 
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Image img_palomita_down 
      Height          =   615
      Left            =   13560
      Picture         =   "forma_moneyreport.frx":A2753
      Stretch         =   -1  'True
      Top             =   13680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_palomita_up 
      Height          =   615
      Left            =   12840
      Picture         =   "forma_moneyreport.frx":A4A82
      Stretch         =   -1  'True
      Top             =   13680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_tab1 
      Height          =   855
      Index           =   2
      Left            =   18840
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label leyenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Money report"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   3
      Left            =   19400
      TabIndex        =   126
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label leyenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Manager"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   495
      Index           =   2
      Left            =   19080
      TabIndex        =   125
      Top             =   960
      Width           =   1230
   End
   Begin VB.Shape tabx 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   2
      Left            =   18840
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1935
   End
   Begin VB.Image btnsend 
      Height          =   1695
      Left            =   8280
      Picture         =   "forma_moneyreport.frx":A873C
      Stretch         =   -1  'True
      Top             =   12120
      Width           =   1455
   End
   Begin VB.Image img_send_down 
      Height          =   615
      Left            =   22560
      Picture         =   "forma_moneyreport.frx":AB4EB
      Stretch         =   -1  'True
      Top             =   12000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_send_up 
      Height          =   615
      Left            =   21960
      Picture         =   "forma_moneyreport.frx":AEB8F
      Stretch         =   -1  'True
      Top             =   12000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image discover 
      Height          =   390
      Left            =   7080
      Picture         =   "forma_moneyreport.frx":B193E
      Top             =   14280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image debito 
      Height          =   390
      Left            =   6360
      Picture         =   "forma_moneyreport.frx":B2E93
      Top             =   14280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image master 
      Height          =   480
      Left            =   5640
      Picture         =   "forma_moneyreport.frx":B43DD
      Top             =   14280
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image american 
      Height          =   390
      Left            =   4920
      Picture         =   "forma_moneyreport.frx":B5CBE
      Top             =   14280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image visa 
      Height          =   390
      Left            =   4200
      Picture         =   "forma_moneyreport.frx":B7382
      Top             =   14280
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Image img_corta_down 
      Height          =   615
      Left            =   22560
      Picture         =   "forma_moneyreport.frx":B8A45
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   615
   End
   Begin VB.Image img_corta_up 
      Height          =   615
      Left            =   21960
      Picture         =   "forma_moneyreport.frx":BB6C5
      Stretch         =   -1  'True
      Top             =   10200
      Width           =   615
   End
   Begin VB.Image img_caja 
      Height          =   615
      Left            =   7800
      Picture         =   "forma_moneyreport.frx":BE633
      Stretch         =   -1  'True
      Top             =   14160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_carga_datos_down 
      Height          =   615
      Left            =   22560
      Picture         =   "forma_moneyreport.frx":C0E37
      Stretch         =   -1  'True
      Top             =   9480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_carga_datos_up 
      Height          =   615
      Left            =   21960
      Picture         =   "forma_moneyreport.frx":C398D
      Stretch         =   -1  'True
      Top             =   9480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_borra_down 
      Height          =   615
      Left            =   22560
      Picture         =   "forma_moneyreport.frx":C6951
      Stretch         =   -1  'True
      Top             =   8880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_borra_up 
      Height          =   615
      Left            =   21960
      Picture         =   "forma_moneyreport.frx":C96DA
      Stretch         =   -1  'True
      Top             =   8880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image btnborra 
      Height          =   1695
      Left            =   6960
      Picture         =   "forma_moneyreport.frx":CD666
      Stretch         =   -1  'True
      Top             =   12120
      Width           =   1455
   End
   Begin VB.Image img_tab1 
      Height          =   855
      Index           =   1
      Left            =   16200
      Top             =   600
      Width           =   2535
   End
   Begin VB.Image img_tab1 
      Height          =   855
      Index           =   0
      Left            =   13680
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label leyenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Receipts/Report"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   495
      Index           =   1
      Left            =   16425
      TabIndex        =   106
      Top             =   960
      Width           =   2115
   End
   Begin VB.Label leyenda 
      BackStyle       =   0  'Transparent
      Caption         =   "Money Report"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   615
      Index           =   0
      Left            =   14040
      TabIndex        =   105
      Top             =   960
      Width           =   1935
   End
   Begin VB.Shape tabx 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   1
      Left            =   16200
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   2535
   End
   Begin VB.Shape tabx 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   0
      Left            =   13800
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   2295
   End
   Begin VB.Image img_load_down 
      Height          =   615
      Left            =   21240
      Picture         =   "forma_moneyreport.frx":D15F2
      Stretch         =   -1  'True
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_load_up 
      Height          =   615
      Left            =   20640
      Picture         =   "forma_moneyreport.frx":D42F6
      Stretch         =   -1  'True
      Top             =   10680
      Visible         =   0   'False
      Width           =   615
   End
   Begin ComctlLib.ImageList imgLarge 
      Left            =   720
      Top             =   14040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "forma_moneyreport.frx":D7F6D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "forma_moneyreport.frx":FD1E3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image img_end_down 
      Height          =   615
      Left            =   21240
      Picture         =   "forma_moneyreport.frx":122459
      Stretch         =   -1  'True
      Top             =   10080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_end_up 
      Height          =   615
      Left            =   20640
      Picture         =   "forma_moneyreport.frx":125956
      Stretch         =   -1  'True
      Top             =   10080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image btnend 
      Height          =   1695
      Left            =   19320
      Picture         =   "forma_moneyreport.frx":12936A
      Stretch         =   -1  'True
      Top             =   12120
      Width           =   1455
   End
   Begin VB.Image btnsave 
      Height          =   1695
      Left            =   9600
      Picture         =   "forma_moneyreport.frx":12CD7E
      Stretch         =   -1  'True
      Top             =   12120
      Width           =   1455
   End
   Begin VB.Image img_disk_down 
      Height          =   615
      Left            =   21240
      Picture         =   "forma_moneyreport.frx":130C5F
      Stretch         =   -1  'True
      Top             =   9480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_disk_up 
      Height          =   615
      Left            =   20640
      Picture         =   "forma_moneyreport.frx":1339F3
      Stretch         =   -1  'True
      Top             =   9480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_calendar_down 
      Height          =   615
      Left            =   21240
      Picture         =   "forma_moneyreport.frx":1378D4
      Stretch         =   -1  'True
      Top             =   8880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_calendar_up 
      Height          =   615
      Left            =   20640
      Picture         =   "forma_moneyreport.frx":13AE85
      Stretch         =   -1  'True
      Top             =   8880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
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
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   44
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image img_print_down 
      Height          =   615
      Left            =   21240
      Picture         =   "forma_moneyreport.frx":13EB14
      Stretch         =   -1  'True
      Top             =   12000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image img_print_up 
      Height          =   615
      Left            =   20640
      Picture         =   "forma_moneyreport.frx":141B21
      Stretch         =   -1  'True
      Top             =   12000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image btnprinter 
      Height          =   1695
      Left            =   10920
      Picture         =   "forma_moneyreport.frx":1458E7
      Stretch         =   -1  'True
      Top             =   12120
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   15600
      Picture         =   "forma_moneyreport.frx":1496AD
      Stretch         =   -1  'True
      Top             =   12000
      Width           =   4095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supervised by: Cintia Cadena"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   4440
      TabIndex        =   42
      Top             =   13440
      Width           =   2130
   End
   Begin VB.Label lbladmon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8880
      TabIndex        =   41
      Top             =   960
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lbladministrator 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   10440
      TabIndex        =   40
      Top             =   960
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "IT Department"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   3120
      TabIndex        =   39
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Printer:"
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
      Height          =   255
      Index           =   2
      Left            =   12360
      TabIndex        =   33
      Top             =   12600
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Created by: Hector Navarro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   1680
      TabIndex        =   21
      Top             =   13440
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manager:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   9240
      TabIndex        =   19
      Top             =   360
      Width           =   990
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (C) 2022-2024"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1680
      TabIndex        =   18
      Top             =   12840
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Just Auto Insurance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   1680
      TabIndex        =   17
      Top             =   13080
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Agent:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   15
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Office:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   360
      Width           =   855
   End
   Begin VB.Image Image3 
      Height          =   2295
      Left            =   240
      Picture         =   "forma_moneyreport.frx":152249
      Stretch         =   -1  'True
      Top             =   11520
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim DesignX As Integer
      Dim DesignY As Integer
Dim primeravez As Integer

Dim lae_office$, estado_carga As Integer, point_date As Integer, permiso_carga As Integer, estado_registro As Integer, arranque As Integer
Dim total_cash As Single, total_debit As Single, total_credit As Single, total_voids As Single, total_general As Single, oficina_trabajada$

Dim operator As Integer
Dim result As Single
Dim sd, seg As Integer
Dim fecha_actual$, id_moneyreport$, fecha_entrada$, segundos As Integer
Dim total_reportado As Single, total_final As Single, total_over As Single, modificado As Integer
Dim guarda$, archivo_selecto$, ruta_archivos$, cargo_impresion As Integer, tipo_guardado As Integer, tipo_vista As Integer

' Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim submitido As Boolean

Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long


Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
  
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Private Const REG_EXPAND_SZ As Long = 4
  
  
Dim OReg As Registro

Public Sub carga_agentes()
On Error Resume Next

If cbo_agentes.Enabled = False Then
   Exit Sub
End If


Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
   
    
    
    Set Rs = New ADODB.Recordset
           
  
  oficina$ = LTrim(RTrim(Right(cbo_oficina.List(cbo_oficina.ListIndex), 30)))
  
  lbloficina_agente.Caption = oficina$
  
  lbloficina_agente.Caption = oficina_trabajada$
  
   sSelect = "select idoffice from officescatalog where office='" + oficina$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_oficina$ = Rs(0)
    Rs.Close
    
    
   fecha_de_revision$ = Format(Now, "mm/dd/yyyy")
  
  Grid2.Clear
    
  
 
  
  If chkagentes(1).Value = 0 Then  ' carga agentes activos y/o inactivos
  
 
  sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (3,6,16,17,18,28,2,29,24,37) and ofc.office='" + oficina$ + "' and empjob.Active='1'"


   
  ' sSelect = "select distinct IDEmployee, Username from EmployeeInfo emp " & _
  ' "inner join ReceiptsHDR rechdr on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
  ' "where emp.Active=1 and IdOffice='" + id_oficina$ + "' and IdJobtitleUSR in (16,17,28, 2) and rechdr.Active=1 " & _
  ' "and cast(rechdr.Date as Date) >= '" + fecha_de_revision$ + "' AND cast( rechdr.DATE as Date) <= '" + fecha_de_revision$ + "'"

  
  Else
  

  
  
  sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where empofc.active=1 and IdJobTitle in (3,6,16,17,18, 28,2,29,24,37) and ofc.office='" + oficina$ + "' and empjob.Active='1'"
  
  
  
 '  sSelect = "select distinct IDEmployee, Username from EmployeeInfo emp " & _
 '  "inner join ReceiptsHDR rechdr on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
 '  "where IdOffice='" + id_oficina$ + "' and IdJobtitleUSR in (16,17,28, 2) and rechdr.Active=1 " & _
 '  "and cast(rechdr.Date as Date) >= '" + fecha_de_revision$ + "' AND cast( rechdr.DATE as Date) <= '" + fecha_de_revision$ + "'"

  
  End If
  
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    
    
    


  
' load agents to Combo_agents
cbo_agentes.Clear

For t = 1 To Grid2.Rows - 1
  Grid2.Row = t
  Grid2.Col = 1
  id_agente$ = Grid2.Text
  
  Grid2.Col = 2
  
  agente$ = Grid2.Text
  
  
  
  
  
  If oficina$ = "JA - HAVEN" Then
    Select Case UCase(agente$)
    Case "HNAVARRO", "DLOPEZ", "MONEYREPORTS", "RECEIPTSCORRECTIONS", "BETZY", "OZUNIGA", "BMARQUEZ", "CPEREZG", "CCADENA"    ' "GJIMENEZ"
    
    Case Else
       cbo_agentes.AddItem agente$ + Space(20) + id_agente$
    
    End Select
    
  Else
    cbo_agentes.AddItem agente$ + Space(20) + id_agente$
  End If
  
Next t


' carga el nombre completo

       sSelect = "select firstname, lastname1 from employeeinfo where username='" + user$ + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       nombre$ = Rs(0)
       apellido$ = Rs(1)
       Rs.Close

       lblname_agent.Caption = nombre$ + Space(1) + apellido$
       lbl_iniciales_agente.Caption = user$
       
       
       valido1 = 0
       carga_manager
       Exit Sub
       
       
   
' carga manager de la oficina


 Grid2.Clear
    
  
   If chkagentes(1).Value = 0 Then
  

  
  
  sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (17, 29, 24) and ofc.office='" + oficina$ + "' and empjob.Active='1'"


  'sSelect = "select distinct IDEmployee, Username from EmployeeInfo emp " & _
  ' "inner join ReceiptsHDR rechdr on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
  ' "where emp.Active=1 and IdOffice='" + id_oficina$ + "' and IdJobtitleUSR in (17) and rechdr.Active=1 " & _
  ' "and cast(rechdr.Date as Date) >= '" + fecha_de_revision$ + "' AND cast( rechdr.DATE as Date) <= '" + fecha_de_revision$ + "'"

  
  
  Else
  
 
  
  sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where empofc.active=1 and IdJobTitle in (17, 29,24) and ofc.office='" + oficina$ + "' and empjob.Active='1'"
  
  
 ' sSelect = "select distinct IDEmployee, Username from EmployeeInfo emp " & _
 '  "inner join ReceiptsHDR rechdr on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
 '  "where IdOffice='" + id_oficina$ + "' and IdJobtitleUSR in (17) and rechdr.Active=1 " & _
 '  "and cast(rechdr.Date as Date) >= '" + fecha_de_revision$ + "' AND cast( rechdr.DATE as Date) <= '" + fecha_de_revision$ + "'"

  
  
  End If
  
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    
  cbo_managers.Clear
 
  
  
  For t = 1 To Grid2.Rows - 1
  Grid2.Row = t
  Grid2.Col = 1
  ID_manager$ = Grid2.Text
  
  Grid2.Col = 2
  manager$ = Grid2.Text
  
  
  If UCase(manager$) = "GJIMENEZ" Then
     GoTo saltado
  End If
  
  
  If oficina$ = "JA - HAVEN" Then
    Select Case UCase(manager$)
    Case "GJIMENEZ"
    
    Case "HNAVARRO", "DLOPEZ", "CCADENA", "MONEYREPORTS", "RECEIPTSCORRECTIONS", "BETZY", "BMARQUEZ"  ' "GJIMENEZ"
    
    Case Else
       cbo_managers.AddItem manager$ + Space(20) + ID_manager$
    End Select
    
  Else
    
    existe = 0
    For Y = 0 To cbo_managers.ListCount - 1
       nombre_manager$ = Left(cbo_managers.List(Y), Len(cbo_managers.List(Y)) - 10)
       If RTrim(UCase(nombre_manager$)) = RTrim(UCase(manager$)) Then
           existe = 1
           Exit For
       End If
    Next Y
    
    If existe = 0 Then
         cbo_managers.AddItem manager$ + Space(20) + ID_manager$
    End If
    
    
  End If
  
saltado:
  
  Next t
   
   
   

If cbo_managers.ListCount = 1 Then
    cbo_managers.ListIndex = 0
    btnlock2.Caption = "Unlock"
    
ElseIf cbo_managers.ListCount = 0 Then
    cbo_managers.AddItem "JOSELIN" + Space(20) + "119"
    
End If
   
   

   
   
End Sub
Public Sub Conecta_SQL()
On Error Resume Next
'  Set cn_ptos = New ADODB.Connection
 '  cn_ptos.Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
 contraseña_ini$ = "Q6XSkLMjy7BUSKdxcE"
 user_ini$ = "payroll"
 bd_ini$ = "laesystemja"
 server_ini$ = "ec2-52-8-179-170.us-west-1.compute.amazonaws.com"

 With base
   .CursorLocation = adUseClient
   ' .Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=CallCenter;Data Source=AICO2-HECTOR"
    .Open "Provider=SQLOLEDB.1;Password=" + contraseña_ini$ + ";Persist Security Info=True;User ID=" + user_ini$ + ";Initial Catalog=" + bd_ini$ + ";Data Source=" + server_ini$
   
   
 End With
End Sub
Public Function GetIPAddress() As String
On Error Resume Next
   Dim sHostName    As String * 256
   Dim lpHost    As Long
   Dim HOST      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
   End If
    
  'gethostname returns the name of the local host into
  'the buffer specified by the name parameter. The host
  'name is returned as a null-terminated string. The
  'form of the host name is dependent on the Windows
  'Sockets provider - it can be a simple host name, or
  'it can be a fully qualified domain name. However, it
  'is guaranteed that the name returned will be successfully
  'parsed by gethostbyname and WSAAsyncGetHostByName.

  'In actual application, if no local host name has been
  'configured, gethostname must succeed and return a token
  'host name that gethostbyname or WSAAsyncGetHostByName
  'can resolve.
   If gethostname(sHostName, 256) = SOCKET_ERROR Then
      GetIPAddress = ""
      MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
              " has occurred. Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   
  'gethostbyname returns a pointer to a HOSTENT structure
  '- a structure allocated by Windows Sockets. The HOSTENT
  'structure contains the results of a successful search
  'for the host specified in the name parameter.

  'The application must never attempt to modify this
  'structure or to free any of its components. Furthermore,
  'only one copy of this structure is allocated per thread,
  'so the application should copy any information it needs
  'before issuing any other Windows Sockets function calls.

  'gethostbyname function cannot resolve IP address strings
  'passed to it. Such a request is treated exactly as if an
  'unknown host name were passed. Use inet_addr to convert
  'an IP address string the string to an actual IP address,
  'then use another function, gethostbyaddr, to obtain the
  'contents of the HOSTENT structure.
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
    
   If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Windows Sockets are not responding. " & _
              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
    
  'to extract the returned IP address, we have to copy
  'the HOST structure and its members
   CopyMemory HOST, lpHost, Len(HOST)
   CopyMemory dwIPAddr, HOST.hAddrList, 4
   
  'create an array to hold the result
   ReDim tmpIPAddr(1 To HOST.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
   
  'and with the array, build the actual address,
  'appending a period between members
   For i = 1 To HOST.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
  
  'the routine adds a period to the end of the
  'string, so remove it here
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   SocketsCleanup
    
End Function
Public Function GetIPHostName() As String
On Error Resume Next
    Dim sHostName As String * 256
    
    If Not SocketsInitialize() Then
        GetIPHostName = ""
        Exit Function
    End If
    
    If gethostname(sHostName, 256) = SOCKET_ERROR Then
        GetIPHostName = ""
        MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
                " has occurred.  Unable to successfully get Host Name."
        SocketsCleanup
        Exit Function
    End If
    
    GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
    SocketsCleanup

End Function



Public Function HiByte(ByVal wParam As Integer) As Byte
  On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function




Public Function LoByte(ByVal wParam As Integer) As Byte
On Error Resume Next
  'note: VB4-32 users should declare this function As Integer
   LoByte = wParam And &HFF&

End Function
Public Sub SocketsCleanup()
On Error Resume Next
    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub
Public Function SocketsInitialize() As Boolean
On Error Resume Next

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
    
  'must be OK, so lets do it
   SocketsInitialize = True
        
End Function



















Private Sub btn0_Click()
If txtoutput.Text = "0" Then txtoutput.Text = ""
txtoutput.Text = txtoutput.Text & 0
End Sub

Private Sub btn1_Click()
If txtoutput.Text = "0" Then txtoutput.Text = ""
txtoutput.Text = txtoutput.Text & 1
End Sub

Private Sub btn2_Click()
If txtoutput.Text = "0" Then txtoutput.Text = ""
txtoutput.Text = txtoutput.Text & 2
End Sub


Private Sub btn3_Click()
If txtoutput.Text = "0" Then txtoutput.Text = ""
txtoutput.Text = txtoutput.Text & 3
End Sub


Private Sub btn4_Click()
If txtoutput.Text = "0" Then txtoutput.Text = ""
txtoutput.Text = txtoutput.Text & 4
End Sub

Private Sub btn5_Click()
If txtoutput.Text = "0" Then txtoutput.Text = ""
txtoutput.Text = txtoutput.Text & 5
End Sub


Private Sub btn6_Click()
If txtoutput.Text = "0" Then txtoutput.Text = ""
txtoutput.Text = txtoutput.Text & 6
End Sub

Private Sub btn7_Click()
If txtoutput.Text = "0" Then txtoutput.Text = ""
txtoutput.Text = txtoutput.Text & 7
End Sub


Private Sub btn8_Click()
If txtoutput.Text = "0" Then txtoutput.Text = ""
txtoutput.Text = txtoutput.Text & 8
End Sub


Private Sub btn9_Click()
If txtoutput.Text = "0" Then txtoutput.Text = ""
txtoutput.Text = txtoutput.Text & 9
End Sub


Private Sub btnadd_agent_Click()
On Error Resume Next
If txtagente.Text = "" Then Exit Sub

Dim sSelect As String
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

sSelect = "select idemployee from employeeinfo where username='" + txtagente.Text + "'"
 Rs.Open sSelect, base, adOpenUnspecified
 id_employee$ = Rs(0)
 Rs.Close

 
 
 If id_employee$ = "" Then Exit Sub


 cbo_agentes.AddItem txtagente.Text + Space(20) + id_employee$

txtagente.Text = ""

End Sub

Private Sub btnborra_Click()
On Error Resume Next


If lbldate_agente.Caption = fecha_actual$ Then

  If modificado = 1 Then
    r$ = MsgBox("You have not saved the data yet. Do you want to save them?", 4, "Attention")
    If r$ = "6" Then
      btnsave_Click
    End If
  End If

End If




r$ = MsgBox("Do you want to start a new form??", 4, "Attention")
If r$ = "7" Then
   Exit Sub
End If

lbldate_agente.Caption = Format(Now, "mm/dd/yyyy")


limpia_tarjetas
limpia_datos


'carga_registros
'carga_datos
calcula_total_LAE
modificado = 0
msg.Visible = False

End Sub

Private Sub btnborra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  btnborra.Picture = img_borra_down.Picture
End Sub


Private Sub btnborra_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnborra.Picture = img_borra_up.Picture
End Sub


Private Sub btnborracash_Click(Index As Integer)
On Error Resume Next
txtdinero(Index).Text = ""
txtdinero(Index).SetFocus
End Sub

Private Sub btnborrar_archivo2_Click()
On Error Resume Next
If archivo_selecto$ = "" Then
    Exit Sub
End If
 Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    
    
 oficina$ = LTrim(Right(UCase(RTrim(cbo_oficina.List(cbo_oficina.ListIndex))), 30))
    sSelect = "SELECT idoffice From officescatalog where office='" + oficina$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    id_office$ = Rs(0)
    Rs.Close
    


 sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idoffice='" + id_office$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreportoffice$ = Rs(0)
     Rs.Close



ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
a$ = archivo_selecto$
Kill "c:\money\" + a$

ruta_archivos$ = "\\192.168.84.215\moneyreport\O-" + id_moneyreportoffice$ + "\"
Kill ruta_archivos$ + a$

archivo_selecto$ = ""

img2.Picture = LoadPicture()
pdf2.src = "c:\"
modificado = 1






End Sub

Private Sub btnborrar_archivo2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnborrar_archivo2.Picture = img_corta_down.Picture
End Sub


Private Sub btnborrar_archivo2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnborrar_archivo2.Picture = img_corta_up.Picture

End Sub


Private Sub btnc_Click()
On Error Resume Next
result = 0
mem.Caption = ""
txtoutput.Text = "0"
txtoutput.SetFocus
End Sub

Private Sub btncalendar_Click()
On Error Resume Next

Dim sSelect As String
    
Dim Rs As ADODB.Recordset

Set Rs = New ADODB.Recordset


lae_office$ = RTrim(LTrim(Right(Form1.cbo_oficina.List(Form1.cbo_oficina.ListIndex), 25)))
 
      
  
   ' sSelect = "select idoffice from officescatalog where office='" + lae_office$ + "'"
   ' Rs.Open sSelect, base, adOpenUnspecified
   ' id_oficina$ = Rs(0)
   ' Rs.Close
     
    
    ' checa la oficina y la actualiza en la barra
    'sSelect = "select office from officescatalog where idoffice='" + id_oficina$ + "'"
    'Rs.Open sSelect, base, adOpenUnspecified
    'oficina$ = Rs(0)
    'Rs.Close
    'lbloficina_agente.Caption = oficina$
    
    'lbloficina_agente.Caption = oficina_trabajada$
    
    
    
    
    'sSelect = "select ofc.Office  from ReceiptsHDR rechdr " & _
    '"inner join OfficesCatalog ofc on ofc.IdOffice= rechdr.IdOffice " & _
    '"inner join EmployeeInfo emp on emp.IDEmployee =  rechdr.IdEmployeeUSR " & _
    '"where  cast(rechdr.Date as Date) >= '" + lbldate_agente.Caption + "' " & _
    '"AND cast( rechdr.Date as Date) <= '" + lbldate_agente.Caption + "' and emp.Username='" + user$ + "' group by ofc.Office"

    'oficina$ = Rs(0)
    'Rs.Close
    'lbloficina_agente.Caption = oficina$
    
    
    
    
    
    'If lbloficina_agente.Caption = "" Then
    '      lbloficina_agente.Caption = lae_office$
    'End If
    


cbo_oficina.Enabled = False

Calendar1.Value = lbldate_agente.Caption
Calendar1.Visible = True
Calendar1.SetFocus

End Sub

Private Sub btncalendar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btncalendar.Picture = img_calendar_down.Picture
End Sub


Private Sub btncalendar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btncalendar.Picture = img_calendar_up.Picture

End Sub


Private Sub btncarga_datos_agente_Click()
On Error Resume Next

carga_registros

 For Y = 1 To grid1.Rows - 1
    
    grid1.Row = Y
    
    grid1.Col = 18
    f$ = Format(grid1.Text, "mm/dd/yyyy")
    
    If Right(f$, 4) = "1900" Then
       f$ = ""
    End If
    
    grid1.Text = f$
    
    Next Y
    
    
    
    
    For Y = 1 To grid3.Rows - 1
    
    grid3.Row = Y
    
    grid3.Col = 18
    f$ = Format(grid3.Text, "mm/dd/yyyy")
    
    If Right(f$, 4) = "1900" Then
       f$ = ""
    End If
    
    grid3.Text = f$
    
    Next Y
    

End Sub

Private Sub btncarga_datos_agente_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btncarga_datos_agente.Picture = img_carga_datos_down.Picture
End Sub


Private Sub btncarga_datos_agente_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btncarga_datos_agente.Picture = img_carga_datos_up.Picture
End Sub


Private Sub btncargar_reportes_Click()
On Error Resume Next
 id_employee = Val(Right(Form1.cbo_agentes.List(Form1.cbo_agentes.ListIndex), 20))
 ID_manager = Val(Right(Form1.cbo_managers.List(Form1.cbo_managers.ListIndex), 20))
  
If id_employee = 0 Then
   MsgBox "Select the Agent from the list", 16, "Attention"
   Exit Sub
End If


If ID_manager = 0 Then
   MsgBox "Select the Manager from the list", 16, "Attention"
   Exit Sub
End If

transfiere$ = ""

Load forma_reportes_pendientes
forma_reportes_pendientes.Show 1

fech$ = Format(transfiere$, "mm/dd/yyyy")
If Val(Right(fech$, 4)) < 2000 Then
   valido1 = 0
   Exit Sub
End If

If transfiere$ <> "" Then
  lbldate_agente.Caption = Format(transfiere$, "mm/dd/yyyy")
  Calendar1.Value = lbldate_agente.Caption
  Calendar1_Click
End If

valido1 = 0

End Sub

Private Sub btncierra_visualizador2_Click()
On Error Resume Next
visualizador2.Visible = False
End Sub

Private Sub btnclear_cust_Click(Index As Integer)
On Error Resume Next
txtcustomer_agente(Index).Text = ""
txtcustomer_agente(Index).SetFocus
End Sub

Private Sub btnclear_void_Click(Index As Integer)
On Error Resume Next
txtamount_agente(Index).Text = ""
txtamount_agente(Index).SetFocus

End Sub

Private Sub btnclose_viewer_Click()
On Error Resume Next
visualizador.Visible = False
End Sub

Private Sub btncopy_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText txtoutput.Text


End Sub

Private Sub btndesbloquear_Click()
On Error Resume Next

Dim sSelect As String
    
Dim Rs As ADODB.Recordset

Set Rs = New ADODB.Recordset


  oficina$ = LTrim(Right(UCase(RTrim(cbo_oficina.List(cbo_oficina.ListIndex))), 30))
    sSelect = "SELECT idoffice From officescatalog where office='" + oficina$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_office$ = Rs(0)
    Rs.Close
    
    
    
    UserName$ = RTrim(Left(cbo_agentes.List(cbo_agentes.ListIndex), Len(cbo_agentes.List(cbo_agentes.ListIndex)) - 5))
    sSelect = "SELECT idemployee From employeeinfo where username='" + UserName$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_employee$ = Rs(0)
    Rs.Close


If tipo_guardado = 1 Then
 
     sSelect = "select idmoneyreport from moneyreport where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "' and idoffice='" + id_office$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreport$ = Rs(0)
     Rs.Close
    


' verifica si ya existe el reporte

      sSelect = "update MoneyReport set submitted='0'" & _
      "where idmoneyreport='" + id_moneyreport$ + "'"  ' datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "' and idoffice='" + id_office$ + "'"

                      
      Rs.Open sSelect, base, adOpenUnspecified
    
      Rs.Close
      
      
Else
      
    
    UserName$ = RTrim(Left(cbo_managers.List(cbo_managers.ListIndex), Len(cbo_managers.List(cbo_managers.ListIndex)) - 5))
    sSelect = "SELECT idemployee From employeeinfo where username='" + UserName$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_employee$ = Rs(0)
    Rs.Close
    
    
    
    sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idmanager='" + id_employee$ + "' and idoffice='" + id_office$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreportoffice$ = Rs(0)
     Rs.Close
     
    
    
    
      sSelect = "update MoneyReportbyoffice set submitted='0' " & _
      "where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"  ' datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "' and idoffice='" + id_office$ + "'"

                      
      Rs.Open sSelect, base, adOpenUnspecified
    
      Rs.Close
    
End If

         btnerase_archivo.Enabled = True
         btnsave.Enabled = True
         chk_dayoff.Enabled = True
         btnsave.Picture = img_disk_up.Picture

         btnsend.Enabled = True
         btnsend.Picture = img_send_up.Picture
    
End Sub

Private Sub btndiv_Click()
On Error Resume Next
operator = 4

sd = txtoutput.Text
If mem.Caption = "" Then
  mem.Caption = Format(sd, "###,##0.00")
Else
  mem.Caption = Format(Val(Format(mem.Caption, "00000.00")) / sd, "###,##0.00")
End If
signo.Caption = "/"
txtoutput.Text = ""
txtoutput.SetFocus

End Sub

Private Sub btnend_Click()
On Error Resume Next

If submitido = "false" Then
 If lbldate_agente.Caption = fecha_actual$ Then

  If modificado = 1 Then
    r$ = MsgBox("You have not saved the data yet. Do you want to save them?", 4, "Attention")
    If r$ = "6" Then
    
      btnsave_Click
    End If
  End If

 End If
End If

base.Close


  

  r$ = Shell("c:\money\reset.exe", vbNormalFocus)

  r$ = Shell("c:\iconos\barra_agent.exe", vbNormalFocus)



'r$ = Shell("c:\money\cierra_money.exe")
X$ = Shell("cmd /c taskkill /f /im money.exe")


End

End Sub













Private Sub btnend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnend.Picture = img_end_down.Picture
End Sub

Private Sub btnend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnend.Picture = img_end_up.Picture

End Sub

Private Sub btnerase_archivo_Click()
On Error Resume Next
If archivo_selecto$ = "" Then
    Exit Sub
End If

ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
a$ = archivo_selecto$
Kill "c:\money\" + a$

ruta_archivos$ = "\\192.168.84.215\moneyreport\" + id_moneyreport$ + "\"
Kill ruta_archivos$ + a$

archivo_selecto$ = ""

img1.Picture = LoadPicture()
pdf1.src = "c:\"
modificado = 1
End Sub

Private Sub btnerase_archivo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnerase_archivo.Picture = img_corta_down.Picture
End Sub

Private Sub btnerase_archivo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnerase_archivo.Picture = img_corta_up.Picture
End Sub


Private Sub btnerase_combo_Click(Index As Integer)
On Error Resume Next
cbo_employees(Index).ListIndex = -1
cbooficina1(Index).ListIndex = -1


Select Case Index
Case 0
  txtcant_ida(0).Text = ""
Case 1
    txtcant_ida(1).Text = ""
Case 2
  txtcant_ida(2).Text = ""
Case 3
  txtcant_venida(0).Text = ""
Case 4
  txtcant_venida(1).Text = ""
Case 5
  txtcant_venida(2).Text = ""
End Select
  
  
End Sub

Private Sub btnerasecredit_Click(Index As Integer)
On Error Resume Next
If Index = 0 Then
   txtdebit_manager.Text = ""
   txtdebit_manager.SetFocus
Else
   txtcredit_manager.Text = ""
   txtcredit_manager.SetFocus
End If
End Sub

Private Sub btneraser_recibo_Click(Index As Integer)
On Error Resume Next
txtrecibos_agente(Index).Text = ""
txtrecibos_agente(Index).SetFocus
End Sub

Private Sub btnfix_Click()
On Error Resume Next
MkDir "c:\money\backup"
FileCopy "c:\money\*.pdf", "c:\money\backup\*.pdf"
FileCopy "c:\money\*.jpg", "c:\money\backup\*.jpg"
FileCopy "c:\money\*.bmp", "c:\money\backup\*.bmp"


Kill "c:\money\*.pdf"
Kill "c:\money\*.jpg"
Kill "c:\money\*.bmp"

ListView1.ListItems.Clear
ListView2.ListItems.Clear



MsgBox "Done. Try it again!", 64, "Fixed up!"



End Sub

Private Sub btnigual_Click()
On Error Resume Next
sd = mem.Caption

If operator = 1 Then

result = Val(Format(sd, "00000.00")) + Val(txtoutput.Text)
txtoutput.Text = result

ElseIf operator = 2 Then

result = Val(Format(sd, "00000.00")) - Val(txtoutput.Text)
txtoutput.Text = result

ElseIf operator = 3 Then

result = Val(Format(sd, "00000.00")) * Val(txtoutput.Text)
txtoutput.Text = result

ElseIf operator = 4 Then

result = Val(Format(sd, "00000.00")) / Val(txtoutput.Text)
txtoutput.Text = result

ElseIf operator = 5 Then

result = Val(Format(sd, "00000.00")) Mod Val(txtoutput.Text)
txtoutput.Text = result

ElseIf operator = 6 Then

result = Val(Format(sd, "00000.00")) * Val(sd)
txtoutput.Text = result

End If
signo.Caption = ""
mem.Caption = ""
txtoutput.SetFocus

End Sub

Private Sub btnlimpiacash_Click(Index As Integer)
On Error Resume Next
txtcash(Index).Text = ""
txtcash(Index).SetFocus
End Sub

Private Sub btnlimpiacredit_Click(Index As Integer)
On Error Resume Next
txtcredit_agente(Index).Text = ""
txtcredit_agente(Index).SetFocus
End Sub

Private Sub btnlimpiadebit_Click(Index As Integer)
On Error Resume Next
txtdebit_agente(Index).Text = ""
txtdebit_agente(Index).SetFocus
End Sub

Private Sub btnlock_Click()
On Error Resume Next


If cbo_oficina.ListCount > 1 And cbo_oficina.Enabled = True Then
   cbo_oficina.Enabled = False
   btnlock.Caption = "Unlock"
   Exit Sub
End If


If cbo_oficina.ListCount > 1 And cbo_oficina.Enabled = False Then
   cbo_oficina.Enabled = True
   btnlock.Caption = "Lock"
   Exit Sub
End If

  
  

  oficina$ = LTrim(RTrim(Right(cbo_oficina.List(cbo_oficina.ListIndex), 30)))
  
  lbloficina_agente.Caption = oficina$
  
  lbloficina_agente.Caption = oficina_trabajada$
  

 
  

End Sub

Private Sub btnlock2_Click()
On Error Resume Next
If cbo_managers.ListCount >= 1 And cbo_managers.Enabled = True Then
   cbo_managers.Enabled = False
   btnlock2.Caption = "Unlock"
   Exit Sub
End If


If cbo_managers.ListCount >= 1 And cbo_managers.Enabled = False Then
   cbo_managers.Enabled = True
   btnlock2.Caption = "Lock"
   Exit Sub
End If


  'oficina$ = LTrim(RTrim(Right(cbo_managers.List(cbo_managers.ListIndex), 30)))
  



End Sub


Private Sub btnmas_Click()
On Error Resume Next
operator = 1

sd = txtoutput.Text
mem.Caption = Format(Val(Format(mem.Caption, "00000.00")) + sd, "###,##0.00")
signo.Caption = "+"

txtoutput.Text = ""
txtoutput.SetFocus

End Sub

Private Sub btnminus_Click()
On Error Resume Next
operator = 2

sd = txtoutput.Text
If mem.Caption = "" Then
   mem.Caption = Format(sd, "###,##0.00")
Else
   mem.Caption = Format(Val(Format(mem.Caption, "00000.00")) - sd, "###,##0.00")
End If
signo.Caption = "-"
txtoutput.Text = ""
txtoutput.SetFocus

End Sub



Private Sub btnmul_Click()
On Error Resume Next
operator = 3

sd = txtoutput.Text
If Val(mem.Caption) = 0 Then mem.Caption = "1"
mem.Caption = Format(Val(Format(mem.Caption, "00000.00")) * sd, "###,##0.00")
signo.Caption = "X"
txtoutput.Text = ""
txtoutput.SetFocus

End Sub



Private Sub btnok_Click()
On Error Resume Next

    Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    
    
    oficina$ = LTrim(Right(UCase(RTrim(cbo_oficina.List(cbo_oficina.ListIndex))), 30))
    sSelect = "SELECT idoffice From officescatalog where office='" + oficina$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    id_office$ = Rs(0)
    Rs.Close
    
    
    
    
    UserName$ = RTrim(Left(cbo_agentes.List(cbo_agentes.ListIndex), Len(cbo_agentes.List(cbo_agentes.ListIndex)) - 5))
    sSelect = "SELECT idemployee From employeeinfo where username='" + UserName$ + "'"  ' and active='1'"
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    id_employee$ = Rs(0)
    Rs.Close
    
    
    
    UserName2$ = RTrim(Left(cbo_managers.List(cbo_managers.ListIndex), Len(cbo_managers.List(cbo_managers.ListIndex)) - 5))
    sSelect = "SELECT idemployee From employeeinfo where username='" + UserName2$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    ID_manager$ = Rs(0)
    Rs.Close
    
    
    If ID_manager$ = "0" Then
       MsgBox "Select the manager", 16, "Attention"
       Exit Sub
    End If
    
    
    If id_employee$ = "0" Then
       MsgBox "Select the Agent", 16, "Attention"
       Exit Sub
    End If
    
    
    Dim valor_revisado As Boolean
    
    If tipo_guardado = 1 Then
     
    
         
        ' verifica si ya existe el reporte
     
        sSelect = "select idmoneyreport from moneyreport where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "' and idoffice='" + id_office$ + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        id_moneyreport$ = Rs(0)
        Rs.Close
     
        If id_moneyreport$ = "" Then Exit Sub
        
        
         If chk_revisado.Value = 1 Then
            valor_revisado2$ = "1"
         Else
            valor_revisado2$ = "0"
         End If
                   
         sSelect = "update moneyreport set reviewed='" + valor_revisado2$ + "' where idmoneyreport='" + id_moneyreport$ + "'"
                 
         Rs.Open sSelect, base, adOpenUnspecified
         Rs.Close
   
     
     
     
     Else
     
        sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idmanager='" + ID_manager$ + "' and idoffice='" + id_office$ + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        id_moneyreportoffice$ = Rs(0)
        Rs.Close
     
        If id_moneyreportoffice$ = "" Then Exit Sub
     
     
        If chk_revisado.Value = 1 Then
            valor_revisado2$ = "1"
         Else
            valor_revisado2$ = "0"
         End If
                   
        
         
                   
         sSelect = "update moneyreportbyoffice set reviewbyaccounting='" + valor_revisado2$ + "' where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"
                 
         Rs.Open sSelect, base, adOpenUnspecified
         Rs.Close
   
     
     
     End If
     
     
     MsgBox "It was saved", 64, "Attention"
     
    
End Sub

Private Sub btnok_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnok.Picture = img_palomita_down.Picture

End Sub

Private Sub btnok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnok.Picture = img_palomita_up.Picture
End Sub


Private Sub btnopen1_Click()
On Error Resume Next

n$ = ""
cd1.FileName = ""
cd1.DialogTitle = "Open File"
    cd1.InitDir = "c:\money"
    cd1.Filter = "ALL Files (*.pdf,*.jpg,*.bmp)|*.pdf;*.jpg;*.bmp|PDF Files (*.pdf)|*.pdf|JPG Files (*.JPG)|*.JPG|Bitmap Files (*.BMP)|*.BMP"
    cd1.FilterIndex = 1
    cd1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    cd1.CancelError = True

  cd1.ShowOpen
  n$ = cd1.FileName
  
  If n$ = "" Then
    Exit Sub
  End If
  
  arrastra_archivo (n$)
  modificado = 1
End Sub

Private Sub btnopen1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnopen1.Picture = img_load_down.Picture
End Sub


Private Sub btnopen1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnopen1.Picture = img_load_up.Picture

End Sub


Private Sub btnopen2_Click()
On Error Resume Next

n$ = ""
cd1.FileName = ""
cd1.DialogTitle = "Open File"
    cd1.InitDir = "c:\money"
    cd1.Filter = "ALL Files (*.pdf,*.jpg,*.bmp)|*.pdf;*.jpg;*.bmp|PDF Files (*.pdf)|*.pdf|JPG Files (*.JPG)|*.JPG|Bitmap Files (*.BMP)|*.BMP"
    cd1.FilterIndex = 1
    cd1.flags = _
        cdlOFNFileMustExist + _
        cdlOFNHideReadOnly + _
        cdlOFNLongNames + _
        cdlOFNExplorer
    cd1.CancelError = True

  cd1.ShowOpen
  n$ = cd1.FileName
  
  If n$ = "" Then
    Exit Sub
  End If
  
  arrastra_archivo2 (n$)
  modificado = 1
  
  
End Sub

Private Sub btnopen2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnopen2.Picture = img_load_down.Picture
End Sub


Private Sub btnopen2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnopen2.Picture = img_load_up.Picture

End Sub


Private Sub btnprinter_Click()
On Error Resume Next




r$ = MsgBox("Do you wish to print it?", 4, "Attention")
If r$ = "7" Then Exit Sub


Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
    
    
    

msg.Visible = True
msg.Refresh

Printer.FontName = "Courier new"
Printer.FontSize = 10

linea = 0

' imprime encabezado

d$ = Format(Now, "MM/DD/YYYY hh:mm am/pm")
pagina = 1

Printer.Orientation = vbPRORPortrait

If cargo_impresion = 0 Then
  ' imprime reporte del agente
   Printer.FontName = "Courier new"
   Printer.FontSize = 16
   Printer.Print Space(1)
   Printer.Print Space(25) + "AGENT REPORT"
   Printer.Print Space(1)
   Printer.FontSize = 10
   
   Printer.Print Space(5) + "Full Name: " + lblname_agent.Caption + Space(40 - Len(lblname_agent.Caption));
   Printer.Print Space(5) + "Agent Initials: " + lbl_iniciales_agente.Caption
   Printer.Print Space(5) + "Date of Report: " + lbldate_agente.Caption
   Printer.Print Space(1)
   Printer.Print Space(5) + "Location: " + lbloficina_agente.Caption
   Printer.Print Space(3) + "----------------------------------------------------------------------------------------------"
   Printer.Print Space(1)
   
   Printer.FontSize = 16
   Printer.Print Space(3) + "CASH" + Space(26) + "DEBIT & CREDIT CARDS"
      
   Printer.FontSize = 10
   
   Dim credito1(10), debito1(10)
   
   
   For z = 0 To 19
   
   If Val(txtdebit_agente(z).Text) > 0 Then
     debito1(z) = Format(Format(txtdebit_agente(z).Text, "$###,##0.00"), "@@@@@@@@@@@")
   Else
     debito1(z) = Format("---", "@@@@@@@@@@@")
   End If
   
   If Val(txtcredit_agente(z).Text) > 0 Then
     credito1(z) = Format(Format(txtcredit_agente(z).Text, "$###,##0.00"), "@@@@@@@@@@@")
   Else
     credito1(z) = Format("---", "@@@@@@@@@@@")
   End If
   
   
   
   Next z
   
   
   
   Printer.Print Space(5) + "Total Cash Deposited: " + Format(Format(txtcash(0).Text, "$###,##0.00"), "@@@@@@@@@@@") + Space(15) + "01" + Space(1) + debito1(0) + Space(5) + credito1(0) 'Format(Format(txtcredit_agente(0).Text, "$###,##0.00"), "@@@@@@@@@@@")
   
   total_money_order = Val(txtcash(1).Text) + Val(txtcash(2).Text)
   Printer.Print Space(5) + "Money Order:          " + Format(Format(total_money_order, "$###,##0.00"), "@@@@@@@@@@@") + Space(15) + "02" + Space(1) + debito1(1) + Space(5) + credito1(1)
   
   total_checks = Val(txtcash(3).Text) + Val(txtcash(4).Text)
   Printer.Print Space(5) + "Checks:               " + Format(Format(total_checks, "$###,##0.00"), "@@@@@@@@@@@") + Space(15) + "03" + Space(1) + debito1(2) + Space(5) + credito1(2)
   Printer.Print Space(5) + "Coins:                " + Format(Format(txtcash(5).Text, "$###,##0.00"), "@@@@@@@@@@@") + Space(15) + "04" + Space(1) + debito1(3) + Space(5) + credito1(3)
   Printer.Print Space(53) + "05" + Space(1) + debito1(4) + Space(5) + credito1(4)
   Printer.Print Space(53) + "06" + Space(1) + debito1(5) + Space(5) + credito1(5)
   Printer.Print Space(53) + "07" + Space(1) + debito1(6) + Space(5) + credito1(6)
   Printer.Print Space(53) + "08" + Space(1) + debito1(7) + Space(5) + credito1(7)
   Printer.Print Space(53) + "09" + Space(1) + debito1(8) + Space(5) + credito1(8)
   Printer.Print Space(53) + "10" + Space(1) + debito1(9) + Space(5) + credito1(9)
   Printer.Print Space(53) + "11" + Space(1) + debito1(10) + Space(5) + credito1(10)
   Printer.Print Space(53) + "12" + Space(1) + debito1(10) + Space(5) + credito1(11)
   Printer.Print Space(53) + "13" + Space(1) + debito1(10) + Space(5) + credito1(12)
   Printer.Print Space(53) + "14" + Space(1) + debito1(10) + Space(5) + credito1(13)
   Printer.Print Space(53) + "15" + Space(1) + debito1(10) + Space(5) + credito1(14)
   Printer.Print Space(53) + "16" + Space(1) + debito1(10) + Space(5) + credito1(15)
   Printer.Print Space(53) + "17" + Space(1) + debito1(10) + Space(5) + credito1(16)
   Printer.Print Space(53) + "18" + Space(1) + debito1(10) + Space(5) + credito1(17)
   Printer.Print Space(53) + "19" + Space(1) + debito1(10) + Space(5) + credito1(18)
   Printer.Print Space(53) + "20" + Space(1) + debito1(10) + Space(5) + credito1(19)
   
   
   
   
   Printer.Print Space(1)
   Printer.Print Space(56) + Format(Format(lbltotal_debit_agent.Caption, "$###,##0.00"), "@@@@@@@@@@@") + Space(5) + Format(Format(lbltotal_credit_agent.Caption, "$###,##0.00"), "@@@@@@@@@@@")
   
   
   Printer.FontSize = 12
   Printer.Print Space(1)
   Printer.FontBold = True
   Printer.Print Space(4) + "TOTAL: " + Format(lbltotal_cash_agente.Caption, "$###,##0.00") + Space(26) + "TOTAL: " + Format(lbltotal_debit_credit_agent.Caption, "$###,##0.00")
   Printer.FontBold = False
   Printer.Print Space(1)
   
   
   
   
   Printer.FontSize = 16
   Printer.Print Space(3) + "Receipts/Pending VOIDS"
   
   Printer.FontSize = 10
   Printer.Print Space(5) + "Cust ID" + Space(8) + "Receipt #" + Space(7) + "Amount"
   For Y = 0 To 1
      Printer.Print Space(5) + Format(txtcustomer_agente(Y).Text, "@@@@@@") + Space(6) + Format(txtrecibos_agente(Y).Text, "@@@@@@@@@@") + Space(5) + Format(Format(txtamount_agente(Y).Text, "$###,##0.00"), "@@@@@@@@@@@")
   Next Y
   
   
   Printer.Print Space(1)
   Printer.Print Space(1)
   
   
   Printer.FontSize = 14
   
   Printer.Print Space(4) + "Total Reported:           " + Format(lbltotal_reported_agent.Caption, "@@@@@@@@@@@")
   Printer.Print Space(4) + "Pending Void Receipts:    " + Format(lbltotal_void_agente.Caption, "@@@@@@@@@@@")
   Printer.FontBold = True
   Printer.Print Space(4) + "Totals:                   " + Format(lbltotal_reportado_menos_voids_agente.Caption, "@@@@@@@@@@@")
   Printer.FontBold = False
   Printer.Print Space(1)
   Printer.Print Space(4) + "Debit and Credit Sales:   " + Format(lblgrantotal_debitcredit_agente.Caption, "@@@@@@@@@@@")
   Printer.Print Space(4) + "Total Cash:               " + Format(lblgrantotal_cash_agente.Caption, "@@@@@@@@@@@")
   
   If Val(Format(lbltotal_over_short_agente.Caption, "000000.00")) >= 0 Then
      Printer.Print Space(4) + "Over (Short):             " + Format(lbltotal_over_short_agente.Caption, "@@@@@@@@@@@")
   Else
      Printer.Print Space(4) + "Over (Short):            (" + Format(lbltotal_over_short_agente.Caption, "@@@@@@@@@@@") + ")"
   End If
   
   
   
   
   
   Printer.FontSize = 12
   Printer.FontBold = True
   Printer.Print
   Printer.FontBold = False
   Printer.FontSize = 10
   
   Printer.Print Space(1)
   Printer.Print Space(5) + "NOTE: "
   conta = 0
   r$ = ""
   For Y = 1 To Len(txtnotas_agente.Text)
        conta = conta + 1
        a$ = Mid(txtnotas_agente.Text, Y, 1)
        
        If a$ = Chr$(10) Then
          GoTo saltado
        End If
        
        If a$ = Chr$(13) Then
          r$ = r$ + Space(1)
          GoTo saltado
        End If
        
        
        If conta < 80 Then
           r$ = r$ + a$
        ElseIf conta >= 80 And a$ = " " Then
           Printer.Print Space(5) + r$
           r$ = ""
           conta = 0
        Else
          
        End If
        
saltado:
        
    Next Y
           
    If conta > 0 Then
         Printer.Print Space(5) + r$
    End If
   
   
    Printer.Print Space(1)
    Printer.Print Space(1)
   
    Printer.Print Space(1)
   Printer.Print Space(1)
   Printer.Print Space(1)
   Printer.Print Space(1)
   Printer.Print Space(5) + "_____________________________                    ______________________________"
   
   
   
   sSelect = "select firstname, lastname1 from employeeinfo where username='" + LTrim(RTrim(Left(cbo_managers.List(cbo_managers.ListIndex), Len(cbo_managers.List(cbo_managers.ListIndex)) - 15))) + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    nombre$ = Rs(0)
    apellido$ = Rs(1)
    Rs.Close
           
    full_name_manager$ = nombre$ + " " + apellido$
    
    
    
    sSelect = "select firstname, lastname1 from employeeinfo where username='" + LTrim(RTrim(Left(cbo_agentes.List(cbo_agentes.ListIndex), Len(cbo_agentes.List(cbo_agentes.ListIndex)) - 15))) + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    nombre$ = Rs(0)
    apellido$ = Rs(1)
    Rs.Close
           
    full_name_agente$ = nombre$ + " " + apellido$
    
   
   
   Printer.Print Space(12) + full_name_agente$ + Space(34) + full_name_manager$
   
    
    
    
  
   
   
  
ElseIf cargo_impresion = 1 Then
  ' imprime reporte del supervisor ---   cargo_impresion=1
  ' ****************************************************************************************************************************************************
  
   sSelect = "select firstname, lastname1 from employeeinfo where username='" + LTrim(RTrim(Left(cbo_managers.List(cbo_managers.ListIndex), Len(cbo_managers.List(cbo_managers.ListIndex)) - 15))) + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    nombre$ = Rs(0)
    apellido$ = Rs(1)
    Rs.Close
           
    full_name$ = nombre$ + " " + apellido$
           
           
  ' imprime reporte del manager
   Printer.FontName = "Courier new"
   Printer.FontSize = 16
   Printer.Print Space(1)
   Printer.Print Space(24) + "MANAGER REPORT"
   Printer.Print Space(1)
   Printer.FontSize = 10
   
   Printer.Print Space(5) + "Full Name: " + full_name$
   Printer.Print Space(5) + "Manager Initials: " + lbl_iniciales_agente.Caption
   Printer.Print Space(5) + "Date of Report: " + lbldate_agente.Caption
   Printer.Print Space(1)
   Printer.Print Space(5) + "Location: " + lbloficina_agente.Caption
   Printer.Print Space(3) + "----------------------------------------------------------------------------------------------"
   Printer.Print Space(1)
   
   Printer.FontSize = 16
   Printer.Print Space(3) + "CASH" + Space(26) + "DEBIT & CREDIT CARDS"
      
   Printer.FontSize = 10
   
   
   
   
   If Val(txtdebit_manager.Text) > 0 Then
     debito1x$ = Format(Format(txtdebit_manager.Text, "$###,##0.00"), "@@@@@@@@@@@")
   Else
     debito1x$ = Format("---", "@@@@@@@@@@@")
   End If
   
   For Y = 0 To 2
     If txtdinero(Y).Text = "" Then
         txtdinero(Y).Text = "0"
     End If
   Next Y
     
   '
   Printer.Print Space(5) + "Cash deposited  :     " + Format(Format(txtdinero(0).Text, "$###,##0.00"), "@@@@@@@@@@@") + Space(15) + "Debit:   " + Space(1) + debito1x$
   
   Printer.Print Space(5) + "Coins:                " + Format(Format(txtdinero(1).Text, "$###,##0.00"), "@@@@@@@@@@@") + Space(15) + "Credit:  " + Space(1) + Format(Format(txtcredit_manager.Text, "$###,##0.00"), "@@@@@@@@@@@")
   
   Printer.Print Space(5) + "Money order:          " + Format(Format(txtdinero(2).Text, "$###,##0.00"), "@@@@@@@@@@@")
   
   
   
   
  
   Printer.FontSize = 12
   Printer.Print Space(1)
   Printer.FontBold = True
   Printer.Print Space(4) + "TOTAL: " + Format(lbltotal_cash_manager.Caption, "$###,##0.00") + Space(24) + "TOTAL: " + Format(lbltotal_debito_credito_manager.Caption, "$###,##0.00")
   Printer.FontBold = False
   Printer.Print Space(1)
   Printer.Print Space(1)
    Printer.Print Space(1)
    Printer.Print Space(1)
   
   
   Printer.FontSize = 16
   Printer.Print Space(3) + "My Agent went to:" + Space(13) + "Agent came to help us:"
   
   Printer.FontSize = 10
   Printer.Print Space(5) + "Agent  " + Space(9) + "Office   " + Space(7) + "Amount" + Space(10) + "Agent  " + Space(9) + "Office   " + Space(7) + "Amount"
   ' Printer.Print Space(4) + "-------------------------------------------" + Space(5) + "--------------------------------------------"
   Printer.Print Space(4) + "¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯" + Space(5) + "¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯"
   For Y = 0 To 2
      nombre_agente$ = LTrim(RTrim(Left(cbo_employees(Y).List(cbo_employees(Y).ListIndex), 15)))
      nombre_oficina$ = LTrim(RTrim(Left(cbooficina1(Y).List(cbooficina1(Y).ListIndex), 15)))
      
      If nombre_agente$ = "" Then
         nombre_agente$ = "---"
      End If
      
      If nombre_oficina$ = "" Then
         nombre_oficina$ = "---"
      End If
      
      nombre_agente2$ = LTrim(RTrim(Left(cbo_employees(Y + 3).List(cbo_employees(Y + 3).ListIndex), 15)))
      nombre_oficina2$ = LTrim(RTrim(Left(cbooficina1(Y + 3).List(cbooficina1(Y + 3).ListIndex), 15)))
      
      If nombre_agente2$ = "" Then
         nombre_agente2$ = "---"
      End If
      
      If nombre_oficina2$ = "" Then
         nombre_oficina2$ = "---"
      End If
      
      
      If Val(txtcant_ida(Y).Text) > 0 Then
         Printer.Print Space(5) + Format(nombre_agente$, "!@@@@@@@@@@@@@@@") + Space(1) + Format(nombre_oficina$, "!@@@@@@@@@@@@@@@") + Format(Format(txtcant_ida(Y).Text, "$#,##0.00"), "@@@@@@@@@");
      Else
         Printer.Print Space(45);
      End If
      
      
      If Val(txtcant_venida(Y).Text) > 0 Then
         Printer.Print Space(8) + Format(nombre_agente2$, "!@@@@@@@@@@@@@@@") + Space(1) + Format(nombre_oficina2$, "!@@@@@@@@@@@@@@@") + Format(Format(txtcant_venida(Y).Text, "$#,##0.00"), "@@@@@@@@@")
      Else
         Printer.Print Space(45)
      End If
      
      
   Next Y
   
   
   Printer.Print Space(1)
  
   Printer.FontSize = 12
   ' Printer.FontBold = True
   Printer.Print Space(4) + "TOTAL: " + Format(lbltotal_agentes_idos.Caption, "$###,##0.00") + Space(26) + "TOTAL: " + Format(lbltotal_agentes_que_vinieron.Caption, "$###,##0.00")
   'Printer.FontBold = False
   
   
   
   
   Printer.Print Space(1)
   Printer.Print Space(1)
   Printer.Print Space(1)
   Printer.Print Space(1)
   Printer.Print Space(1)
   
   
   Printer.FontSize = 14
   
   Printer.Print Space(4) + "Total LAE System:         " + Format(lbltotal_LAE_oficina.Caption, "@@@@@@@@@@@")
   Printer.Print Space(4) + "Debit & Credit:           " + Format(lbltotal_debit_and_credit_oficina.Caption, "@@@@@@@@@@@")
   Printer.Print Space(4) + "Cash:                     " + Format(lbltotal_needed_oficina.Caption, "@@@@@@@@@@@")
   Printer.Print Space(1)
   Printer.Print Space(4) + "My agent left:            " + Format(lbltotal_dejado_por_agentes_de_oficina.Caption, "@@@@@@@@@@@")
   Printer.Print Space(4) + "Agent Came:               " + Format(lbltotal_dejado_por_agentes_que_vinieron.Caption, "@@@@@@@@@@@")
   Printer.Print Space(1)
   
   over_short = Val(Format(lblover_short_oficina.Caption, "0000000.00"))
   If over_short >= 0 Then
   
       Printer.Print Space(4) + "Over (Short):             " + Format(lblover_short_oficina.Caption, "@@@@@@@@@@@")
   ElseIf over_short < 0 Then
       Printer.Print Space(4) + "Over (Short):            (" + Format(lblover_short_oficina.Caption, "@@@@@@@@@@@") + ")"
   
   End If
   
   
   
   Printer.Print Space(1)
   Printer.FontBold = True
   
   Printer.FontBold = False
   
   
   
   
   
   Printer.FontSize = 12
   'Printer.FontBold = True
   Printer.Print
   Printer.FontItalic = True
   
   
   Printer.Print Space(1)
   Printer.Print Space(1)
   Printer.Print Space(7) + "N O T E : "
    Printer.FontItalic = False
   Printer.FontSize = 10
   Printer.FontBold = False
   
   conta = 0
   r$ = ""
   For Y = 1 To Len(txtnotas_manager.Text)
        conta = conta + 1
        a$ = Mid(txtnotas_manager.Text, Y, 1)
        
        If a$ = Chr$(10) Then
          GoTo saltado2
        End If
        
        If a$ = Chr$(13) Then
          r$ = r$ + Space(1)
          GoTo saltado2
        End If
        
        
        If conta < 80 Then
           r$ = r$ + a$
        ElseIf conta >= 80 And a$ = " " Then
           Printer.Print Space(8) + r$
           r$ = ""
           conta = 0
        Else
          
        End If
        
saltado2:
        
    Next Y
           
    If conta > 0 Then
         Printer.Print Space(8) + r$
    End If
   
   
    Printer.Print Space(1)
    Printer.Print Space(1)
    Printer.Print Space(1)
   
    Printer.Print Space(1)
   Printer.Print Space(1)
   Printer.Print Space(1)
   Printer.Print Space(1)
   Printer.Print Space(5) + "_____________________________                    ______________________________"
   
   'etiqueta_agente$ = Format(Left(cbo_agentes.List(cbo_agentes.ListIndex), Len(cbo_agentes.List(cbo_agentes.ListIndex)) - 10), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
   'etiqueta_manager$ = Format(Left(cbo_managers.List(cbo_managers.ListIndex), Len(cbo_managers.List(cbo_managers.ListIndex)) - 10), "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@")
   
   Printer.Print Space(12) + full_name$ + Space(34) + "Verified by:"
   
  
  
ElseIf cargo_impresion = 3 Then

Printer.Orientation = vbPRORLandscape
Printer.FontName = "courier new"
Printer.Print " "
Printer.Print " "
Printer.FontSize = 12
Printer.Print Space(5) + "RECEIPTS:"
Printer.Print " "
Printer.FontSize = 5



For t = 0 To grid1.Rows
  grid1.Row = t
  lineac$ = ""
       If t = 0 Then
       
       grid1.Col = 0  ' num
       lineac$ = lineac$ + Format("Row#", "@@@") + Space$(1)
       
       Else
       grid1.Col = 0  ' num
       lineac$ = lineac$ + Format(grid1.Text, "@@@") + Space$(1)
  
       End If
       
  
       grid1.Col = 1  ' Receipt#
       lineac$ = lineac$ + Format(grid1.Text, "@@@@@@@@@") + Space$(1)
       
       
       grid1.Col = 2  ' Date
       a$ = Left(grid1.Text, 10)
       lineac$ = lineac$ + Format(a$, "!@@@@@@@@@@") + Space$(1)
  
       grid1.Col = 3  ' IDcust
       lineac$ = lineac$ + Format(grid1.Text, "!@@@@@@") + Space$(1)
       
       grid1.Col = 4  ' Customer name
       lineac$ = lineac$ + Format(grid1.Text, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 5  ' Policy#
       lineac$ = lineac$ + Format(grid1.Text, "!@@@@@@@@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 6  ' IDCompany
       lineac$ = lineac$ + Space(3) + Format(grid1.Text, "@@@@@") + Space$(4)
       
       grid1.Col = 7  ' Company
       lineac$ = lineac$ + Format(grid1.Text, "!@@@@@@@@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 8  ' idemployee
       lineac$ = lineac$ + Format(grid1.Text, "@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 9  ' USR
       lineac$ = lineac$ + Format(grid1.Text, "!@@@@@@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 10  ' CSR
       lineac$ = lineac$ + Format(grid1.Text, "!@@@@@@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 11  ' idoffice
       lineac$ = lineac$ + Format(grid1.Text, "@@@@@@@@") + Space$(1)
       
       grid1.Col = 12  ' Office
       lineac$ = lineac$ + Format(grid1.Text, "!@@@@@@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 13  ' Fiduciary
       lineac$ = lineac$ + Format(Format(grid1.Text, "###,##0.00"), "@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 14  ' Total Receipt
       lineac$ = lineac$ + Format(Format(grid1.Text, "###,##0.00"), "@@@@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 15  ' Amount paid
       lineac$ = lineac$ + Format(Format(grid1.Text, "###,##0.00"), "@@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 16  ' PYMT Method
       lineac$ = lineac$ + Format(Format(grid1.Text, "###,##0.00"), "!@@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 17  ' Balance due
       lineac$ = lineac$ + Format(Format(grid1.Text, "###,##0.00"), "@@@@@@@@@@@") + Space$(1)
       
       grid1.Col = 18  ' BD date
       lineac$ = lineac$ + Format(grid1.Text, "!@@@@@@@@@@") + Space$(1)
       
        
       Printer.Print lineac$
       
       If grid1.Rows <= 1 Then Exit For
Next t

Printer.Print " "
Printer.Print " "
Printer.Print " "
Printer.Print " "

Printer.FontSize = 9
Printer.Print Space(5) + "VOID RECEIPTS:"
Printer.Print " "
Printer.FontSize = 5



For t = 0 To grid3.Rows
  grid3.Row = t
  lineac$ = ""
       If t = 0 Then
       
       grid3.Col = 0  ' num
       lineac$ = lineac$ + Format("Row#", "@@@") + Space$(1)
       
       Else
       grid3.Col = 0  ' num
       lineac$ = lineac$ + Format(grid3.Text, "@@@") + Space$(1)
  
       End If
       
  
       grid3.Col = 1  ' Receipt#
       lineac$ = lineac$ + Format(grid3.Text, "@@@@@@@@@") + Space$(1)
       
       
       grid3.Col = 2  ' Date
       a$ = Left(grid3.Text, 10)
       lineac$ = lineac$ + Format(a$, "!@@@@@@@@@@") + Space$(1)
  
       grid3.Col = 3  ' IDcust
       lineac$ = lineac$ + Format(grid3.Text, "!@@@@@@") + Space$(1)
       
       grid3.Col = 4  ' Customer name
       lineac$ = lineac$ + Format(grid3.Text, "!@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 5  ' Policy#
       lineac$ = lineac$ + Format(grid3.Text, "!@@@@@@@@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 6  ' IDCompany
       lineac$ = lineac$ + Space(3) + Format(grid3.Text, "@@@@@") + Space$(4)
       
       grid3.Col = 7  ' Company
       lineac$ = lineac$ + Format(grid3.Text, "!@@@@@@@@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 8  ' idemployee
       lineac$ = lineac$ + Format(grid3.Text, "@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 9  ' USR
       lineac$ = lineac$ + Format(grid3.Text, "!@@@@@@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 10  ' CSR
       lineac$ = lineac$ + Format(grid3.Text, "!@@@@@@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 11  ' idoffice
       lineac$ = lineac$ + Format(grid3.Text, "@@@@@@@@") + Space$(1)
       
       grid3.Col = 12  ' Office
       lineac$ = lineac$ + Format(grid3.Text, "!@@@@@@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 13  ' Fiduciary
       lineac$ = lineac$ + Format(Format(grid3.Text, "###,##0.00"), "@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 14  ' Total Receipt
       lineac$ = lineac$ + Format(Format(grid3.Text, "###,##0.00"), "@@@@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 15  ' Amount paid
       lineac$ = lineac$ + Format(Format(grid3.Text, "###,##0.00"), "@@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 16  ' PYMT Method
       lineac$ = lineac$ + Format(Format(grid3.Text, "###,##0.00"), "!@@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 17  ' Balance due
       lineac$ = lineac$ + Format(Format(grid3.Text, "###,##0.00"), "@@@@@@@@@@@") + Space$(1)
       
       grid3.Col = 18  ' BD date
       lineac$ = lineac$ + Format(grid3.Text, "!@@@@@@@@@@") + Space$(1)
       
        
       Printer.Print lineac$
       
       If grid3.Rows <= 1 Then Exit For
       
Next t


End If





Printer.EndDoc
msg.Visible = False

End Sub

Private Sub btnprinter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnprinter.Picture = img_print_down.Picture

End Sub

Private Sub btnprinter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnprinter.Picture = img_print_up.Picture
End Sub












Private Sub btnpunto_Click()
On Error Resume Next
pos = InStr(1, txtoutput.Text, ".")
If pos > 0 Then Exit Sub
txtoutput.Text = txtoutput.Text & "."
End Sub

Private Sub btnrefresh_total_lae_agente_Click()
On Error Resume Next

calcula_total_LAE


    
End Sub

Private Sub btnrevisar_Click()
On Error Resume Next
btncargar_reportes_Click

End Sub

Private Sub btnsave_Click()
On Error Resume Next

lblmsg1.Caption = "Saving the information"

msg.Visible = True
msg.Refresh

 If cbo_managers.ListIndex = -1 Then
     cbo_managers.ListIndex = 0
 End If

 graba_datos


msg.Visible = False
msg.Refresh

lblmsg1.Caption = "Loading the information"
modificado = 0




End Sub

Private Sub btnsave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnsave.Picture = img_disk_down.Picture
End Sub


Private Sub btnsave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnsave.Picture = img_disk_up.Picture

End Sub


Private Sub btnsend_Click()
On Error Resume Next

If chkmanager.Value = 0 And (Val(cargo$) = 17 Or Val(cargo$) = 24) And tipo_guardado = 2 Then
  MsgBox "You need to sign the report", 16, "Attention"
  Exit Sub
End If

If chk_firma_agente.Value = 0 And Val(cargo$) = 16 And tipo_guardado = 1 Then
  MsgBox "You need to sign the report", 16, "Attention"
  Exit Sub
End If





If bloqueado = 1 And Format(lbldate_agente.Caption, "mm/dd/yyyy") = Format(Now, "mm/dd/yyyy") Then
   MsgBox "Cannot close, you have previous days that haven't been closed.", 16, "Attention"
   Exit Sub
End If


If (Val(Format(lbldate_agente.Caption, "y")) > Val(Format(Now, "y"))) And (Val(Format(lbldate_agente.Caption, "yyyy")) = Val(Format(Now, "yyyy"))) Then
    MsgBox "The selected date is greater than the current date", 16, "Attention"
    Exit Sub
End If



ruta_archivos$ = "\\192.168.84.215\moneyreport\"


If Dir$(ruta_archivos$ + "encendido.txt") = "" Then
     MsgBox "Files cannot be saved. Apparently the storage server is down. Contact the IT department to solve this problem. Thanks", 16, "ATTENTION"
     Exit Sub
End If



If chk_dayoff.Value = 1 And Val(Format(txttotal_LAE_agente.Text, "000000.00")) > 0 Then
      MsgBox "It's not possible to close this day as a rest day because it has a sold balance.", 16, "Attention"
      Exit Sub
End If



r$ = MsgBox("Do you want to send the money report?. Once sent you will not be able to make changes.", 4, "Attention")
If r$ = "7" Then Exit Sub


   btnsave_Click

   Dim sSelect As String
   Dim Rs As ADODB.Recordset


    Set Rs = New ADODB.Recordset
    
   
    oficina$ = LTrim(Right(UCase(RTrim(cbo_oficina.List(cbo_oficina.ListIndex))), 30))
    sSelect = "SELECT idoffice From officescatalog where office='" + oficina$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_office$ = Rs(0)
    Rs.Close
    
    
    
    UserName$ = RTrim(Left(cbo_agentes.List(cbo_agentes.ListIndex), Len(cbo_agentes.List(cbo_agentes.ListIndex)) - 5))
    sSelect = "SELECT idemployee From employeeinfo where username='" + UserName$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_employee$ = Rs(0)
    Rs.Close






If tipo_guardado = 1 Then


   cuenta = 0


  If ListView1.ListItems.Count = 0 And chk_dayoff.Value = 0 And Val(Format(txttotal_LAE_agente.Text, "000000.00")) > 0 Then
      MsgBox "Scanned documents have not been attached yet", 16, "Attention"
      Exit Sub
  End If



    
una_vez_mas:
    
    
    
     ' verifica si ya existe el reporte
     
     sSelect = "select idmoneyreport from moneyreport where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "' and idoffice='" + id_office$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreport$ = Rs(0)
     Rs.Close
     
    
     If id_moneyreport$ = "" Then
        btnsave_Click
        cuenta = cuenta + 1
        If cuenta = 1 Then
           GoTo una_vez_mas
        Else
           MsgBox "The money report couldn't save the info", 16, "Attention"
           Exit Sub
        End If
     End If
     
    
    submitido = ""
    
    sSelect = "SELECT submitted From moneyreport where idmoneyreport='" + id_moneyreport$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    submitido = Rs(0)
    Rs.Close
    
    
    
    
    
    If submitido = "False" Then
    
      sSelect = "update MoneyReport set submitted='1' " & _
      "where idmoneyreport='" + id_moneyreport$ + "'"  ' datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "' and idoffice='" + id_office$ + "'"

                      
      Rs.Open sSelect, base, adOpenUnspecified
    
      Rs.Close
    
    End If
    
    
  
    
    
    ' *****************************************************************************************************************************
    ' graba las tablas de SQL de receipts/reports
    
    
    
    For t = 1 To grid1.Rows - 1
       grid1.Row = t
       grid1.Col = 1
       recibox$ = grid1.Text
       
       grid1.Col = 2
       fechax$ = Format(grid1.Text, "mm/dd/yyyy")
       
       grid1.Col = 3
       idcustx$ = grid1.Text
       
       grid1.Col = 4
       Nombrex$ = grid1.Text
       
       grid1.Col = 5
       polizax$ = grid1.Text
       
       grid1.Col = 6
       idcompanyx$ = grid1.Text
       
       grid1.Col = 7
       companyx$ = grid1.Text
       
       grid1.Col = 8
       idemployeex$ = grid1.Text
       
       grid1.Col = 9
       USRx$ = grid1.Text
       
       grid1.Col = 10
       CSRx$ = grid1.Text
       
       grid1.Col = 11
       Idoficinax$ = grid1.Text
       
        grid1.Col = 12
       Oficinax$ = grid1.Text
       
        grid1.Col = 13
       Fiduciaryx$ = grid1.Text
       
        grid1.Col = 14
       Total_recibox$ = grid1.Text
       
        grid1.Col = 15
       Cantidadx$ = grid1.Text
       
        grid1.Col = 16
       PYMTx$ = grid1.Text
       
        grid1.Col = 17
       Balancex$ = grid1.Text
       
        grid1.Col = 18
       BDx$ = grid1.Text
       
    
       sSelect = "insert into MoneyReportreceipts (idreceipthdr, date, idcustomer, customername, policynumber, idcompany,companyname, idemployee, usr, csr, idoffice, office, " & _
       "fiduciary, totalreceipt, amountpaid, paymentmethod, balancedue, balanceduedate, void, active)  VALUES ('" & _
       recibox$ + "', '" + fechax$ + "', '" + idcustx$ + "', '" + Nombrex$ + "', '" + polizax$ + "', '" + idcompanyx$ + "', '" + companyx$ + "', '" + idemployeex$ + "', '" & _
       USRx$ + "', '" + CSRx$ + "', '" + Idoficinax$ + "', '" + Oficinax$ + "', '" + Fiduciaryx$ + "', '" + Total_recibox$ + "', '" + Cantidadx$ + "', '" + PYMTx$ + "', '" & _
       Balancex$ + "', '" + BDx$ + "', '0', '1')"
       
       Rs.Open sSelect, base, adOpenUnspecified
    
       Rs.Close
    
    Next t
    
    
    
    For t = 1 To grid3.Rows - 1
       grid3.Row = t
       grid3.Col = 1
       recibox$ = grid3.Text
       
       grid3.Col = 2
       fechax$ = Format(grid3.Text, "mm/dd/yyyy")
       
       grid3.Col = 3
       idcustx$ = grid3.Text
       
       grid3.Col = 4
       Nombrex$ = grid3.Text
       
       grid3.Col = 5
       polizax$ = grid3.Text
       
       grid3.Col = 6
       idcompanyx$ = grid3.Text
       
       grid3.Col = 7
       companyx$ = grid3.Text
       
       grid3.Col = 8
       idemployeex$ = grid3.Text
       
       grid3.Col = 9
       USRx$ = grid3.Text
       
       grid3.Col = 10
       CSRx$ = grid3.Text
       
       grid3.Col = 11
       Idoficinax$ = grid3.Text
       
        grid3.Col = 12
       Oficinax$ = grid3.Text
       
        grid3.Col = 13
       Fiduciaryx$ = grid3.Text
       
        grid3.Col = 14
       Total_recibox$ = grid3.Text
       
        grid3.Col = 15
       Cantidadx$ = grid3.Text
       
        grid3.Col = 16
       PYMTx$ = grid3.Text
       
        grid3.Col = 17
       Balancex$ = grid3.Text
       
        grid3.Col = 18
       BDx$ = grid3.Text
       
    
       sSelect = "insert into MoneyReportreceipts (idreceipthdr, date, idcustomer, customername, policynumber, idcompany,companyname, idemployee, usr, csr, idoffice, office, " & _
       "fiduciary, totalreceipt, amountpaid, paymentmethod, balancedue, balanceduedate, void, active)  VALUES ('" & _
       recibox$ + "', '" + fechax$ + "', '" + idcustx$ + "', '" + Nombrex$ + "', '" + polizax$ + "', '" + idcompanyx$ + "', '" + companyx$ + "', '" + idemployeex$ + "', '" & _
       USRx$ + "', '" + CSRx$ + "', '" + Idoficinax$ + "', '" + Oficinax$ + "', '" + Fiduciaryx$ + "', '" + Total_recibox$ + "', '" + Cantidadx$ + "', '" + PYMTx$ + "', '" & _
       Balancex$ + "', '" + BDx$ + "', '1', '1')"
       
       Rs.Open sSelect, base, adOpenUnspecified
    
       Rs.Close
    
    Next t
    
    
    carga_registros
    
    
 ElseIf tipo_guardado = 2 Then
 
 ' ***********************************************************************************************************
   ' verifica en reporte de oficina
   
   ' verifica que todos los usuarios hayan cerrado su money report
   
   
   If ListView2.ListItems.Count = 0 Then
      MsgBox "Scanned documents have not been attached yet", 16, "Attention"
      Exit Sub
   End If
   
   Dim aviso_agentes$(10), submitidito As Boolean
   
   Erase aviso_agentes$
   
   sSelect = "select distinct IDEmployee, Username from EmployeeInfo emp " & _
   "inner join ReceiptsHDR rechdr on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
   "where emp.Active=1 and IdOffice='" + id_office$ + "' and IdJobtitleUSR in (16,17,28, 2,37) and rechdr.Active=1 " & _
   "and cast(rechdr.Date as Date) >= '" + lbldate_agente.Caption + "' AND cast( rechdr.DATE as Date) <= '" + lbldate_agente.Caption + "'"

   ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    

   
   PERMISO$ = "1"
   conta = 0
   For z = 1 To Grid2.Rows - 1
    Grid2.Row = z
    Grid2.Col = 2
    UserName$ = Grid2.Text
    
    Grid2.Col = 1
    id_employee$ = Grid2.Text
    
   
    'UserName$ = RTrim(Left(cbo_agentes.List(z), Len(cbo_agentes.List(z)) - 5))
    'sSelect = "SELECT idemployee From employeeinfo where username='" + UserName$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    'Rs.Open sSelect, base, adOpenUnspecified
    
    'id_employee$ = Rs(0)
    'Rs.Close
   
    id_moneyreport$ = ""
    sSelect = "select idmoneyreport from moneyreport where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "'"  ' and idoffice='" + id_office$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_moneyreport$ = Rs(0)
    Rs.Close
     
        
    If id_moneyreport$ = "" Then
      submitidito = False
    
    Else
        
     sSelect = "select submitted from moneyreport where idmoneyreport='" + id_moneyreport$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     submitidito = Rs(0)
     Rs.Close
    
    End If
    
    
    If submitidito = False Then
        PERMISO$ = "0"
        aviso_agentes$(conta) = UserName$
        conta = conta + 1
    End If
   
   
   Next z
   
   
   If PERMISO$ = "0" Then
      a$ = "The following employees have not closed their money report for the day, therefore the office money report cannot be closed." + Chr$(13)
      For z = 0 To conta
          a$ = a$ + Chr$(13) + aviso_agentes$(z)
      Next z
      MsgBox a$, 64, "Attention"
      Exit Sub
   End If
   
   
   
   
   
    UserName$ = RTrim(Left(cbo_managers.List(cbo_managers.ListIndex), Len(cbo_managers.List(cbo_managers.ListIndex)) - 5))
    sSelect = "SELECT idemployee From employeeinfo where username='" + UserName$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    ID_manager$ = Rs(0)
    Rs.Close
   
   
     
     sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idmanager='" + ID_manager$ + "'"   '  " and idoffice='" + id_office$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreportoffice$ = Rs(0)
     Rs.Close
     
    
     If id_moneyreportoffice$ = "" Then
        btnsave_Click
        cuenta = cuenta + 1
        If cuenta = 1 Then
           GoTo una_vez_mas
        Else
           MsgBox "The money report couldn't save the info", 16, "Attention"
           Exit Sub
        End If
     End If
     
    
    submitido = ""
    
    sSelect = "SELECT submitted From moneyreportbyoffice where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    submitido = Rs(0)
    Rs.Close
    
    
   
    
    
    If submitido = "False" Then
    
      sSelect = "update MoneyReportbyoffice set submitted='1'" & _
      "where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"  ' datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "' and idoffice='" + id_office$ + "'"

                      
      Rs.Open sSelect, base, adOpenUnspecified
    
      Rs.Close
    
    End If
    
    
    carga_registros
 
 
 
 End If
 
      
    
      obtener_fecha_real
   
      MsgBox "Done!. Your report was sent successfully.", 64, "A thumb up!"

End Sub

Private Sub btnsend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnsend.Picture = img_send_down.Picture

End Sub


Private Sub btnsend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnsend.Picture = img_send_up.Picture

End Sub


Private Sub btnupdate_LAE_Click()
On Error Resume Next

If valido1 = 777 Then Exit Sub

Dim sSelect As String
    
Dim Rs As ADODB.Recordset
    
    
Set Rs = New ADODB.Recordset
           
  
  'oficina$ = LTrim(RTrim(Right(cbo_oficina.List(cbo_oficina.ListIndex), 30)))
  '
  
  oficina$ = lbloficina_agente.Caption
  
  oficina$ = LTrim(Right(UCase(RTrim(cbo_oficina.List(cbo_oficina.ListIndex))), 30))
  
  lbloficina_agente.Caption = oficina$
  
  Grid2.Clear
    
  
 
   sSelect = "Select sum(AmountPaid) as 'Total Amount', rechdr.IdOffice, ofc.Office from ReceiptsHDR rechdr " & _
   "inner join OfficesCatalog ofc on rechdr.IdOffice=ofc.IdOffice " & _
   "where cast(Date as Date) = '" + lbldate_agente.Caption + "' and rechdr.Void=0 group by rechdr.IdOffice, ofc.Office"
   
 
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    cantidad$ = ""
    For t = 1 To Grid2.Rows - 1
       Grid2.Row = t
       Grid2.Col = 3
       oficina$ = UCase(Grid2.Text)
    
       If UCase(lbloficina_agente.Caption) = oficina$ Then
           Grid2.Col = 1
           cantidad$ = Grid2.Text
           Exit For
       End If
       
    Next t
       
    txttotal_venta_manager.Text = cantidad$
    
    lbltotal_LAE_oficina.Caption = Format(txttotal_venta_manager.Text, "$###,##0.00")

    valido1 = 888

End Sub



Private Sub Calendar1_Click()
On Error Resume Next
msg.Visible = True
lbldate_agente.Caption = Calendar1.Value
Calendar1.Visible = False
valido1 = 0
hoja1.Visible = False
Hoja2.Visible = False
hoja3.Visible = False
  
  
  
  
Dim sSelect As String
    
Dim Rs As ADODB.Recordset

Set Rs = New ADODB.Recordset


lae_office$ = RTrim(LTrim(Right(Form1.cbo_oficina.List(Form1.cbo_oficina.ListIndex), 25)))
 
      
   If cbo_agentes.Enabled = True Then
    For t = 0 To cbo_agentes.ListCount - 1
       If UCase(RTrim(LTrim(Left(cbo_agentes.List(t), 30)))) = user$ Then
          cbo_agentes.ListIndex = t
          Exit For
       End If
    Next t
   End If
    
    ' checa la oficina y la actualiza en la barra
    'sSelect = "select office from officescatalog where idoffice='" + id_oficina$ + "'"
    'Rs.Open sSelect, base, adOpenUnspecified
    'oficina$ = Rs(0)
    'Rs.Close
    'lbloficina_agente.Caption = oficina$
    
    'lbloficina_agente.Caption = oficina_trabajada$
    

    
    sSelect = "select ofc.Office from ReceiptsHDR rechdr " & _
    "inner join OfficesCatalog ofc on ofc.IdOffice= rechdr.IdOffice " & _
    "inner join EmployeeInfo emp on emp.IDEmployee =  rechdr.IdEmployeeUSR " & _
    "where  cast(rechdr.Date as Date) >= '" + lbldate_agente.Caption + "' " & _
    "AND cast( rechdr.Date as Date) <= '" + lbldate_agente.Caption + "' and emp.Username='" + user$ + "'"
    
    ' group by ofc.Office"
    Rs.Open sSelect, base, adOpenUnspecified
    oficina$ = Rs(0)
    Rs.Close
    lbloficina_agente.Caption = oficina$
    
    
    
    
    
    If lbloficina_agente.Caption = "" Then
          lbloficina_agente.Caption = lae_office$
    End If
    
  
  
    sSelect = "select idoffice from officescatalog where office='" + oficina$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_oficina$ = Rs(0)
    Rs.Close
  
    
   
 
  
    existe = 0
    For t = 0 To cbo_oficina.ListCount - 1
      oficina_en_lista$ = LTrim(RTrim(Right(cbo_oficina.List(t), 30)))
      If UCase(oficina_en_lista$) = UCase(oficina$) Then
          cbo_oficina.ListIndex = t
          existe = 1
          Exit For
      End If
    Next t
    
    
    If existe = 0 Then
       If UCase(oficina$) <> "" Then
    
        cbo_oficina.AddItem UCase(oficina$) + Space(30) + UCase(oficina$)
        
       End If
    
        
       
        For Y = 0 To cbo_oficina.ListCount - 1
            oficina_en_lista$ = LTrim(RTrim(Right(cbo_oficina.List(Y), 30)))
            If UCase(oficina_en_lista$) = UCase(oficina$) Then
                cbo_oficina.ListIndex = Y
                existe = 1
                Exit For
            End If
        Next Y
       
    End If
    
  
           
       
  
  
carga_todo


valido1 = 999
Load forma_reportes_pendientes
forma_reportes_pendientes.Show 1

If Val(transfiere$) > 0 Then
  Timer1.Enabled = True
  Picture2.Visible = True
Else
  Picture2.Visible = False
  Timer1.Enabled = False
End If

Refresh

obtener_fecha_real


If tipo_guardado = 1 Then
  chk_dayoff.Visible = True
End If


msg.Visible = False
End Sub

Private Sub Calendar1_LostFocus()
Calendar1.Visible = False

End Sub


Private Sub cbo_agentes_Click()
On Error Resume Next
' verifica si es agente o manager el usuario

If valido1 = 777 Then
   Exit Sub
End If

If agente$ = manager$ Then
  estado_carga = 1
Else
  estado_carga = 0
End If


valido1 = 0




carga_archivos
carga_archivos2

obtener_fecha_real
 
limpia_datos

img_tab1_Click (0)


carga_registros

carga_datos

calcula_total_LAE




valido1 = 999
Load forma_reportes_pendientes
forma_reportes_pendientes.Show 1

If Val(transfiere$) > 0 Then
  Timer1.Enabled = True
  Picture2.Visible = True
Else
  Picture2.Visible = False
  Timer1.Enabled = False
End If

Refresh



lbl_iniciales_agente.Caption = LTrim(RTrim(Left(cbo_agentes.List(cbo_agentes.ListIndex), Len(cbo_agentes.List(cbo_agentes.ListIndex)) - 10)))

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
           
  
 
' carga el nombre completo

       sSelect = "select firstname, lastname1 from employeeinfo where username='" + lbl_iniciales_agente.Caption + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       nombre$ = Rs(0)
       apellido$ = Rs(1)
       Rs.Close

       lblname_agent.Caption = nombre$ + Space(1) + apellido$
       'lbl_iniciales_agente.Caption = user$
   



Desactiva_sel

End Sub


Private Sub cbo_managers_Click()
On Error Resume Next

If valido1 = 777 Then Exit Sub

' verifica si es agente o manager el usuario
If agente$ = manager$ Then
  estado_carga = 1
Else
  estado_carga = 0
End If

valido1 = 0

' estado_registro = 3
estado_registro = 1


' obtener_fecha_real

carga_archivos
carga_archivos2



limpia_datos


 
carga_registros

carga_datos

calcula_total_LAE

Desactiva_sel


End Sub


Private Sub cbo_oficina_Click()
On Error Resume Next

If valido1 = 777 Then
   'valido1 = 1
   Exit Sub
End If





guarda_cbo_agente = cbo_agentes.List(cbo_agentes.ListIndex)

msg.Visible = True
'lbldate_agente.Caption = Calendar1.Value
Calendar1.Visible = False

hoja1.Visible = False
Hoja2.Visible = False
hoja3.Visible = False
  
  
carga_todo
carga_agentes

valido1 = 777


If cbo_agentes.Enabled = True Then
 For Y = 0 To cbo_agentes.ListCount - 1
   If cbo_agentes.List(Y) = guarda_cbo_agente Then
           cbo_agentes.ListIndex = Y
           Exit For
   End If
 Next Y
End If


valido1 = 0
calcula_total_LAE



 
cbo_managers.ListIndex = 0



      oficina_trabajada$ = LTrim(RTrim(Right(cbo_oficina.List(cbo_oficina.ListIndex), 30)))
      lbloficina_agente.Caption = oficina_trabajada$


   
   

valido1 = 0
msg.Visible = False
End Sub

Private Sub cbo_year_Change()
On Error Resume Next
txtdate1.Text = ""
txtdate2.Text = ""
    
End Sub

Private Sub cboimpre_Click()
On Error Resume Next


For Each xprint In Printers
           If xprint.DeviceName = cboimpre.Text Then
              ' La define como predeterminada del sistema.
              Set Printer = xprint
              DoEvents
              Exit For
           End If
Next


nf = FreeFile
 Open "c:\discrepancy\printer" For Output Shared As #nf
 Lock #nf
 Print #nf, Printer.DeviceName
 Print #nf, Printer.Port
 Unlock #nf
 Close #nf
 
 
End Sub


Public Sub carga_impresoras()
On Error Resume Next

Dim cImprGen As String
    cImprGen = cboimpre.Text
    
cboimpre.Clear
' ruta$ = "c:\discrepancy\"
    
If Dir$(ruta$ + "printer") <> "" Then
 nf = FreeFile
 Open ruta$ + "printer" For Input Shared As #nf
 Lock #nf
 Line Input #nf, P1$
 Line Input #nf, P2$
 Unlock #nf
 Close #nf
 
 cImprGen = P1$
 cboimpre.Text = P1$

End If
    
    
    
    
For Each xprint In Printers
           If xprint.DeviceName = cImprGen Then
              ' La define como predeterminada del sistema.
              Set Printer = xprint
              DoEvents
              Exit For
           End If
Next
        
        
        
For Each xprint In Printers
        cboimpre.AddItem xprint.DeviceName
Next
        
        
nf = FreeFile
 Open ruta$ + "printer" For Output Shared As #nf
 Lock #nf
 Print #nf, Printer.DeviceName
 Print #nf, Printer.Port
 Unlock #nf
 Close #nf
 
 
 For t = 0 To cboimpre.ListCount - 1
   If cboimpre.List(t) = Printer.DeviceName Then
       cboimpre.ListIndex = t
       Exit For
   End If
 Next t
        
        
        
        
End Sub

Private Sub chk_dayoff_Click()
On Error Resume Next
msgdescanso.Visible = chk_dayoff.Value
End Sub

Private Sub chk_firma_agente_Click()
On Error Resume Next
Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
    
    
     
 
  id_agente = Val(Right(Form1.cbo_agentes.List(Form1.cbo_agentes.ListIndex), 20))
 
 ' carga el nombre completo

       sSelect = "select firstname, lastname1 from employeeinfo where idemployee='" + id_agente + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       nombre$ = Rs(0)
       apellido$ = Rs(1)
       Rs.Close
 
    
If chk_firma_agente.Value = 1 Then
  Firma_agente.Caption = nombre$ + " " + apellido$
Else
  Firma_agente.Caption = ""
End If
End Sub

Private Sub chkagentes_Click(Index As Integer)
On Error Resume Next
carga_agentes
End Sub






Private Sub chkmanager_Click()
On Error Resume Next
Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
    
    
     
 
  ID_manager = Val(Right(Form1.cbo_managers.List(Form1.cbo_managers.ListIndex), 20))
 
 ' carga el nombre completo

       sSelect = "select firstname, lastname1 from employeeinfo where idemployee='" + ID_manager + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       nombre$ = Rs(0)
       apellido$ = Rs(1)
       Rs.Close
 
    
If chkmanager.Value = 1 Then
  firma.Caption = nombre$ + " " + apellido$
Else
  firma.Caption = ""
End If



End Sub

Private Sub Form_Load()
On Error Resume Next
Left = (Screen.Width - Width) / 2
Top = 0

If (App.PrevInstance = True) Then
  X$ = Shell("cmd /c taskkill /f /im money.exe")

  base.Close
  End
End If

MkDir "c:\money"
ruta$ = "c:\money\"


    
    
permiso_carga = 0
limpia_tarjetas
tipo_guardado = 1
cbo_oficina.Clear

Calendar1.Value = Format(Now, "mm/dd/yyyy")

If administrator = 1 Then
   tipo_guardado = 2
End If

actualiza_google


X$ = Shell("cmd /c taskkill /f /im barra_agent.exe")
X$ = Shell("cmd /c taskkill /f /im barra_agent.exe")
X$ = Shell("cmd /c taskkill /f /im barra_agent.exe")

    Kill "c:\iconos\barra_agent.exe"
    
    FileCopy "\\192.168.84.215\moneyreport\barra_agent.exe", "c:\iconos\barra_agent.exe"
    



If Dir$("\\192.168.84.215\moneyreport\copia_barra") <> "" Then
    
    X$ = Shell("cmd /c taskkill /f /im barra_agent.exe")
    
    Kill "c:\iconos\barra_agent.exe"
    FileCopy "\\192.168.84.215\moneyreport\barra_agent.exe", "c:\iconos\barra_agent.exe"
    
     

    GoTo salta
End If


If administrador = 1 Then
  txtagente.Visible = True
  btnadd_agent.Visible = True
End If


If Dir$("\\192.168.84.215\moneyreport\copia_activada") <> "" Then
    X$ = Shell("cmd /c taskkill /f /im reset.exe")
    FileCopy "\\192.168.84.215\moneyreport\reset.exe", "c:\money\reset.exe"
    
    X$ = Shell("cmd /c taskkill /f /im barra_agent.exe")
    
    Kill "c:\iconos\barra_agent.exe"
    FileCopy "\\192.168.84.215\moneyreport\JA-authorize_X.ico", "c:\iconos\JA-authorize_X.ico"
    
    FileCopy "\\192.168.84.215\moneyreport\barra_agent.exe", "c:\iconos\barra_agent.exe"
    
    
    

    GoTo salta
End If





If Dir$("c:\money\reset.exe") = "" Then
  
  X$ = Shell("cmd /c taskkill /f /im reset.exe")
  FileCopy "\\192.168.84.215\moneyreport\reset.exe", "c:\money\reset.exe"
  
End If


If Dir$("c:\iconos\barra_agent.exe") = "" Then
  
  X$ = Shell("cmd /c taskkill /f /im barra_agent.exe")
  Kill "c:\iconos\barra_agent.exe"
  FileCopy "\\192.168.84.215\moneyreport\barra_agent.exe", "c:\iconos\barra_agent.exe"
  
End If



salta:



Carga_todas_las_oficinas

fecha_entrada$ = Format(Now, "mm/dd/yyyy")

bloqueado = 0





'For Y = 0 To 5
' cbooficina1(Y).Clear
 
' For t = 1 To grid4.Rows - 1
'    grid4.Row = t
'    grid4.Col = 2
'    cbooficina1(Y).AddItem grid4.Text + Space(30) + grid4.Text
    
' Next t
'Next Y





ano_actual = Val(Format(Now, "yyyy"))


cbo_year.Clear
cbo_year.AddItem Str(ano_actual - 3)
cbo_year.AddItem Str(ano_actual - 2)
cbo_year.AddItem Str(ano_actual - 1)
cbo_year.AddItem Str(ano_actual)

cbo_year.ListIndex = cbo_year.ListCount - 1





'Conecta_SQL

estado_registro = 1


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
          
           ScaleFactorX = 1360 / DesignX
           ScaleFactorY = 1 ' 1024 / DesignY
        
        
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


 a$ = GetIPHostName()
  'a$ = InputBox("nombre de pc")
  b$ = GetIPAddress()
  
  
veces = 0
cont = 0
' remueve la ultima parte del ip
For t = Len(b$) To 1 Step -1
  c$ = Mid$(b$, t, 1)
  cont = cont + 1
  If c$ = "." Then
     veces = veces + 1
     If veces = 2 Then
        
        last_digits$ = Right(b$, cont - 1)
        pos = InStr(1, last_digits$, ".")
        last_digits$ = Left(last_digits$, pos - 1)
        Exit For
     End If
  End If
Next t
  
'Frame1.Visible = False

  


valido1 = 777

aprobado = 0

For z = 0 To 4

 If UCase$(oficina_guardada$(z)) = "" Then
   Exit For
 End If
 
 If UCase$(oficina_guardada$(z)) = "INDEPENDENT" Then
   aprobado = 1
   cbo_oficina.ListIndex = 11
   Exit For
 End If
 
 If UCase$(oficina_guardada$(z)) = "RETENTIONS" Then
   aprobado = 1
   cbo_oficina.ListIndex = 12
   Exit For
 End If
 

Select Case last_digits$
Case 39
  cbo_oficina.ListIndex = 0
  If UCase$(oficina_guardada$(z)) = "JA - ARLETA" Then
     aprobado = 1
  End If
  
Case 43
  cbo_oficina.ListIndex = 1
  If UCase$(oficina_guardada$(z)) = "JA - COMPTON" Then
     aprobado = 1
  End If
  
Case 54
  'cbo_oficina.ListIndex = 1
  'If UCase$(oficina_guardada$(z)) = "JA - ECHO PARK" Then
  '   aprobado = 1
  'End If
  
Case 45
  cbo_oficina.ListIndex = 2
   If UCase$(oficina_guardada$(z)) = "JA - CITRUS" Then
     aprobado = 1
  End If
  
Case 49
  cbo_oficina.ListIndex = 3
  If UCase$(oficina_guardada$(z)) = "JA - FLORENCE" Then
     aprobado = 1
  End If
  
Case 84
  
  'Frame1.Visible = True
  If UCase$(oficina_guardada$(z)) = "JA - HAVEN" Then
     cbo_oficina.ListIndex = 4
     aprobado = 1
  ElseIf UCase$(oficina_guardada$(z)) = "JA - PHONE SALES" Then
     cbo_oficina.ListIndex = 5
     aprobado = 1
  End If
    
Case 47
  cbo_oficina.ListIndex = 6
    If UCase$(oficina_guardada$(z)) = "JA - SAN BERNARDINO" Then
     aprobado = 1
  End If
  
Case 41
  cbo_oficina.ListIndex = 7
  If UCase$(oficina_guardada$(z)) = "JA - 17TH ST" Then
     aprobado = 1
  End If
  
Case 46
  cbo_oficina.ListIndex = 8
  If UCase$(oficina_guardada$(z)) = "JA - VANOWEN" Then
     aprobado = 1
  End If
  
Case 23
  cbo_oficina.ListIndex = 9
   If UCase$(oficina_guardada$(z)) = "JA - WHITTIER" Then
     aprobado = 1
  End If
  
Case 31
  cbo_oficina.ListIndex = 10
   If UCase$(oficina_guardada$(z)) = "JA - MONTERREY" Then
     aprobado = 1
  End If
  
Case 22
  cbo_oficina.ListIndex = 10
   If UCase$(oficina_guardada$(z)) = "JA - PONDEROSA" Then
     aprobado = 1
  End If
  
  
  
End Select

Next z



If aprobado = 0 And administrador = 0 Then
   'MsgBox "This user can not use this program in this office", 16, "Access denied"
   'base.Close
   'End
End If



carga_oficinas
carga_agentes


  
For t = 0 To cbo_oficina.ListCount - 1
  pos = InStr(1, cbo_oficina.List(t), " ")
  a$ = RTrim(LTrim(Mid$(cbo_oficina.List(t), pos, Len(cbo_oficina.List(t)) - pos + 1)))
  If UCase(a$) = UCase(oficina_guardada$(0)) Then
      cbo_oficina.ListIndex = t
      Exit For
  End If
Next t
  
  
  
 modificado = 0
  



  
  
  
lae_office$ = RTrim(LTrim(Right(Form1.cbo_oficina.List(Form1.cbo_oficina.ListIndex), 25)))
  


carga_impresoras


For t = 0 To cbo_agentes.ListCount - 1
   a$ = RTrim(UCase(Left(cbo_agentes.List(t), Len(cbo_agentes.List(t)) - 15)))
   If user$ = a$ Then
       cbo_agentes.ListIndex = t
       Exit For
   End If
Next t

cbo_managers.ListIndex = 0



' verifica si es agente o manager el usuario
If agente$ = manager$ Then
  estado_carga = 1
Else
  estado_carga = 0
End If


lbladministrator.Caption = name_admon$
  
If administrador = 1 Then
  txtnum_nota.Visible = True
  lblnota.Visible = True
  btnlimpia_comen2.Visible = True
  
  lbladministrator.Visible = True
  lbladmon.Visible = True
  cbo_oficina.Enabled = True
 
  btnmostrar_todo.Visible = True
  
  marco_revisado.Visible = True
  
  
  For Y = 0 To cbo_oficina.ListCount - 1
     If Left(UCase(cbo_oficina.List(Y)), 5) = "HAVEN" Then
          cbo_oficina.ListIndex = Y
          Exit For
     End If
  Next Y
  
 If cbo_agentes.Enabled = True Then
  cbo_agentes.ListIndex = 0
 End If
 
  cbo_managers.ListIndex = 0
  
Else
  
  
End If
  
  
  ' 16,17,28,2
  
If (Val(cargo$) = 37 Or Val(cargo$) = 16 Or Val(cargo$) = 28 Or Val(cargo$) = 18 Or Val(cargo$) = 2 Or (Val(cargo$) = 3)) And administrador = 0 Then
  cbo_agentes.Enabled = False
  
  chkagentes(1).Enabled = False

  
ElseIf (Val(cargo$) = 17 Or Val(cargo$) = 24) And administrador = 0 Then
  cbo_agentes.Enabled = True
  
  chkagentes(1).Enabled = True
  cbo_oficina.Enabled = True
  
End If


ano_actual = Val(Format(Now, "yyyy"))
mes_actual = Val(Format(Now, "mm"))

Select Case mes_actual
Case 1, 3, 5, 7, 8, 10, 12
  dia_final = 31
Case 2
  residuo = ((ano_actual / 4) - Int(ano_actual / 4))
  If residuo > 0 Then
     dia_final = 28
  Else
     dia_final = 29
  End If
  
Case 4, 6, 9, 11
  dia_final = 30
End Select





arranque = -1
permiso_carga = 1

lbldate_agente.Caption = Format(Now, "mm/dd/yyyy")

carga_archivos
carga_archivos2

obtener_fecha_real
 
limpia_datos

carga_registros

carga_datos



calcula_total_LAE

carga_combo_de_empleados

btnupdate_LAE_Click



 If (Val(cargo$) = 17 Or Val(cargo$) = 24) Or administrador = 1 Then
 
   
   tabx(0).FillColor = &HFF&
   tabx(2).FillColor = &HC0&
   leyenda(2).ForeColor = &H8080FF
   leyenda(3).ForeColor = &H8080FF
 End If


valido1 = 999
Load forma_reportes_pendientes
forma_reportes_pendientes.Show 1

If Val(transfiere$) > 0 Then
  Timer1.Enabled = True
End If



If Val(transfiere$) > 3 Then
  bloqueado = 1
End If


valido1 = 0

msg.Visible = False
msg.Refresh

modificado = 0


If Left(cbo_oficina.List(0), 9) = "Monterrey" Then
  
  For Y = 0 To cbo_managers.ListCount - 1
      If UCase(Left(cbo_managers.List(Y), 7)) = "TLLANAS" Then
           cbo_managers.ListIndex = Y
           
           Exit For
      End If
  Next Y
End If


If cbo_oficina.ListCount > 1 Then
   btnlock.Visible = True
   btnlock2.Visible = True
Else
  If administrador = 0 Then
   cbo_oficina.Enabled = False
  End If
End If





  
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


Private Sub grid_EnterCell()
On Error Resume Next

grid.Col = 11
IDdisc$ = grid.Text
txtnum_comentario.Text = IDdisc$
txtnum_nota.Text = IDdisc$
txtcalcula.Text = IDdisc$
End Sub









Public Sub carga_registros()
On Error Resume Next
If permiso_carga = 0 Then
   Exit Sub
End If


msg.Visible = True
msg.Refresh

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
           
  
  id_employee = Val(Right(Form1.cbo_agentes.List(Form1.cbo_agentes.ListIndex), 20))
  ID_manager = Val(Right(Form1.cbo_managers.List(Form1.cbo_managers.ListIndex), 20))
  lae_office$ = RTrim(LTrim(Right(Form1.cbo_oficina.List(Form1.cbo_oficina.ListIndex), 25)))
 
  
  If (agente$ = manager$ And agente$ <> "") Or (id_employee = ID_manager And id_employee > 0) Then
    estado_carga = 1
  Else
    estado_carga = 0
  End If
 
  grid1.Visible = False
   
    
  grid3.Visible = False
  grid1.Clear
    
    
     
    oficina$ = LTrim(Right(UCase(RTrim(cbo_oficina.List(cbo_oficina.ListIndex))), 30))
    sSelect = "SELECT idoffice From officescatalog where office='" + oficina$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    id_office_lae$ = Rs(0)
    Rs.Close
     
    
    
    
    sSelect = "select idoffice from moneyreport where idemployee='" + Format(id_employee, "###0") + "' and datereport=convert(datetime, '" + lbldate_agente.Caption + " ') and idoffice='" + id_office_lae$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_oficina$ = Rs(0)
    Rs.Close
    
    
    If id_oficina$ = "" Then
       sSelect = "select idoffice from officescatalog where office='" + lae_office$ + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        id_oficina$ = Rs(0)
        Rs.Close
    End If
    
    
    
    ' checa la oficina y la actualiza en la barra
    sSelect = "select office from officescatalog where idoffice='" + id_oficina$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    oficina$ = Rs(0)
    Rs.Close
    lbloficina_agente.Caption = oficina$
    
    lbloficina_agente.Caption = oficina_trabajada$
    
    
    
    
    
    
    
    fecha_de_revision$ = lbldate_agente.Caption
    
    
       
    
      
    
    id_moneyreport$ = ""
    sSelect = "select idmoneyreport from moneyreport where idoffice='" + id_oficina$ + "' and idemployee='" + Format(id_employee, "###0") + "' and datereport=convert(datetime, '" + lbldate_agente.Caption + " ')"
    Rs.Open sSelect, base, adOpenUnspecified
    id_moneyreport$ = Rs(0)
    Rs.Close
    
    
    id_moneyreportoffice$ = ""
    sSelect = "select idmoneyreportoffice from moneyreportbyoffice where idoffice='" + id_oficina$ + "' and datereport=convert(datetime, '" + lbldate_agente.Caption + " ')"
    Rs.Open sSelect, base, adOpenUnspecified
    id_moneyreportoffice$ = Rs(0)
    Rs.Close
    
    
    
    If tipo_guardado = 1 Then
    
      sSelect = "SELECT submitted From moneyreport where idmoneyreport='" + id_moneyreport$ + "'"
      ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
      submitido = Rs(0)
      Rs.Close
      
    Else
    
      sSelect = "SELECT submitted From moneyreportbyoffice where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"
      ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
      Rs.Open sSelect, base, adOpenUnspecified
      submitido = Rs(0)
      Rs.Close
        
    End If
    
    
    
    
    
    
        
    
    If submitido = "True" And Format(fecha_de_revision$, "mm/dd/yyyy") <> Format(Now, "mm/dd/yyyy") Then
        
     
         ' CARGA TODOS no importa la fecha
        'If administrador = 0 Then
     
                 If (tipo_vista = 1) Then
                 
                    sSelect = "Select idreceipthdr, date, idcustomer, customername, policynumber, idcompany, companyname, idemployee, usr, csr, idoffice, office, fiduciary, totalreceipt, amountpaid, paymentmethod, balancedue, balanceduedate from moneyreportreceipts where idemployee='" + Format(id_employee, "####0") + "' and idoffice='" + id_oficina$ + "' and date='" + lbldate_agente.Caption + "' and void='0' and active='1'"
         
                 Else
                 
                    sSelect = "Select idreceipthdr, date, idcustomer, customername, policynumber, idcompany, companyname, idemployee, usr, csr, idoffice, office, fiduciary, totalreceipt, amountpaid, paymentmethod, balancedue, balanceduedate from moneyreportreceipts where idoffice='" + id_oficina$ + "' and date='" + lbldate_agente.Caption + "' and void='0' and active='1'"
                 
                 End If
                          
                         
    Else
     
recarga:
     
                   If (tipo_vista = 1) Then
                   
                      sSelect = "SELECT rechdr.[IdReceiptHDR],rechdr.Date,rechdr.[IdCustomer],CONCAT(cus.FirstName+' ',cus.MiddleName,+' '+cus.LastName1,+' '+cus.LastName2) as [Name] " & _
                      ",polhdr.PolicyNumber,polhdr.IdCompany,ins.CompanyName,emp.IDEmployee,emp.Username as USR,csr.Username as CSR " & _
                      ",rechdr.IdOffice,ofc.Office,rechdr.Fiduciary,rechdr.TotalAmntReceipt,rechdr.AmountPaid,PaymentMethod=STUFF " & _
                      "((SELECT DISTINCT ', ' + CAST(t3.PayMethodName AS VARCHAR(MAX)) FROM ReceiptsPayments t2 " & _
                      "join PayMethodCatalog t3 on t3.IdPayMethod=t2.IdPayMethod " & _
                      "Where t2.IdReceiptHDR = rechdr.IdReceiptHDR " & _
                      "FOR XML PATH('')),1,1,''),rechdr.BalanceDue,rechdr.BalanceDueDate " & _
                      ",rechdr.Void from [ReceiptsHDR] rechdr " & _
                      "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
                      "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
                      "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
                      "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
                      "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
                      "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
                      "where rechdr.active=1 and cast(rechdr.Date as Date) >= '" + fecha_de_revision$ + "' " & _
                      "AND cast( rechdr.DATE as Date) <= '" + fecha_de_revision$ + "' " & _
                      "and emp.IDEmployee='" + Format(id_employee, "####0") + "' and ofc.IdOffice='" + id_oficina$ + "'"    ' MODIFIQUE AQUI

                          
                         
                          
                   Else
                   
                      sSelect = "SELECT rechdr.[IdReceiptHDR],rechdr.Date,rechdr.[IdCustomer],CONCAT(cus.FirstName+' ',cus.MiddleName,+' '+cus.LastName1,+' '+cus.LastName2) as [Name] " & _
                      ",polhdr.PolicyNumber,polhdr.IdCompany,ins.CompanyName,emp.IDEmployee,emp.Username as USR,csr.Username as CSR " & _
                      ",rechdr.IdOffice,ofc.Office,rechdr.Fiduciary,rechdr.TotalAmntReceipt,rechdr.AmountPaid,PaymentMethod=STUFF " & _
                      "((SELECT DISTINCT ', ' + CAST(t3.PayMethodName AS VARCHAR(MAX)) FROM ReceiptsPayments t2 " & _
                      "join PayMethodCatalog t3 on t3.IdPayMethod=t2.IdPayMethod " & _
                      "Where t2.IdReceiptHDR = rechdr.IdReceiptHDR " & _
                      "FOR XML PATH('')),1,1,''),rechdr.BalanceDue,rechdr.BalanceDueDate " & _
                      ",rechdr.Void from [ReceiptsHDR] rechdr " & _
                      "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
                      "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
                      "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
                      "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
                      "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
                      "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
                      "where rechdr.active=1 and cast(rechdr.Date as Date) >= '" + fecha_de_revision$ + "' " & _
                      "AND cast( rechdr.DATE as Date) <= '" + fecha_de_revision$ + "' and ofc.IdOffice='" + id_oficina$ + "'"
                      
                      
                   End If
       End If
                          
                          
                         
    
                 ' carga todos los campos del registro para el administrador
        
         '                  If estado_carga = 1 Then ' Es manager
                                      
        
    
    
    
    
    
    
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid1.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid1.DataSource = Rs
                         
    Rs.Close
    
    ' acomoda grid y encabezados
    
    'grid1.cols = grid1.cols - 1
    
    contador = contador + 1
    If grid1.Rows <= 1 And contador < 5 Then
       GoTo recarga
    
    End If
    
    
    
   enca1
     
    
   For t = 1 To grid1.Rows - 1
      grid1.Row = t
      grid1.Col = 0
      grid1.Text = Format(t, "###,##0")
   Next t
   
   
   
   grid1.Row = 1
   grid1.Col = 12
   If grid1.Text <> "" And grid1.Text <> "Office" Then
     oficina_trabajada$ = grid1.Text
   Else
     oficina_trabajada$ = LTrim(RTrim(Right(cbo_oficina.List(cbo_oficina.ListIndex), 20)))
   End If
      
   
   
   
   
    Set Rs = New ADODB.Recordset
    
     
     
     
                        
                          
     If submitido = "True" And Format(fecha_de_revision$, "mm/dd/yyyy") <> Format(Now, "mm/dd/yyyy") Then
                          
                    If (tipo_vista = 1) Then
                                     
                         sSelect = "Select idreceipthdr, date, idcustomer, customername, policynumber, idcompany, companyname, idemployee, usr, csr, idoffice, office, fiduciary, totalreceipt, amountpaid, paymentmethod, balancedue, balanceduedate from moneyreportreceipts where idemployee='" + Format(id_employee, "####0") + "' and idoffice='" + id_oficina$ + "' and date='" + lbldate_agente.Caption + "' and void='1' and active='1'"
                    Else
                         sSelect = "Select idreceipthdr, date, idcustomer, customername, policynumber, idcompany, companyname, idemployee, usr, csr, idoffice, office, fiduciary, totalreceipt, amountpaid, paymentmethod, balancedue, balanceduedate from moneyreportreceipts where idoffice='" + id_oficina$ + "' and date='" + lbldate_agente.Caption + "' and void='1' and active='1'"
                    
                    End If
                    
                          
     Else
recarga2:
                   If (tipo_vista = 1) Then
     
                        sSelect = "SELECT rechdr.[IdReceiptHDR],rechdr.Date,rechdr.[IdCustomer],CONCAT(cus.FirstName+' ',cus.MiddleName,+' '+cus.LastName1,+' '+cus.LastName2) as [Name] " & _
                      ",polhdr.PolicyNumber,polhdr.IdCompany,ins.CompanyName,emp.IDEmployee,emp.Username as USR,csr.Username as CSR " & _
                      ",rechdr.IdOffice,ofc.Office,rechdr.Fiduciary,rechdr.TotalAmntReceipt,rechdr.AmountPaid,PaymentMethod=STUFF " & _
                      "((SELECT DISTINCT ', ' + CAST(t3.PayMethodName AS VARCHAR(MAX)) FROM ReceiptsPayments t2 " & _
                      "join PayMethodCatalog t3 on t3.IdPayMethod=t2.IdPayMethod " & _
                      "Where t2.IdReceiptHDR = rechdr.IdReceiptHDR " & _
                      "FOR XML PATH('')),1,1,''),rechdr.BalanceDue,rechdr.BalanceDueDate " & _
                      ",rechdr.Void from [ReceiptsHDR] rechdr " & _
                      "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
                      "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
                      "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
                      "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
                      "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
                      "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
                      "where rechdr.active=1 and rechdr.void=1 and cast(rechdr.Date as Date) >= '" + fecha_de_revision$ + "' " & _
                      "AND cast( rechdr.DATE as Date) <= '" + fecha_de_revision$ + "' and emp.IDEmployee='" + Format(id_employee, "####0") + "' and ofc.IdOffice='" + id_oficina$ + "'"
                      
                          
                   Else
                   
                          sSelect = "SELECT rechdr.[IdReceiptHDR],rechdr.Date,rechdr.[IdCustomer],CONCAT(cus.FirstName+' ',cus.MiddleName,+' '+cus.LastName1,+' '+cus.LastName2) as [Name] " & _
                      ",polhdr.PolicyNumber,polhdr.IdCompany,ins.CompanyName,emp.IDEmployee,emp.Username as USR,csr.Username as CSR " & _
                      ",rechdr.IdOffice,ofc.Office,rechdr.Fiduciary,rechdr.TotalAmntReceipt,rechdr.AmountPaid,PaymentMethod=STUFF " & _
                      "((SELECT DISTINCT ', ' + CAST(t3.PayMethodName AS VARCHAR(MAX)) FROM ReceiptsPayments t2 " & _
                      "join PayMethodCatalog t3 on t3.IdPayMethod=t2.IdPayMethod " & _
                      "Where t2.IdReceiptHDR = rechdr.IdReceiptHDR " & _
                      "FOR XML PATH('')),1,1,''),rechdr.BalanceDue,rechdr.BalanceDueDate " & _
                      ",rechdr.Void from [ReceiptsHDR] rechdr " & _
                      "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
                      "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
                      "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
                      "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
                      "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
                      "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
                      "where rechdr.active=1 and rechdr.void=1 and cast(rechdr.Date as Date) >= '" + fecha_de_revision$ + "' " & _
                      "AND cast( rechdr.DATE as Date) <= '" + fecha_de_revision$ + "' and ofc.IdOffice='" + id_oficina$ + "'"


                      
                   
                   End If

     
     End If
     
     
     
         
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid3.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid3.DataSource = Rs
                         
    Rs.Close
    
    ' acomoda grid y encabezados
    
    'grid1.cols = grid1.cols - 1
    
    contador = contador + 1
    If grid3.Rows <= 1 And contador < 5 Then
       GoTo recarga2
    
    End If
    
    
   enca3
     
    
   For t = 1 To grid3.Rows - 1
      grid3.Row = t
      grid3.Col = 0
      grid3.Text = Format(t, "###,##0")
   Next t
   
   
   
   
   
salida:
    
    'Refresh
    grid1.Visible = True
    grid3.Visible = True
    permiso_carga = 1
    
    
    For Y = 1 To grid1.Rows - 1
    
    grid1.Row = Y
    
    grid1.Col = 18
    f$ = Format(grid1.Text, "mm/dd/yyyy")
    
    If Right(f$, 4) = "1900" Then
       f$ = ""
    End If
    
    grid1.Text = f$
    
    Next Y
    
    
    
    
    For Y = 1 To grid3.Rows - 1
    
    grid3.Row = Y
    
    grid3.Col = 18
    f$ = Format(grid3.Text, "mm/dd/yyyy")
    
    If Right(f$, 4) = "1900" Then
       f$ = ""
    End If
    
    grid3.Text = f$
    
    Next Y
    
    
    
     ' obtiene el ID_moneyreport
    
    
    id_moneyreport$ = ""
    sSelect = "select idmoneyreport from moneyreport where idoffice='" + id_oficina$ + "' and idemployee='" + Format(id_employee, "###0") + "' and datereport=convert(datetime, '" + lbldate_agente.Caption + " ')"
    Rs.Open sSelect, base, adOpenUnspecified
    id_moneyreport$ = Rs(0)
    Rs.Close
    
    ' carga los archivos
    
    ListView1.ListItems.Clear
    

    ruta_archivos$ = "\\192.168.84.215\moneyreport\" + id_moneyreport$ + "\"
 


File1.Path = "c:\"
File1.Path = ruta_archivos$
 For t = 0 To File1.ListCount - 1
    
    If UCase(Right(File1.List(t), 3)) = "PDF" Then
       Set list_item = ListView1.ListItems.Add(, , File1.List(t))
       list_item.Icon = 1
       list_item.SmallIcon = 1
       list_item.SubItems(1) = File1.List(t)
    Else
       Set list_item = ListView1.ListItems.Add(, , File1.List(t))
       list_item.Icon = 2
       list_item.SmallIcon = 2
       list_item.SubItems(1) = File1.List(t)
    End If
                
         

    
 Next t
    
    
    
 If administrador = 3 Then
    
 valido1 = 777
   cbo_oficina.Clear
    
   oficina_JA$ = oficina_trabajada$
     
     If oficina_JA$ <> "" Then
       cbo_oficina.AddItem oficina_JA$ + Space(30) + oficina_JA$
     End If
      
    
    cbo_oficina.ListIndex = 0

   valido1 = 0
   carga_manager
   
   
   valido1 = 777
   
 End If
    
 msg.Visible = False
    
    
End Sub

Public Sub enca1()
On Error Resume Next


grid1.cols = grid1.cols + 1

grid1.ColWidth(0) = 600
grid1.ColAlignment(0) = flexAlignLeftCenter


grid1.ColWidth(1) = 1000 'idreceiptHDR
grid1.ColAlignment(1) = flexAlignRightCenter

grid1.ColWidth(2) = 2000   ' Date
grid1.ColAlignment(2) = flexAlignLeftCenter

grid1.ColWidth(3) = 900   ' Idcustomer
grid1.ColAlignment(3) = flexAlignCenterCenter

grid1.ColWidth(4) = 3200   ' Name
grid1.ColAlignment(4) = flexAlignLeftCenter

grid1.ColWidth(5) = 2200   'Policynumber
grid1.ColAlignment(5) = flexAlignLeftCenter

grid1.ColWidth(6) = 1200   ' idcompany
grid1.ColAlignment(6) = flexAlignCenterCenter

grid1.ColWidth(7) = 2000   ' companyname
grid1.ColAlignment(7) = flexAlignLeftCenter

grid1.ColWidth(8) = 1200   ' idemployee
grid1.ColAlignment(8) = flexAlignCenterCenter

grid1.ColWidth(9) = 1600   ' USR
grid1.ColAlignment(9) = flexAlignLeftCenter

grid1.ColWidth(10) = 1600   ' CSR
grid1.ColAlignment(10) = flexAlignLeftCenter

grid1.ColWidth(11) = 800   ' IdOffice
grid1.ColAlignment(11) = flexAlignCenterCenter



  grid1.ColWidth(12) = 1840   ' Office
  grid1.ColAlignment(12) = flexAlignLeftCenter
  grid1.ColWidth(13) = 1150  ' Fiduciary  1200
  grid1.ColAlignment(13) = flexAlignRightCenter
  
  

  
  grid1.ColWidth(14) = 1300   ' total amount receipt
  grid1.ColAlignment(14) = flexAlignRightCenter
  
  grid1.ColWidth(15) = 1300  ' Amount paid
  grid1.ColAlignment(15) = flexAlignRightCenter
  
  grid1.ColWidth(16) = 1500   ' Payment Method
  grid1.ColAlignment(16) = flexAlignRightCenter

  
    grid1.ColWidth(17) = 1300   ' balance due
  grid1.ColAlignment(17) = flexAlignRightCenter

  grid1.ColWidth(18) = 1200   ' balance due date
  grid1.ColAlignment(18) = flexAlignLeftCenter
  
  grid1.ColWidth(19) = 10   '
  grid1.ColAlignment(19) = flexAlignCenterCenter

 grid1.ColWidth(20) = 10   '
  grid1.ColAlignment(19) = flexAlignCenterCenter


grid1.Row = 0

grid1.Col = 1
grid1.Text = "Receipt#"

grid1.Col = 2
grid1.Text = "Date"

grid1.Col = 3
grid1.Text = "IdCust"

grid1.Col = 4
grid1.Text = "Name"

grid1.Col = 5
grid1.Text = "Policy#"

grid1.Col = 6
grid1.Text = "IdCompany"

grid1.Col = 7
grid1.Text = "Company"


grid1.Col = 8
grid1.Text = "IdEmployee"


grid1.Col = 9
grid1.Text = "USR"


grid1.Col = 10
grid1.Text = "CSR"

grid1.Col = 11
grid1.Text = "IdOffice"

  grid1.Col = 12
  grid1.Text = "Office"
  
  grid1.Col = 13
  grid1.Text = "Fiduciary"
  
  grid1.Col = 14
  grid1.Text = "Total Receipt"
  grid1.Col = 15
  grid1.Text = "Amount Paid"
  
  grid1.Col = 16
  grid1.Text = "PYMT Method"
  
  
  grid1.Col = 17
  grid1.Text = "Balance Due"
  grid1.Col = 18
  grid1.Text = "BD Date"
    
   grid1.Col = 19
  grid1.Text = "Void"
      
  



grid1.FixedRows = 1
grid1.FixedCols = 1

grid1.Row = 1
grid11.Col = 1

End Sub

Public Sub Desactiva_sel()
On Error Resume Next
grid.Row = 1
grid1.Col = 1
End Sub































Public Sub verifica_existencia_notas()
On Error Resume Next
    Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    Set Rs = New ADODB.Recordset

grid.Visible = False

For t = 1 To grid.Rows - 1
      grid.Row = t
      grid.Col = 11
      IDdisc$ = grid.Text
      
      
    Grid2.Clear
    sSelect = "select iddiscrepancy, iddiscrepancynotes from discrepancynotesrel where iddiscrepancy='" + IDdisc$ + "' and active=1"
  
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close

  If administrador = 0 Then
    grid.Col = 13
    If Grid2.Rows > 1 Then
       
       grid.Text = "Y"
    Else
       grid.Text = ""
    End If
    
  Else
    grid.Col = 17
    If Grid2.Rows > 1 Then
       
       grid.Text = "Y"
    Else
       grid.Text = ""
    End If
  
  End If
    
Next t

grid.Visible = True


End Sub

Public Sub enca_printer()
On Error Resume Next

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
           
           
    num = Val(Right(cbo_managers.List(cbo_managers.ListIndex), 5))
           
    sSelect = "select firstname from employeeinfo where idemployee='" + Format(num, "###0") + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    nombre$ = Rs(0)
    Rs.Close
           
    sSelect = "select lastname1 from employeeinfo where idemployee='" + Format(num, "###0") + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    apellido$ = Rs(0)
    Rs.Close
           
           
           
Printer.FontName = "Courier new"

Printer.Orientation = 2 ' 1=portrait   2=landscape

Printer.FontSize = 18
Printer.Print Space(1)
Printer.Print Space(5) + "Office: " + Left(UCase(cbo_oficina.List(cbo_oficina.ListIndex)), 20)
Printer.FontSize = 12
Printer.Print Space(7) + "Manager: " + nombre$ + Space(1) + apellido$
Printer.Print Space(1)

Printer.FontSize = 8
Printer.Print Space(5) + "   DATE   " + Space(2) + "C O M P A N Y       " + Space(1) + "POLICY No.          " + Space(1) + "C U S T O M E R               " + "CUST-ID";
Printer.Print Space(3) + "TYPE-TRANSACTION    " + Space(3) + "  AMOUNT   " + Space(1) + "A G E N T         " + Space(1) + "STATUS"


Printer.Print Space(5) + "-----------------------------------------------------------------------------------------------------------------------------------------------------------"

Printer.Print Space(1)

' Printer.Print Space(1) + "XXXX-XX-XX" + Space(2) + "XXXXXXXXXXXXXXXXXXXX" + Space(1) + "XXXXXXXXXXXXXXXXXXXX" + Space(1) + "XXXXXXXXXXXXXXXXXXXXXXXXX" + Space(2) + "XXXXX";
' Printer.Print Space(1) + "XXXXXXXXXXXXXXXXXXXX" + Space(1) + "$XXX,XXX.00" + Space(1) + "XXXXXXXXXXXXXXXXXXXX"





End Sub




Public Sub enca_printer2()
On Error Resume Next

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
           
           
    num = Val(Right(cbo_managers.List(cbo_managers.ListIndex), 5))
           
    sSelect = "select firstname from employeeinfo where idemployee='" + Format(num, "###0") + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    nombre$ = Rs(0)
    Rs.Close
           
    sSelect = "select lastname1 from employeeinfo where idemployee='" + Format(num, "###0") + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    apellido$ = Rs(0)
    Rs.Close
           
           
           
Printer.FontName = "Courier new"

Printer.Orientation = 2 ' 1=portrait   2=landscape

Printer.FontSize = 18
Printer.Print Space(1)

If btnmostrar_todo.Value = False Then
    Printer.Print Space(5) + "Office: " + Left(UCase(cbo_oficina.List(cbo_oficina.ListIndex)), 20)
Else
    Printer.Print Space(5) + "Office: ALL"
End If

Printer.FontSize = 12
Printer.Print Space(7) + "Manager: "
Printer.Print Space(1)

Printer.FontSize = 7
Printer.Print Space(5) + "   DATE   " + Space(2) + "C O M P A N Y       " + Space(1) + "POLICY No.          " + Space(1) + "C U S T O M E R               " + "CUST-ID";
Printer.Print Space(3) + "TYPE-TRANSACTION    " + Space(3) + "  AMOUNT   " + Space(1) + "A G E N T         " + Space(1) + "DRAFT DATE" + Space(1) + "   COLLECTED  " + Space(1) + "STATUS"


Printer.Print Space(5) + "--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

Printer.Print Space(1)

' Printer.Print Space(1) + "XXXX-XX-XX" + Space(2) + "XXXXXXXXXXXXXXXXXXXX" + Space(1) + "XXXXXXXXXXXXXXXXXXXX" + Space(1) + "XXXXXXXXXXXXXXXXXXXXXXXXX" + Space(2) + "XXXXX";
' Printer.Print Space(1) + "XXXXXXXXXXXXXXXXXXXX" + Space(1) + "$XXX,XXX.00" + Space(1) + "XXXXXXXXXXXXXXXXXXXX"



End Sub

Public Sub carga_oficinas()
On Error Resume Next



Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
           
  
 
  Grid2.Clear
    
  
  
  
    sSelect = "select idemployee from employeeinfo where username='" + user$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_employee = Rs(0)
    Rs.Close
    
    
If administrador = 0 Then
  
  If chkagentes(1).Value = 0 Then
  
  
  
   sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (3,6,16,17,18, 28,2,24,37) and Username='" + user$ + "'" ' and empjob.Active='1'





  
  Else
  

  
  
   sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where empofc.active=1 and IdJobTitle in (3,6,16,17,18,28,2,24,37) and Username='" + user$ + "'" ' and empjob.Active='1'
  
  
  

  
  End If
  
Else


  If chkagentes(1).Value = 0 Then
  
  
  
   sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (3,6,16,17,18, 28,2,24,37)"





  
  Else
  

  
  
   sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where empofc.active=1 and IdJobTitle in (3,6,16,17,18,28,2,24,37)"
  
  
  

  
  End If



End If
  
  
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    
    
    
'If administrador = 0 Then
    cbo_oficina.Clear
    
    For t = 1 To Grid2.Rows - 1
     
     Grid2.Row = t
     Grid2.Col = 4
     titulo = Val(Grid2.Text)
     
     Grid2.Col = 3
     oficina_JA$ = Grid2.Text
     
     
     
     existe = 0
     For Y = 0 To cbo_oficina.ListCount - 1
        office$ = LTrim(RTrim(Right(cbo_oficina.List(Y), 20)))
        If UCase(office$) = UCase(oficina_JA$) Then
             existe = 1
             Exit For
        End If
     Next Y
     
     
     
     If existe = 0 Then
     
      If UCase(oficina_JA$) <> "" Then
       cbo_oficina.AddItem UCase(oficina_JA$) + Space(30) + UCase(oficina_JA$)
      End If
       
       
     End If
     
    Next t
    cbo_oficina.ListIndex = 0
'Else

    
'End If





End Sub

Public Sub carga_archivos()




Dim column_header As ColumnHeader
Dim list_item As ListItem

msg.Visible = True



    ' Create the column headers.
    Set column_header = ListView1. _
        ColumnHeaders.Add(, , "Abbrev", _
        TextWidth("Abbrev"))
    Set column_header = ListView1. _
        ColumnHeaders.Add(, , "Title", _
        TextWidth("Ready-to-Run Visual Basic Algorithms"))
    Set column_header = ListView1. _
        ColumnHeaders.Add(, , "ISBN", _
        TextWidth("0-000-00000-0"))

    ' Start with report view.
    ' mnuViewChoice_Click lvwReport
     ListView1.View = 0

    ' Associate the ImageLists with the
    ' ListView's Icons and SmallIcons properties.
    ListView1.Icons = imgLarge
    ListView1.SmallIcons = imgSmall
    
    
    Exit Sub
    

    Set list_item = ListView1.ListItems.Add(, , "VBA")
    list_item.Icon = 2
    list_item.SmallIcon = 1
    list_item.SubItems(1) = "Ready-to-Run Visual Basic Algorithms"
    list_item.SubItems(2) = "0-471-24268-3"

    Set list_item = ListView1.ListItems.Add(, , "VBGP")
    list_item.Icon = 1
    list_item.SmallIcon = 1
    list_item.SubItems(1) = "Visual Basic Graphics Programming"
    list_item.SubItems(2) = "0-471-15533-0"

    Set list_item = ListView1.ListItems.Add(, , "CCL")
    list_item.Icon = 1
    list_item.SmallIcon = 1
    list_item.SubItems(1) = "Custom Controls Library"
    list_item.SubItems(2) = "0-471-24267-5"

    Set list_item = ListView1.ListItems.Add(, , "AVBT")
    list_item.Icon = 1
    list_item.SmallIcon = 1
    list_item.SubItems(1) = "Advanced Visual Basic Techniques"
    list_item.SubItems(2) = "0-471-18881-6"
    
    
     msg.Visible = False

End Sub





Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

base.Close



  r$ = Shell("c:\money\reset.exe", vbNormalFocus)

  r$ = Shell("c:\iconos\barra_agent.exe", vbNormalFocus)



X$ = Shell("cmd /c taskkill /f /im money.exe")
End
End Sub

Private Sub Image12_Click()
On Error Resume Next
cbo_oficina.Enabled = True
End Sub

Private Sub img_tab1_Click(Index As Integer)
On Error Resume Next
 hoja1.Visible = False
 Hoja2.Visible = False
 hoja3.Visible = False
 panel.Visible = False
 Refresh

valido1 = 0

Dim sSelect As String
    
Dim Rs As ADODB.Recordset

Set Rs = New ADODB.Recordset
           



  id_employee = Val(Right(Form1.cbo_agentes.List(Form1.cbo_agentes.ListIndex), 20))
  ID_manager = Val(Right(Form1.cbo_managers.List(Form1.cbo_managers.ListIndex), 20))
  lae_office$ = RTrim(LTrim(Right(Form1.cbo_oficina.List(Form1.cbo_oficina.ListIndex), 25)))
 
 sSelect = "select idoffice from officescatalog where office='" + lae_office$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_oficina$ = Rs(0)
    Rs.Close
 


cargo_impresion = 0

If Index = 0 Then   'And administrador = 0 Then
     tabx(0).FillColor = &HFF&
     tipo_guardado = 1
 
     lblleyenda.Caption = "Agent Report Form"

     If Val(cargo$) = 17 Or Val(cargo$) = 24 Then
         tabx(1).FillColor = &HC0&
         tabx(2).FillColor = &HC0&
         leyenda(2).ForeColor = &H8080FF
         leyenda(3).ForeColor = &H8080FF
 
     Else
         tabx(1).FillColor = &HC0&
         tabx(2).FillColor = &H808080
         leyenda(2).ForeColor = &HC0C0C0
         leyenda(3).ForeColor = &HC0C0C0
     End If
  
     leyenda(0).ForeColor = &HC0C0FF
     leyenda(1).ForeColor = &H8080FF
 
     hoja1.Visible = True
     hoja1.Refresh
 
     chk_dayoff.Visible = True
     msgdescanso.Visible = chk_dayoff.Value
     msgdescanso.Refresh


 
ElseIf Index = 1 Then  ' And administrador = 0 Then

     cargo_impresion = 3

     If tipo_guardado = 2 Then
         panel.Visible = True
     End If

     lblleyenda.Caption = "Receipts/reports"
     chk_dayoff.Visible = False
     msgdescanso.Visible = False
      msgdescanso.Refresh
    

     btncarga_datos_agente_Click
     tabx(1).FillColor = &HFF&
  
     leyenda(1).ForeColor = &HC0C0FF
     leyenda(0).ForeColor = &H8080FF
  
     If Val(cargo$) = 17 Or Val(cargo$) = 24 Then
         tabx(0).FillColor = &HC0&
         tabx(2).FillColor = &HC0&
         leyenda(2).ForeColor = &H8080FF
         leyenda(3).ForeColor = &H8080FF
       
     Else
         tabx(0).FillColor = &HC0&
         tabx(2).FillColor = &H808080
         leyenda(2).ForeColor = &HC0C0C0
         leyenda(3).ForeColor = &HC0C0C0
     End If
 
     tabx(0).Refresh
     tabx(1).Refresh
     tabx(2).Refresh
     leyenda(0).Refresh
     leyenda(2).Refresh
     leyenda(3).Refresh
  
     'hoja1.Visible = False
     Hoja2.Visible = True
     Hoja2.Refresh
     'hoja3.Visible = False
 
ElseIf (Index = 2 And (Val(cargo$) = 17 Or Val(cargo$) = 24)) Or administrador = 1 Then


     If tipo_guardado = 2 Then
         panel.Visible = True
     End If

     lblleyenda.Caption = "Manager Report Form"
     chk_dayoff.Visible = False
     msgdescanso.Visible = False
      msgdescanso.Refresh


     cargo_impresion = 1
     tipo_guardado = 2
     tabx(2).FillColor = &HFF&
 
     tabx(0).FillColor = &HC0&
     tabx(1).FillColor = &HC0&
 
     leyenda(2).ForeColor = &HC0C0FF
     leyenda(3).ForeColor = &HC0C0FF
 
     leyenda(0).ForeColor = &H8080FF
     leyenda(1).ForeColor = &H8080FF
 
     tabx(0).Refresh
     tabx(1).Refresh
     tabx(2).Refresh
     leyenda(0).Refresh
     leyenda(1).Refresh
     leyenda(2).Refresh
     leyenda(3).Refresh
   
     'hoja1.Visible = False
     'Hoja2.Visible = False
     hoja3.Visible = True
     hoja3.Refresh
 
ElseIf (Index = 2 And Val(cargo$) = 16) Then
 
     tabx(0).FillColor = &HFF&
     chk_dayoff.Visible = False
     msgdescanso.Visible = False
      msgdescanso.Refresh

   
     tabx(1).FillColor = &HC0&
     tabx(2).FillColor = &H808080
     leyenda(2).ForeColor = &HC0C0C0
     leyenda(3).ForeColor = &HC0C0C0
 
     leyenda(0).ForeColor = &HC0C0FF
     leyenda(1).ForeColor = &H8080FF

     hoja1.Visible = True
     hoja1.Refresh
     ' Hoja2.Visible = False
     ' hoja3.Visible = False
 
End If








If tipo_guardado = 1 Then
          
     ' verifica si ya existe el reporte
     
     sSelect = "select idmoneyreport from moneyreport where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + Format(id_employee, "###0") + "' and idoffice='" + id_oficina$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreport$ = Rs(0)
     Rs.Close
        
     lblnum.Caption = id_moneyreport$
    
Else
      
    
     sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idoffice='" + id_oficina$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreportoffice$ = Rs(0)
     Rs.Close
    
     lblnum.Caption = "O-" + id_moneyreportoffice$
     
End If
  
  
  
btnupdate_LAE_Click

carga_totales_oficina
'Refresh
obtener_fecha_real

carga_revisado
 
 
 
 
End Sub

Private Sub ListView1_Click()
On Error Resume Next
archivo_selecto$ = ListView1.SelectedItem

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset

If archivo_selecto$ = "" Then Exit Sub



 id_employee = Val(Right(Form1.cbo_agentes.List(Form1.cbo_agentes.ListIndex), 20))

oficina$ = LTrim(Right(UCase(RTrim(cbo_oficina.List(cbo_oficina.ListIndex))), 30))
    sSelect = "SELECT idoffice From officescatalog where office='" + oficina$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    id_office$ = Rs(0)
    Rs.Close
    

 id_moneyreport$ = ""

 sSelect = "select idmoneyreport from moneyreport where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idoffice='" + id_office$ + "' and idemployee='" + Format(id_employee, "###0") + "'"
 Rs.Open sSelect, base, adOpenUnspecified
 id_moneyreport$ = Rs(0)
 Rs.Close


If id_moneyreport$ = "" Then Exit Sub

 ruta_archivos$ = "\\192.168.84.215\moneyreport\" + id_moneyreport$ + "\"




visualizador.Visible = True

If Right(UCase(archivo_selecto$), 3) = "PDF" Then
   If Dir$("c:\money\" + archivo_selecto$) <> "" Then
      pdf1.src = "c:\money\" + archivo_selecto$
      pdf1.gotoFirstPage
   Else
      pdf1.src = ruta_archivos$ + archivo_selecto$
      pdf1.gotoFirstPage
    
   End If
     
    pdf1.Visible = True
    img1.Visible = False
Else
    If Dir$("c:\money\" + archivo_selecto$) <> "" Then
       img1.Picture = LoadPicture("c:\money\" + archivo_selecto$)
    Else
       img1.Picture = LoadPicture(ruta_archivos$ + archivo_selecto$)
    End If
    
    
    pdf1.Visible = False
    img1.Visible = True
End If


'strDir = "c:\money"
'strFile = archivo_selecto$
'ShellExecute 0, "OPEN", strDir & strFile, "", strDir, 1

End Sub

Private Sub ListView1_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error Resume Next

 ' See if we know what to do with the data.

    If Data.GetFormat(vbCFBitmap) Then

        ' Copy the bitmap.

       ' picDragTo(Index).Picture = Data.GetData(vbCFBitmap)

        Effect = vbDropEffectCopy

    ElseIf Data.GetFormat(vbCFFiles) Then

        ' See if this is a file name ending in

        ' bmp, gif, jpg, pdf or jpeg.

        extension1$ = LCase$(Right$(Data.Files(1), 4))
        
        guarda$ = ""
        For Y = Len(Data.Files(1)) To 1 Step -1
           If Mid$(Data.Files(1), Y, 1) = "\" Then
               Exit For
           Else
              guarda$ = Mid$(Data.Files(1), Y, 1) + guarda$
           End If
        Next Y
        
        nombre_archivo$ = guarda$
        Kill "c:\money\" + nombre_archivo$
        
        Select Case extension1$
        Case ".pdf"
        
                If Dir$("c:\money\" + guarda$) <> "" Then
                   GoTo brinca
                End If
                
                
                For z = 1 To ListView1.ListItems.Count
                                 
                   a$ = ListView1.ListItems.Item(z)
                   If a$ = guarda$ Then
                       ListView1.ListItems.Remove (z)
                       Exit For
                   End If
                
                Next z
                
                Set list_item = ListView1.ListItems.Add(, , guarda$)
                list_item.Icon = 1
                list_item.SmallIcon = 1
                list_item.SubItems(1) = Data.Files(1)
                'list_item.SubItems(2) = "0-471-24267-5"
                
            Effect = vbDropEffectCopy
            
            FileCopy Data.Files(1), "c:\money\" + guarda$
            


        Case ".bmp", ".png", ".jpg", "jpeg"

            ' Load the file.

        '    picDragTo(Index).Picture = LoadPicture(Data.Files(1))
        
                If Dir$("c:\money\" + guarda$) <> "" Then
                   GoTo brinca
                End If
    
        
           Set list_item = ListView1.ListItems.Add(, , guarda$)
                list_item.Icon = 2
                list_item.SmallIcon = 2
                list_item.SubItems(1) = Data.Files(1)
                'list_item.SubItems(2) = "0-471-24267-5"
                
         

            Effect = vbDropEffectCopy
            
            FileCopy Data.Files(1), "c:\money\" + guarda$
         
            
       
        Case Else

            ' Tell the source we did nothing.

            Effect = vbDropEffectNone

        End Select

    Else

        ' Tell the source we did nothing.

        Effect = vbDropEffectNone

    End If
    
    
brinca:
modificado = 1

End Sub


Public Sub arrastra_archivo(nombre_archivo As String)
On Error Resume Next

 ' See if we know what to do with the data.

   
         
   
        ' See if this is a file name ending in

        ' bmp, gif, jpg, pdf or jpeg.

        extension1$ = LCase$(Right$(nombre_archivo, 4))
        
        guarda$ = ""
        For Y = Len(nombre_archivo) To 1 Step -1
           If Mid$(nombre_archivo, Y, 1) = "\" Then
               Exit For
           Else
              guarda$ = Mid$(nombre_archivo, Y, 1) + guarda$
           End If
        Next Y
        
        nombre_archivo$ = guarda$
        
       Kill "c:\money\" + nombre_archivo$
        
        
        Select Case extension1$
        Case ".pdf"
        
                If Dir$("c:\money\" + guarda$) <> "" Then
                   GoTo brinca
                End If
    
                For z = 1 To ListView1.ListItems.Count
                                 
                   a$ = ListView1.ListItems.Item(z)
                   If a$ = guarda$ Then
                       ListView1.ListItems.Remove (z)
                       Exit For
                   End If
                
                Next z
                
                
                Set list_item = ListView1.ListItems.Add(, , guarda$)
                list_item.Icon = 1
                list_item.SmallIcon = 1
                list_item.SubItems(1) = nombre_archivo
                'list_item.SubItems(2) = "0-471-24267-5"
                
            Effect = vbDropEffectCopy
            
            FileCopy nombre_archivo, "c:\money\" + guarda$
            
            
        Case ".bmp", ".png", ".jpg", "jpeg"

            ' Load the file.

        '    picDragTo(Index).Picture = LoadPicture(Data.Files(1))
        
                If Dir$("c:\money\" + guarda$) <> "" Then
                   GoTo brinca
                End If
    
        
           Set list_item = ListView1.ListItems.Add(, , guarda$)
                list_item.Icon = 2
                list_item.SmallIcon = 2
                list_item.SubItems(1) = nombre_archivo
                'list_item.SubItems(2) = "0-471-24267-5"
                
         

            Effect = vbDropEffectCopy
            
            FileCopy nombre_archivo, "c:\money\" + guarda$
         
            
       
        Case Else

            ' Tell the source we did nothing.

            Effect = vbDropEffectNone

        End Select
    


    
brinca:

End Sub






























Private Sub txtagente_venido_Click(Index As Integer)

End Sub

Private Sub ListView2_Click()
On Error Resume Next
archivo_selecto$ = ListView2.SelectedItem



Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    
    
 If archivo_selecto$ = "" Then Exit Sub
 

 id_employee = Val(Right(Form1.cbo_agentes.List(Form1.cbo_agentes.ListIndex), 20))

oficina$ = LTrim(Right(UCase(RTrim(cbo_oficina.List(cbo_oficina.ListIndex))), 30))
    sSelect = "SELECT idoffice From officescatalog where office='" + oficina$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    id_office$ = Rs(0)
    Rs.Close
    


 sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idoffice='" + id_office$ + "'"
 Rs.Open sSelect, base, adOpenUnspecified
 id_moneyreportoffice$ = Rs(0)
 Rs.Close


If id_moneyreportoffice$ = "" Then Exit Sub


 ruta_archivos$ = "\\192.168.84.215\moneyreport\O-" + id_moneyreportoffice$ + "\"



 

visualizador2.Visible = True
If Right(UCase(archivo_selecto$), 3) = "PDF" Then

     
  If Dir$("c:\money\" + archivo_selecto$) <> "" Then
     pdf2.src = "c:\money\" + archivo_selecto$
     pdf2.gotoFirstPage
  Else
     pdf2.src = ruta_archivos$ + archivo_selecto$
     pdf2.gotoFirstPage
    
  End If
     
    pdf2.Visible = True
    img2.Visible = False
Else
    If Dir$("c:\money\" + archivo_selecto$) <> "" Then
       img2.Picture = LoadPicture("c:\money\" + archivo_selecto$)
    Else
       img2.Picture = LoadPicture(ruta_archivos$ + archivo_selecto$)
    End If
       
       
    pdf2.Visible = False
    img2.Visible = True
End If
End Sub


Private Sub ListView2_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

 ' See if we know what to do with the data.

    If Data.GetFormat(vbCFBitmap) Then

        ' Copy the bitmap.

       ' picDragTo(Index).Picture = Data.GetData(vbCFBitmap)

        Effect = vbDropEffectCopy

    ElseIf Data.GetFormat(vbCFFiles) Then

        ' See if this is a file name ending in

        ' bmp, gif, jpg, pdf or jpeg.

        extension1$ = LCase$(Right$(Data.Files(1), 4))
        
        guarda$ = ""
        For Y = Len(Data.Files(1)) To 1 Step -1
           If Mid$(Data.Files(1), Y, 1) = "\" Then
               Exit For
           Else
              guarda$ = Mid$(Data.Files(1), Y, 1) + guarda$
           End If
        Next Y
        
        nombre_archivo$ = guarda$
         Kill "c:\money\" + nombre_archivo$
        
        Select Case extension1$
        Case ".pdf"
        
                If Dir$("c:\money\" + guarda$) <> "" Then
                   GoTo brinca
                End If
    
                
                For z = 1 To ListView2.ListItems.Count
                                 
                   a$ = ListView2.ListItems.Item(z)
                   If a$ = guarda$ Then
                       ListView2.ListItems.Remove (z)
                       Exit For
                   End If
                
                Next z
                
                
                Set list_item = ListView2.ListItems.Add(, , guarda$)
                list_item.Icon = 1
                list_item.SmallIcon = 1
                list_item.SubItems(1) = Data.Files(1)
                'list_item.SubItems(2) = "0-471-24267-5"
                
            Effect = vbDropEffectCopy
            
            FileCopy Data.Files(1), "c:\money\" + guarda$
            


        Case ".bmp", ".png", ".jpg", "jpeg"

            ' Load the file.

        '    picDragTo(Index).Picture = LoadPicture(Data.Files(1))
        
                If Dir$("c:\money\" + guarda$) <> "" Then
                   GoTo brinca
                End If
    
        
           Set list_item = ListView2.ListItems.Add(, , guarda$)
                list_item.Icon = 2
                list_item.SmallIcon = 2
                list_item.SubItems(1) = Data.Files(1)
                'list_item.SubItems(2) = "0-471-24267-5"
                
         

            Effect = vbDropEffectCopy
            
            FileCopy Data.Files(1), "c:\money\" + guarda$
         
            
       
        Case Else

            ' Tell the source we did nothing.

            Effect = vbDropEffectNone

        End Select

    Else

        ' Tell the source we did nothing.

        Effect = vbDropEffectNone

    End If
    
    
brinca:
modificado = 1

End Sub





Private Sub op_cards_Click(Index As Integer)
On Error Resume Next

tabcard(0).Visible = False
tabcard(1).Visible = False

tabcard(Index).Visible = True

If Index = 0 Then
  txtcredit_agente(0).SetFocus
Else
  txtcredit_agente(10).SetFocus
End If

End Sub

Private Sub op_view_Click(Index As Integer)
On Error Resume Next
tipo_vista = Index + 1
btncarga_datos_agente_Click
End Sub



Private Sub Timer1_Timer()
On Error Resume Next

seg = seg + 1
If seg = 1 Then
  Picture2.Visible = True
  Picture2.BackColor = &HFFFF&
  mensaje.ForeColor = &HFF&
  
ElseIf seg >= 2 Then
  Picture2.BackColor = &HFF&
  mensaje.ForeColor = &HFFFF&
  seg = 0


Else


End If

End Sub

Private Sub Timer2_Timer()
On Error Resume Next

segundos = segundos + 1



If segundos = 60 Then
  If Dir$("c:\iconos\barra_agent.exe") = "" Then
     FileCopy "\\192.168.84.215\moneyreport\barra_agent.exe", "c:\iconos\barra_agent.exe"
  End If


End If



If segundos >= 3600 Then

  

  fecha_hoy$ = Format(Now, "mm/dd/yyyy")
  If fecha_hoy$ <> fecha_entrada$ Then
    End
  End If
  segundos = 0
End If





End Sub

Private Sub txtamount_agente_Change(Index As Integer)
On Error Resume Next
total_voids = 0
For t = 0 To 1
  total_voids = total_voids + Val(txtamount_agente(t).Text)
Next t

lbltotal_recibos_void_agente.Caption = Format(total_voids, "$###,##0.00")
lbltotal_void_agente.Caption = Format(total_voids, "$###,##0.00")
lbltotal_reportado_menos_voids_agente.Caption = Format(total_reportado - total_voids, "$###,##0.00")

total_dinero = Val(Format(lblgrantotal_debitcredit_agente.Caption, "000000.00")) + Val(Format(lblgrantotal_cash_agente.Caption, "000000.00"))
lbltotal_over_short_agente.Caption = Format(total_dinero - (total_reportado - total_voids), "$###,##0.00")

modificado = 1
End Sub

Private Sub txtamount_agente_GotFocus(Index As Integer)
On Error Resume Next
If Val(txtamount_agente(Index).Text) = 0 Then txtamount_agente(Index).Text = ""

End Sub


Private Sub txtamount_agente_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
    If Index = 1 Then
       txttotal_LAE_agente.SetFocus
       
    Else
      txtamount_agente(Index + 1).SetFocus
    End If
End If

If KeyAscii = Asc(".") Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If


End Sub


Private Sub txtamount_agente_LostFocus(Index As Integer)
On Error Resume Next


pos = InStr(1, txtamount_agente(Index).Text, ".")

If pos > 0 Then
  txtamount_agente(Index).Text = Left(txtamount_agente(Index).Text, pos + 2)
  If Mid(txtamount_agente(Index).Text, pos + 1, 1) = "." Then
     
  Else
    If Mid(txtamount_agente(Index).Text, pos + 2, 1) = "." Then
       txtamount_agente(Index).Text = Left(txtamount_agente(Index).Text, Len(txtamount_agente(Index).Text) - 1)
    End If
  End If
  
  
  
  If Val(txtamount_agente(Index).Text) = 0 Then
      txtamount_agente(Index).Text = "0"
  End If
End If

txtamount_agente(Index).Text = Format(txtamount_agente(Index).Text, "#####0.00")

End Sub


Private Sub txtcant_ida_Change(Index As Integer)
On Error Resume Next
If valido1 = 66 Then Exit Sub
total_dinero_agent_went = 0
For t = 0 To 2
  total_dinero_agent_went = total_dinero_agent_went + Val(txtcant_ida(t).Text)
Next t

lbltotal_agentes_idos.Caption = Format(total_dinero_agent_went, "$###,##0.00")


''total_dinero = Val(Format(lblgrantotal_debitcredit_agente.Caption, "000000.00")) + Val(Format(lblgrantotal_cash_agente.Caption, "000000.00"))
''lbltotal_over_short_agente.Caption = Format((total_reportado - total_voids) - total_dinero, "$###,##0.00")

modificado = 1

calcula_total_oficina_manager
End Sub

Private Sub txtcant_ida_GotFocus(Index As Integer)
On Error Resume Next
If Val(txtcant_ida(Index).Text) = 0 Then txtcant_ida(Index).Text = ""

End Sub

Private Sub txtcant_ida_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
    If Index = 2 Then
       txtcant_venida(0).SetFocus
       
    Else
      txtcant_ida(Index + 1).SetFocus
    End If
End If

If KeyAscii = Asc(".") Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If

End Sub


Private Sub txtcant_venida_Change(Index As Integer)
On Error Resume Next
If valido1 = 66 Then Exit Sub
total_dinero_agent_came = 0
For t = 0 To 2
  total_dinero_agent_came = total_dinero_agent_came + Val(txtcant_venida(t).Text)
Next t

lbltotal_agentes_que_vinieron.Caption = Format(total_dinero_agent_came, "$###,##0.00")

calcula_total_oficina_manager

''total_dinero = Val(Format(lblgrantotal_debitcredit_agente.Caption, "000000.00")) + Val(Format(lblgrantotal_cash_agente.Caption, "000000.00"))
''lbltotal_over_short_agente.Caption = Format((total_reportado - total_voids) - total_dinero, "$###,##0.00")

modificado = 1
End Sub

Private Sub txtcant_venida_GotFocus(Index As Integer)
On Error Resume Next
If Val(txtcant_venida(Index).Text) = 0 Then txtcant_venida(Index).Text = ""

End Sub


Private Sub txtcant_venida_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
    If Index = 2 Then
       txtdinero(0).SetFocus
       
    Else
      txtcant_venida(Index + 1).SetFocus
    End If
End If

If KeyAscii = Asc(".") Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If
End Sub

Private Sub txtcash_Change(Index As Integer)
On Error Resume Next
If valido1 = 66 Then Exit Sub
total_cash = 0
For t = 0 To 5
  total_cash = total_cash + Val(txtcash(t).Text)
Next t

lbltotal_cash_agente.Caption = Format(total_cash, "$###,##0.00")
lblgrantotal_cash_agente.Caption = Format(total_cash, "$###,##0.00")

total_dinero = Val(Format(lblgrantotal_debitcredit_agente.Caption, "000000.00")) + Val(Format(lblgrantotal_cash_agente.Caption, "000000.00"))
lbltotal_over_short_agente.Caption = Format(total_dinero - (total_reportado - total_voids), "$###,##0.00")

modificado = 1
End Sub

Private Sub txtcash_GotFocus(Index As Integer)
On Error Resume Next
If Val(txtcash(Index).Text) = 0 Then txtcash(Index).Text = ""

End Sub


Private Sub txtcash_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
    If Index = 5 Then
       txtdebit_agente(0).SetFocus
       
    Else
      txtcash(Index + 1).SetFocus
    End If
End If

If KeyAscii = Asc(".") Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If


End Sub


Private Sub txtcash_LostFocus(Index As Integer)
On Error Resume Next


pos = InStr(1, txtcash(Index).Text, ".")

If pos > 0 Then
  txtcash(Index).Text = Left(txtcash(Index).Text, pos + 2)
  If Mid(txtcash(Index).Text, pos + 1, 1) = "." Then
     
  Else
    If Mid(txtcash(Index).Text, pos + 2, 1) = "." Then
       txtcash(Index).Text = Left(txtcash(Index).Text, Len(txtcash(Index).Text) - 1)
    End If
  End If
  
  
  
  If Val(txtcash(Index).Text) = 0 Then
      txtcash(Index).Text = "0"
  End If
End If

txtcash(Index).Text = Format(txtcash(Index).Text, "#####0.00")

End Sub


Private Sub txtcredit_agente_Change(Index As Integer)
On Error Resume Next
total_credit = 0
For t = 0 To 19
  total_credit = total_credit + Val(txtcredit_agente(t).Text)
 Next t

lbltotal_credit_agent.Caption = Format(total_credit, "$###,##0.00")

lbltotal_debit_credit_agent.Caption = Format(total_debit + total_credit, "$###,##0.00")
lblgrantotal_debitcredit_agente.Caption = Format(total_debit + total_credit, "$###,##0.00")


total_dinero = Val(Format(lblgrantotal_debitcredit_agente.Caption, "000000.00")) + Val(Format(lblgrantotal_cash_agente.Caption, "000000.00"))
lbltotal_over_short_agente.Caption = Format(total_dinero - (total_reportado - total_voids), "$###,##0.00")

modificado = 1

End Sub

Private Sub txtcredit_agente_GotFocus(Index As Integer)
On Error Resume Next
If Val(txtcredit_agente(Index).Text) = 0 Then txtcredit_agente(Index).Text = ""
End Sub

Private Sub txtcredit_agente_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
    If Index = 19 Then
       txtcustomer_agente(0).SetFocus
       
    ElseIf Index < 10 Then
      txtcredit_agente(Index + 1).SetFocus
      tabcard(0).Visible = True
      tabcard(1).Visible = False
      
      
    ElseIf Index > 9 And Index < 19 Then
      txtcredit_agente(Index + 1).SetFocus
      tabcard(1).Visible = True
      tabcard(0).Visible = False
      
      
    End If
End If

If KeyAscii = Asc(".") Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If


End Sub


Private Sub txtcredit_agente_LostFocus(Index As Integer)
On Error Resume Next

img_credit(Index).Picture = LoadPicture()

For t = 1 To grid1.Rows - 1
   grid1.Row = t
   
   grid1.Col = 15
   cantidad_pagada$ = grid1.Text
   
   
   grid1.Col = 16
   metodo$ = grid1.Text
   
   
   
   If Val(txtcredit_agente(Index).Text) = Val(cantidad_pagada$) Then
      
      
      Select Case UCase(RTrim(LTrim(metodo$)))

      Case "VISA"
           img_credit(Index).Picture = visa.Picture
      Case "MASTERCARD"
        img_credit(Index).Picture = master.Picture
      Case "AMEX"
      img_credit(Index).Picture = american.Picture
      Case "DISCOVER"
      img_credit(Index).Picture = discover.Picture
      Case "DEBIT"
      img_credit(Index).Picture = debito.Picture
     
      End Select
   End If
   
Next t





pos = InStr(1, txtcredit_agente(Index).Text, ".")

If pos > 0 Then
  txtcredit_agente(Index).Text = Left(txtcredit_agente(Index).Text, pos + 2)
  If Mid(txtcredit_agente(Index).Text, pos + 1, 1) = "." Then
     
  Else
    If Mid(txtcredit_agente(Index).Text, pos + 2, 1) = "." Then
       txtcredit_agente(Index).Text = Left(txtcredit_agente(Index).Text, Len(txtcredit_agente(Index).Text) - 1)
    End If
  End If
  
  
  
  If Val(txtcredit_agente(Index).Text) = 0 Then
      txtcredit_agente(Index).Text = "0"
  End If
End If

txtcredit_agente(Index).Text = Format(txtcredit_agente(Index).Text, "#####0.00")


End Sub

Private Sub txtcredit_manager_Change()
On Error Resume Next
If valido1 = 66 Then Exit Sub
total_credito_oficina = 0

total_credito_oficina = Val(txtdebit_manager.Text) + Val(txtcredit_manager.Text)


lbltotal_debito_credito_manager.Caption = Format(total_credito_oficina, "$###,##0.00")

calcula_total_oficina_manager
''total_dinero = Val(Format(lblgrantotal_debitcredit_agente.Caption, "000000.00")) + Val(Format(lblgrantotal_cash_agente.Caption, "000000.00"))
''lbltotal_over_short_agente.Caption = Format((total_reportado - total_voids) - total_dinero, "$###,##0.00")

modificado = 1
End Sub

Private Sub txtcredit_manager_GotFocus()
On Error Resume Next
If Val(txtcredit_manager.Text) = 0 Then txtcredit_manager.Text = ""
End Sub

Private Sub txtcredit_manager_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
    
       cbo_employees(0).SetFocus
       Exit Sub
 
End If


If KeyAscii = Asc(".") Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If

End Sub


Private Sub txtcustomer_agente_Change(Index As Integer)
modificado = 1
End Sub

Private Sub txtcustomer_agente_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
    If Index = 1 Then
       txtrecibos_agente(0).SetFocus
       
    Else
      txtcustomer_agente(Index + 1).SetFocus
    End If
End If



If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If
End Sub


Private Sub txtdebit_agente_Change(Index As Integer)
On Error Resume Next
If valido1 = 66 Then Exit Sub
total_debit = 0
For t = 0 To 9
  total_debit = total_debit + Val(txtdebit_agente(t).Text)
Next t

lbltotal_debit_agent.Caption = Format(total_debit, "$###,##0.00")

lbltotal_debit_credit_agent.Caption = Format(total_debit + total_credit, "$###,##0.00")
lblgrantotal_debitcredit_agente.Caption = Format(total_debit + total_credit, "$###,##0.00")

total_dinero = Val(Format(lblgrantotal_debitcredit_agente.Caption, "000000.00")) + Val(Format(lblgrantotal_cash_agente.Caption, "000000.00"))
lbltotal_over_short_agente.Caption = Format(total_dinero - (total_reportado - total_voids), "$###,##0.00")

modificado = 1
End Sub

Private Sub txtdebit_agente_GotFocus(Index As Integer)
On Error Resume Next
If Val(txtdebit_agente(Index).Text) = 0 Then txtdebit_agente(Index).Text = ""
End Sub


Private Sub txtdebit_agente_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
    If Index = 9 Then
       txtcredit_agente(0).SetFocus
       
    Else
      txtdebit_agente(Index + 1).SetFocus
    End If
End If

If KeyAscii = Asc(".") Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If


End Sub


Private Sub txtdebit_agente_LostFocus(Index As Integer)
On Error Resume Next




img_debito(Index).Picture = LoadPicture()

For t = 1 To grid1.Rows - 1
   grid1.Row = t
   
   grid1.Col = 15
   cantidad_pagada$ = grid1.Text
   
   
   grid1.Col = 16
   metodo$ = grid1.Text
   
   
   
   If Val(txtdebit_agente(Index).Text) = Val(cantidad_pagada$) Then
      
      
      Select Case UCase(RTrim(LTrim(metodo$)))

      Case "VISA"
           img_debito(Index).Picture = visa.Picture
      Case "MASTERCARD"
        img_debito(Index).Picture = master.Picture
      Case "AMEX"
      img_debito(Index).Picture = american.Picture
      Case "DISCOVER"
      img_debito(Index).Picture = discover.Picture
      Case "DEBIT"
      img_debito(Index).Picture = debito.Picture
     
      End Select
   End If
   
Next t




pos = InStr(1, txtdebit_agente(Index).Text, ".")

If pos > 0 Then
  txtdebit_agente(Index).Text = Left(txtdebit_agente(Index).Text, pos + 2)
  If Mid(txtdebit_agente(Index).Text, pos + 1, 1) = "." Then
     
  Else
    If Mid(txtdebit_agente(Index).Text, pos + 2, 1) = "." Then
       txtdebit_agente(Index).Text = Left(txtdebit_agente(Index).Text, Len(txtdebit_agente(Index).Text) - 1)
    End If
  End If
  
  
  
  If Val(txtdebit_agente(Index).Text) = 0 Then
      txtdebit_agente(Index).Text = "0"
  End If
End If

txtdebit_agente(Index).Text = Format(txtdebit_agente(Index).Text, "#####0.00")


End Sub




Private Sub txtdebit_manager_Change()
On Error Resume Next
If valido1 = 66 Then Exit Sub
total_credito_oficina = 0

total_credito_oficina = Val(txtdebit_manager.Text) + Val(txtcredit_manager.Text)


lbltotal_debito_credito_manager.Caption = Format(total_credito_oficina, "$###,##0.00")

calcula_total_oficina_manager
''total_dinero = Val(Format(lblgrantotal_debitcredit_agente.Caption, "000000.00")) + Val(Format(lblgrantotal_cash_agente.Caption, "000000.00"))
''lbltotal_over_short_agente.Caption = Format((total_reportado - total_voids) - total_dinero, "$###,##0.00")

modificado = 1
End Sub

Private Sub txtdebit_manager_GotFocus()
On Error Resume Next
If Val(txtdebit_manager.Text) = 0 Then txtdebit_manager.Text = ""
End Sub

Private Sub txtdebit_manager_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
    
       txtcredit_manager.SetFocus
       Exit Sub
 
End If


If KeyAscii = Asc(".") Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If


End Sub


Private Sub txtdinero_Change(Index As Integer)
On Error Resume Next
If valido1 = 66 Then Exit Sub
total_dinero_oficina = 0
For t = 0 To 2
  total_dinero_oficina = total_dinero_oficina + Val(txtdinero(t).Text)
Next t

lbltotal_cash_manager.Caption = Format(total_dinero_oficina, "$###,##0.00")

calcula_total_oficina_manager
''total_dinero = Val(Format(lblgrantotal_debitcredit_agente.Caption, "000000.00")) + Val(Format(lblgrantotal_cash_agente.Caption, "000000.00"))
''lbltotal_over_short_agente.Caption = Format((total_reportado - total_voids) - total_dinero, "$###,##0.00")

modificado = 1
End Sub

Private Sub txtdinero_GotFocus(Index As Integer)
On Error Resume Next
If Val(txtdinero(Index).Text) = 0 Then txtdinero(Index).Text = ""

End Sub

Private Sub txtdinero_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
    If Index = 2 Then
       txtdebit_manager.SetFocus
       
    Else
      txtdinero(Index + 1).SetFocus
    End If
End If

If KeyAscii = Asc(".") Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If


End Sub


Private Sub txtnotas_agente_Change()
modificado = 1
End Sub

Private Sub txtoficina_ida_Change(Index As Integer)

End Sub

Private Sub txtoutput_Change()
On Error Resume Next
If Left(txtoutput.Text, 1) = "0" Then txtoutput.Text = Right(txtoutput.Text, Len(txtoutput.Text) - 1)
'If mem.Caption = "" Then txtoutput.Text = ""

End Sub

Private Sub txtoutput_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If


If KeyAscii = Asc("C") Or KeyAscii = Asc("c") Then
  KeyAscii = 0
  btnc_Click
  Exit Sub
End If



If KeyAscii = Asc("+") Then
  KeyAscii = 0
  btnmas_Click
  Exit Sub
End If

If KeyAscii = Asc("-") Then
KeyAscii = 0
btnminus_Click
Exit Sub
End If

If KeyAscii = Asc("*") Then
   KeyAscii = 0
   btnmul_Click
   Exit Sub
End If

If KeyAscii = Asc("/") Then
 KeyAscii = 0
 btndiv_Click
 Exit Sub
End If


   
If KeyAscii = 13 Or KeyAscii = Asc("=") Then
   KeyAscii = 0
   btnigual_Click
   Exit Sub
End If



If KeyAscii = Asc(".") Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If

End Sub


Private Sub txtrecibos_agente_Change(Index As Integer)
modificado = 1
End Sub

Private Sub txtrecibos_agente_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
    If Index = 1 Then
       txtamount_agente(0).SetFocus
       
    Else
      txtrecibos_agente(Index + 1).SetFocus
    End If
End If


If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If
End Sub





















Private Sub txttotal_LAE_agente_Change()
On Error Resume Next
total_reportado = Val(Format(txttotal_LAE_agente.Text, "0000000.00"))
lbltotal_reported_agent.Caption = Format(total_reportado, "$###,##0.00")

lbltotal_reportado_menos_voids_agente.Caption = Format(total_reportado - total_voids, "$###,##0.00")


total_dinero = Val(Format(lblgrantotal_debitcredit_agente.Caption, "000000.00")) + Val(Format(lblgrantotal_cash_agente.Caption, "000000.00"))
lbltotal_over_short_agente.Caption = Format(total_dinero - (total_reportado - total_voids), "$###,##0.00")

modificado = 1

End Sub

Private Sub txttotal_LAE_agente_GotFocus()
On Error Resume Next

txttotal_LAE_agente.Text = Format(txttotal_LAE_agente.Text, "#####0.00")

End Sub


Private Sub txttotal_LAE_agente_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 8 Then
  Exit Sub
End If

If KeyAscii = 13 Then
   txttotal_dmv_agente.SetFocus
End If

If KeyAscii = Asc(".") Then
  Exit Sub
End If

If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
  KeyAscii = 0
End If
End Sub


Private Sub txttotal_LAE_agente_LostFocus()
On Error Resume Next


pos = InStr(1, txttotal_LAE_agente.Text, ".")

If pos > 0 Then
  txttotal_LAE_agente.Text = Left(txttotal_LAE_agente.Text, pos + 2)
  If Mid(txttotal_LAE_agente.Text, pos + 1, 1) = "." Then
     
  Else
    If Mid(txttotal_LAE_agente.Text, pos + 2, 1) = "." Then
       txttotal_LAE_agente.Text = Left(txttotal_LAE_agente.Text, Len(txttotal_LAE_agente.Text) - 1)
    End If
  End If
  
  
  
  If Val(txttotal_LAE_agente.Text) = 0 Then
      txttotal_LAE_agente.Text = "0"
  End If
End If

txttotal_LAE_agente.Text = Format(txttotal_LAE_agente.Text, "$###,##0.00")

End Sub



Public Sub limpia_datos()
On Error Resume Next

 msg.Visible = True

For t = 0 To 5
txtcash(t).Text = ""
Next t

For t = 0 To 19
 txtdebit_agente(t).Text = ""
 txtcredit_agente(t).Text = ""
Next t

tabcard(0).Visible = True
tabcard(1).Visible = False
      
op_cards(0).Value = True

txtnotas_agente.Text = ""

For t = 0 To 1
 txtcustomer_agente(t).Text = ""
 txtrecibos_agente(t).Text = ""
 txtamount_agente(t).Text = ""
Next t

txttotal_LAE_agente.Text = ""


ListView1.ListItems.Clear

With ListView1
.ColumnHeaders.Clear
End With
modificado = 0

chk_dayoff.Value = False
grid1.Clear
grid3.Clear
visualizador.Visible = False
limpia_tarjetas
arrow(0).Visible = False
arrow(1).Visible = False
 msg.Visible = False
 
 ' limpia tercer folder
 
 For t = 0 To 2
   txtdinero(t).Text = ""
   txtcant_ida(t).Text = ""
   txtcant_venida(t).Text = ""
 Next t
 
 txtdebit_manager.Text = ""
 txtcredit_manager.Text = ""
 
 txttotal_venta_manager.Text = ""
 txtnotas_manager.Text = ""
 
 For t = 0 To 5
    cbo_employees(t).ListIndex = -1
    cbooficina1(t).ListIndex = -1
 Next t
 
 firma.Caption = ""
 chkmanager.Value = False
 Firma_agente.Caption = ""
 chk_firma_agente.Value = False
 tipo_vista = 1
 op_view(0).Value = True
 
 tipo_guardado = 1
End Sub

Public Sub graba_datos()
On Error Resume Next

    Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    Set Rs = New ADODB.Recordset
    
    
    ruta_archivos$ = "\\192.168.84.215\moneyreport\"


If Dir$(ruta_archivos$ + "encendido.txt") = "" Then
     MsgBox "Files cannot be saved. Apparently the storage server is down. Contact the IT department to solve this problem. Thanks", 16, "ATTENTION"
     Exit Sub
End If

   
   
    oficina$ = LTrim(Right(UCase(RTrim(cbo_oficina.List(cbo_oficina.ListIndex))), 30))
    sSelect = "SELECT idoffice From officescatalog where office='" + oficina$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    id_office$ = Rs(0)
    Rs.Close
    
    
    
    
    UserName$ = RTrim(Left(cbo_agentes.List(cbo_agentes.ListIndex), Len(cbo_agentes.List(cbo_agentes.ListIndex)) - 5))
    sSelect = "SELECT idemployee From employeeinfo where username='" + UserName$ + "'"  ' and active='1'"
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    id_employee$ = Rs(0)
    Rs.Close
    
    
    
    UserName2$ = RTrim(Left(cbo_managers.List(cbo_managers.ListIndex), Len(cbo_managers.List(cbo_managers.ListIndex)) - 5))
    sSelect = "SELECT idemployee From employeeinfo where username='" + UserName2$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    ID_manager$ = Rs(0)
    Rs.Close
    
    
    
     ' obtiene fecha actual
     sSelect = "SELECT GETDATE()"
     Rs.Open sSelect, base, adOpenUnspecified
     fecha_real$ = Format(Rs(0), "mm/dd/yyyy hh:mm am/pm")
     Rs.Close
         
     fecha_creada$ = Format(Now, "mm/dd/yyyy")
    
     
     
If tipo_guardado = 1 Then
     
     
     ' verifica si ya existe el reporte
     
     sSelect = "select idmoneyreport from moneyreport where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "' and idoffice='" + id_office$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreport$ = Rs(0)
     Rs.Close
     
     
     
     
   
  tot$ = Format(txttotal_LAE_agente.Text, "000000.00")
  f1$ = fecha_real$
  f2$ = lbldate_agente.Caption
  
  
   ' quita las comillas simples a la nota
  r$ = txtnotas_agente.Text
  texto$ = ""
  For z = 1 To Len(r$)
    If Mid$(r$, z, 1) <> "'" Then
         texto$ = texto$ + Mid$(r$, z, 1)
    End If
  Next z
  txtnotas_agente.Text = texto$
  notas$ = txtnotas_agente.Text
  
  
  
  
  Dim descanso As Boolean
  
  If chk_dayoff.Value = False Then
     descanso = 0
  Else
     descanso = 1
  End If
  
  chk_dayoff.Value = descanso
  msgdescanso.Visible = descanso
  
  
  
  rev$ = Format(chk_revisado.Value, "0")
    
  Set Rs = New ADODB.Recordset
    
  If id_moneyreport$ = "" Then  ' administrador = 0
    
    sSelect = "insert into MoneyReport (idoffice, idemployee, idmanager, datecreated, year, month, datereport, Notes, TotalLAE, Pathfiles, Reviewed, Lastupdated, active, submitted, DayOff)  VALUES ('" & _
    id_office$ + "', '" + id_employee$ + "', '" + ID_manager$ + "', convert(datetime, '" + f1$ + "'), '" + Mid$(f1$, 7, 4) + "', '" + Left(f1$, 2) & _
    "', convert(datetime, '" + f2$ + "'), '" + notas$ + "', convert(money, '" + tot$ + "'), '" & _
    ruta_archivos$ + "', '" + rev$ + "',  convert(datetime, '" + f1$ + "'), '1', '0', '" + Format(descanso, "0") + "')"
    
    
 Else
     
  
    sSelect = "update MoneyReport set idoffice='" + id_office$ + "', idemployee='" + id_employee$ + "',  idmanager='" + ID_manager$ + "', " & _
    "datereport=convert(datetime, '" + f2$ + "'), notes='" + notas$ + "', TotalLAE=convert(money, '" + tot$ + "'), Pathfiles='" + ruta_archivos$ + "', Reviewed='" & _
    rev$ + "', lastupdated= convert(datetime, '" + fecha_real$ + "'), active='1', submitted='0', DayOff='" + Format(descanso, "0") + "' " & _
    "where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "' and idoffice='" + id_office$ + "'"

  
       
  End If
  
                      
    Rs.Open sSelect, base, adOpenUnspecified
    
    Rs.Close
   
   
   ' obtiene el ID del money report que se grabó
    sSelect = "select idmoneyreport from moneyreport where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "' and idoffice='" + id_office$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreport$ = Rs(0)
     Rs.Close
     
        
     
     
     
     
     
     
     
   
   ' graba los pagos
   '********************************************************************************************************************
   
        idmoneyreportpaymetrel$ = ""
    
        sSelect = "select idmoneyreportpaymetrel from moneyreportpaymetrel where idmoneyreport='" + id_moneyreport$ + "' and idpaymethod='2' and indice='1'"
        Rs.Open sSelect, base, adOpenUnspecified
        idmoneyreportpaymetrel$ = Rs(0)
        Rs.Close
                 
        If idmoneyreportpaymetrel$ <> "" Then
             ' Graba CASh=2
             sSelect = "update moneyreportpaymetrel set idmoneyreport='" + id_moneyreport$ + "', idpaymethod='2', amount=convert(money, '" + txtcash(0).Text + "'), " & _
             "datecreated=convert(datetime, '" + f2$ + "'), active='1', indice='1' where idmoneyreportpaymetrel='" + idmoneyreportpaymetrel$ + "'"
        
        Else
        
             sSelect = "insert into moneyreportpaymetrel (idmoneyreport, idpaymethod, amount, datecreated, active, indice) VALUES ('" & _
             id_moneyreport$ + "', '2', convert(money, '" + txtcash(0).Text + "'), convert(datetime, '" + f2$ + "'), '1', '1')"
               
        End If
    
   
   
      
   Rs.Open sSelect, base, adOpenUnspecified
   Rs.Close
   
   
   
' -----------------------  MONEY ORDER

        idmoneyreportpaymetrel_1$ = ""
    
        sSelect = "select idmoneyreportpaymetrel from moneyreportpaymetrel where idmoneyreport='" + id_moneyreport$ + "' and idpaymethod='9' and indice='1'"
        Rs.Open sSelect, base, adOpenUnspecified
        idmoneyreportpaymetrel_1$ = Rs(0)
        Rs.Close
                 
        idmoneyreportpaymetrel_2$ = ""
                 
        sSelect = "select idmoneyreportpaymetrel from moneyreportpaymetrel where idmoneyreport='" + id_moneyreport$ + "' and idpaymethod='9' and indice='2'"
        Rs.Open sSelect, base, adOpenUnspecified
        idmoneyreportpaymetrel_2$ = Rs(0)
        Rs.Close
                 
                 
        If idmoneyreportpaymetrel_1$ <> "" Then
             
             sSelect = "update moneyreportpaymetrel set idmoneyreport='" + id_moneyreport$ + "', idpaymethod='9', amount=convert(money, '" + txtcash(1).Text + "'), " & _
             "datecreated=convert(datetime, '" + f2$ + "'), active='1', indice='1' where idmoneyreportpaymetrel='" + idmoneyreportpaymetrel_1$ + "'"
        
        Else
        
             sSelect = "insert into moneyreportpaymetrel (idmoneyreport, idpaymethod, amount, datecreated, active, indice) VALUES ('" & _
             id_moneyreport$ + "', '9', convert(money, '" + txtcash(1).Text + "'), convert(datetime, '" + f2$ + "'), '1', '1')"
               
        End If
    
   
      
        Rs.Open sSelect, base, adOpenUnspecified
        Rs.Close




         If idmoneyreportpaymetrel_2$ <> "" Then
             
             sSelect = "update moneyreportpaymetrel set idmoneyreport='" + id_moneyreport$ + "', idpaymethod='9', amount=convert(money, '" + txtcash(2).Text + "'), " & _
             "datecreated=convert(datetime, '" + f2$ + "'), active='1', indice='2' where idmoneyreportpaymetrel='" + idmoneyreportpaymetrel_2$ + "'"
        
        Else
        
             sSelect = "insert into moneyreportpaymetrel (idmoneyreport, idpaymethod, amount, datecreated, active,indice) VALUES ('" & _
             id_moneyreport$ + "', '9', convert(money, '" + txtcash(2).Text + "'), convert(datetime, '" + f2$ + "'), '1', '2')"
               
        End If
    
          
        Rs.Open sSelect, base, adOpenUnspecified
        Rs.Close
   
   
       
   
' -----------------------  CHECKS

        idmoneyreportpaymetrel_1$ = ""
    
        sSelect = "select idmoneyreportpaymetrel from moneyreportpaymetrel where idmoneyreport='" + id_moneyreport$ + "'  and idpaymethod='10' and indice='1'"
        Rs.Open sSelect, base, adOpenUnspecified
        idmoneyreportpaymetrel_1$ = Rs(0)
        Rs.Close
                 
        If idmoneyreportpaymetrel_1$ <> "" Then
             
             sSelect = "update moneyreportpaymetrel set idmoneyreport='" + id_moneyreport$ + "', idpaymethod='10', amount=convert(money, '" + txtcash(3).Text + "'), " & _
             "datecreated=convert(datetime, '" + f2$ + "'), active='1', indice='1' where idmoneyreportpaymetrel='" + idmoneyreportpaymetrel_1$ + "'"
        
        Else
        
             sSelect = "insert into moneyreportpaymetrel (idmoneyreport, idpaymethod, amount, datecreated, active, indice) VALUES ('" & _
             id_moneyreport$ + "', '10', convert(money, '" + txtcash(3).Text + "'), convert(datetime, '" + f2$ + "'), '1', '1')"
               
        End If
    
   
      
   Rs.Open sSelect, base, adOpenUnspecified
   Rs.Close
   
   
        idmoneyreportpaymetrel_2$ = ""
    
        sSelect = "select idmoneyreportpaymetrel from moneyreportpaymetrel where idmoneyreport='" + id_moneyreport$ + "'  and idpaymethod='10' and indice='2'"
        Rs.Open sSelect, base, adOpenUnspecified
        idmoneyreportpaymetrel_2$ = Rs(0)
        Rs.Close
                 
        If idmoneyreportpaymetrel_2$ <> "" Then
             
             sSelect = "update moneyreportpaymetrel set idmoneyreport='" + id_moneyreport$ + "', idpaymethod='10', amount=convert(money, '" + txtcash(4).Text + "'), " & _
             "datecreated=convert(datetime, '" + f2$ + "'), active='1', indice='2' where idmoneyreportpaymetrel='" + idmoneyreportpaymetrel_2$ + "'"
        
        Else
        
             sSelect = "insert into moneyreportpaymetrel (idmoneyreport, idpaymethod, amount, datecreated, active, indice) VALUES ('" & _
             id_moneyreport$ + "', '10', convert(money, '" + txtcash(4).Text + "'), convert(datetime, '" + f2$ + "'), '1', '2')"
               
        End If
    
   
      
   Rs.Open sSelect, base, adOpenUnspecified
   Rs.Close
   
   
   
' -----------------------  COINS

        idmoneyreportpaymetrel$ = ""
        
        sSelect = "select idmoneyreportpaymetrel from moneyreportpaymetrel where idmoneyreport='" + id_moneyreport$ + "' and idpaymethod='12' and indice='1'"
        Rs.Open sSelect, base, adOpenUnspecified
        idmoneyreportpaymetrel$ = Rs(0)
        Rs.Close
                 
        If idmoneyreportpaymetrel$ <> "" Then
             
             sSelect = "update moneyreportpaymetrel set idmoneyreport='" + id_moneyreport$ + "', idpaymethod='12', amount=convert(money, '" + txtcash(5).Text + "'), " & _
             "datecreated=convert(datetime, '" + f2$ + "'), active='1', indice='1' where idmoneyreportpaymetrel='" + idmoneyreportpaymetrel$ + "'"
        
        Else
        
             sSelect = "insert into moneyreportpaymetrel (idmoneyreport, idpaymethod, amount, datecreated, active, indice) VALUES ('" & _
             id_moneyreport$ + "', '12', convert(money, '" + txtcash(5).Text + "'), convert(datetime, '" + f2$ + "'), '1', '1')"
               
        End If
    
   
      
   Rs.Open sSelect, base, adOpenUnspecified
   Rs.Close
   
   
   
   
   
   
   
   
   ' -----------------------  DEBIT CARDS
   
 For z = 0 To 9

        idmoneyreportpaymetrel$ = ""
          
        sSelect = "select idmoneyreportpaymetrel from moneyreportpaymetrel where idmoneyreport='" + id_moneyreport$ + "' and idpaymethod='5' and indice='" + Format(z + 1, "0") + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        idmoneyreportpaymetrel$ = Rs(0)
        Rs.Close
                 
        If idmoneyreportpaymetrel$ <> "" Then
             
             sSelect = "update moneyreportpaymetrel set idmoneyreport='" + id_moneyreport$ + "', idpaymethod='5', amount=convert(money, '" + txtdebit_agente(z).Text + "'), " & _
             "datecreated=convert(datetime, '" + f2$ + "'), active='1', indice='" + Format(z + 1, "0") + "' where idmoneyreportpaymetrel='" + idmoneyreportpaymetrel$ + "'"
        
        Else
        
             sSelect = "insert into moneyreportpaymetrel (idmoneyreport, idpaymethod, amount, datecreated, active, indice) VALUES ('" & _
             id_moneyreport$ + "', '5', convert(money, '" + txtdebit_agente(z).Text + "'), convert(datetime, '" + f2$ + "'), '1', '" + Format(z + 1, "0") + "')"
               
        End If
    
   
      
   Rs.Open sSelect, base, adOpenUnspecified
   Rs.Close
   
 Next z
   
   
   
   
     
   ' -----------------------  CREDIT CARDS
   
   
   
 For z = 0 To 19

        idmoneyreportpaymetrel$ = ""
    
        sSelect = "select idmoneyreportpaymetrel from moneyreportpaymetrel where idmoneyreport='" + id_moneyreport$ + "' and idpaymethod='4' and indice='" + Format(z + 1, "0") + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        idmoneyreportpaymetrel$ = Rs(0)
        Rs.Close
                 
        If idmoneyreportpaymetrel$ <> "" Then
             
             sSelect = "update moneyreportpaymetrel set idmoneyreport='" + id_moneyreport$ + "', idpaymethod='4', amount=convert(money, '" + txtcredit_agente(z).Text + "'), " & _
             "datecreated=convert(datetime, '" + f2$ + "'), active='1', indice='" + Format(z + 1, "0") + "' where idmoneyreportpaymetrel='" + idmoneyreportpaymetrel$ + "'"
        
        Else
        
             sSelect = "insert into moneyreportpaymetrel (idmoneyreport, idpaymethod, amount, datecreated, active, indice) VALUES ('" & _
             id_moneyreport$ + "', '4', convert(money, '" + txtcredit_agente(z).Text + "'), convert(datetime, '" + f2$ + "'), '1', '" + Format(z + 1, "0") + "')"
               
        End If
    
   
      
   Rs.Open sSelect, base, adOpenUnspecified
   Rs.Close
   
 Next z
   
   
' GoTo salta
   
   
  ' -----------------------  VOIDS
   
    ' Graba VOIDS
   
   
   
    Grid2.Clear
    
    sSelect = "select IDReceiptHDR, IdCustomer, AmountPaid, Date, Void from ReceiptsHDR " & _
    "Where Void = 1 and cast(Date as Date) >= '" + f2$ + "' AND cast( DATE as Date) <= '" + f2$ + "' " & _
    "and IdEmployeeUSR='" + id_employee$ + "' order by Date"
    
       ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
   
   
 For z = 0 To 1
   
    
    ' verifica que exista la poliza en LAE y que este VOID
    
    existe = 0
    For Y = 1 To Grid2.Rows - 1
       Grid2.Row = Y
       Grid2.Col = 1
       id_receiptHDR$ = Grid2.Text
       
       Grid2.Col = 2
       id_customer$ = Grid2.Text
       
       Grid2.Col = 3
       cantidad_void$ = Grid2.Text
       
       If LTrim(RTrim(txtcustomer_agente(z).Text)) = LTrim(RTrim(id_customer$)) Then
             If Val(txtamount_agente(z).Text) = Val(cantidad_void$) Then
                       existe = 1
                       Exit For
             End If
       End If
                       
       
     Next Y
    
    
    
     If existe = 1 Then
    
    
            sSelect = "select idmoneyreprecvoidrel from moneyreportrecvoidrel  where idmoneyreport='" + id_moneyreport$ + "' and idreceiptHDR='" + id_receiptHDR$ + "'"
            Rs.Open sSelect, base, adOpenUnspecified
            idmoneyreprecvoidrel$ = Rs(0)
            Rs.Close
                 
            If idmoneyreprecvoidrel$ <> "" Then
             
               sSelect = "update moneyreportrecvoidrel set idmoneyreport='" + id_moneyreport$ + "', idreceiptHDR='" + id_receiptHDR$ + "', " & _
               "datecreated=convert(datetime, '" + f2$ + "'), active='1' where idmoneyreprecvoidrel='" + idmoneyreprecvoidrel$ + "'"
        
            Else
        
               sSelect = "insert into moneyreportrecvoidrel (idmoneyreport, idreceiptHDR, datecreated, active) VALUES ('" & _
               id_moneyreport$ + "', '" + id_receiptHDR$ + "',  convert(datetime, '" + f2$ + "'), '1') "
                      
            End If
    
            Rs.Open sSelect, base, adOpenUnspecified
            Rs.Close
   
     End If
   
 Next z
     
   
   
   
   
ElseIf tipo_guardado = 2 Then
        
        
   ' ********************************************************************************************************************************************************
   ' ********************************************************************************************************************************************************
   ' ********************************************************************************************************************************************************
   
        
   
     ' verifica si ya existe el reporte de la oficina
     
     sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idoffice='" + id_office$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreportoffice$ = Rs(0)
     Rs.Close
     
   
    f1$ = fecha_real$
    f2$ = lbldate_agente.Caption
  
    
     
     
   ' graba el reporte diario de la oficina
   ' --------------------------------------------------------------------------------------------------------------------------------
   
  Dim verificado_por_manager As Boolean
  Dim verificado_por_Accounting As Boolean
  
  Dim part2 As Variant, part3 As Variant
  
  verificado_por_manager = chkmanager.Value
  verificado_por_Accounting = chk_revisado.Value
  
  
  ' quita las comillas simples a la nota
  r$ = txtnotas_manager.Text
  texto$ = ""
  For z = 1 To Len(r$)
    If Mid$(r$, z, 1) <> "'" Then
         texto$ = texto$ + Mid$(r$, z, 1)
    End If
  Next z
  txtnotas_manager.Text = texto$
  
  
   
  Set Rs = New ADODB.Recordset
  
   
   
    
  If id_moneyreportoffice$ = "" Then    ' administrador = 0
    
    
    sSelect = "insert into moneyreportbyoffice (idoffice, idmanager, datecreated, year, month, datereport, Dunbarnet, coins, moneyorder, debit, credit, laetotal) Values (' " & _
    id_office$ + "', '" + ID_manager$ + "', convert(datetime, '" + f1$ + "'), '" + Mid$(f1$, 7, 4) + "', '" + Left(f1$, 2) & _
    "', convert(datetime, '" + f2$ + "'), '" + txtdinero(0).Text + "', '" + txtdinero(1).Text + "', '" + txtdinero(2).Text + "', '" & _
    txtdebit_manager.Text + "', '" + txtcredit_manager.Text + "', '" + lbltotal_LAE_oficina.Caption + "')"
    
    Rs.Open sSelect, base, adOpenUnspecified
    Rs.Close
    
    
    ' agarra el IDmoneyreportOffice
    sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idoffice='" + id_office$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_moneyreportoffice$ = Rs(0)
    Rs.Close
    
    
    ' inserta la nota
    
     sSelect = "update moneyreportbyoffice set notes='" + txtnotas_manager.Text + " where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idoffice='" + id_office$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreportoffice$ = Rs(0)
     Rs.Close
    
    
    
    r$ = LTrim(RTrim(Right(cbooficina1(0).List(cbooficina1(0).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice0$ = Rs(0)
    Rs.Close
    
    r$ = LTrim(RTrim(Right(cbooficina1(1).List(cbooficina1(1).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice1$ = Rs(0)
    Rs.Close
    
    r$ = LTrim(RTrim(Right(cbooficina1(2).List(cbooficina1(2).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice2$ = Rs(0)
    Rs.Close
    
    r$ = LTrim(RTrim(Right(cbooficina1(3).List(cbooficina1(3).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice3$ = Rs(0)
    Rs.Close
    
    r$ = LTrim(RTrim(Right(cbooficina1(4).List(cbooficina1(4).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice4$ = Rs(0)
    Rs.Close
    
    r$ = LTrim(RTrim(Right(cbooficina1(5).List(cbooficina1(5).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice5$ = Rs(0)
    Rs.Close
    
    
    
    
    
    
    part2 = "agentto3='" + LTrim(RTrim(Right(cbo_employees(2).List(cbo_employees(2).ListIndex), 10))) + "', officeto3='" + IDOffice2$ + "', amountto3='" + txtcant_ida(2).Text & _
    "', agentfrom1='" + LTrim(RTrim(Right(cbo_employees(3).List(cbo_employees(3).ListIndex), 10))) + "', officefrom1='" + IDOffice3$ + "', amountfrom1='" + txtcant_venida(0).Text & _
    "', agentfrom2='" + LTrim(RTrim(Right(cbo_employees(4).List(cbo_employees(4).ListIndex), 10))) + "', officefrom2='" + IDOffice4$ + "', amountfrom2='" + txtcant_venida(1).Text + "', "
    
    part3 = "agentfrom3='" + LTrim(RTrim(Right(cbo_employees(5).List(cbo_employees(5).ListIndex), 10))) + "', officefrom3='" + IDOffice5$ + "', amountfrom3='" + txtcant_venida(2).Text & _
    "', reviewbymanager='" + Format(verificado_por_manager, "0") + "', active='1', submitted='0' " & _
    "where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"
    
    ' *************************************************************************************************************************
    ' reviewbyaccounting='" + Format(verificado_por_Accounting, "0")  debe asignarlo JOSELIN
    ' *************************************************************************************************************************
    
    
    sSelect = "update moneyreportbyoffice set agentto1='" + LTrim(RTrim(Right(cbo_employees(0).List(cbo_employees(0).ListIndex), 10))) + "', officeto1='" + IDOffice0$ + "', amountto1='" + txtcant_ida(0).Text & _
    "', agentto2='" + LTrim(RTrim(Right(cbo_employees(1).List(cbo_employees(1).ListIndex), 10))) + "', officeto2='" + IDOffice1$ + "', amountto2='" + txtcant_ida(1).Text + "'," + part2 + part3
    
    Rs.Open sSelect, base, adOpenUnspecified
    Rs.Close
   
    
 Else
     
  
  
 ' agarra el IDmoneyreportOffice
    sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idoffice='" + id_office$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_moneyreportoffice$ = Rs(0)
    Rs.Close
     
    
    sSelect = "update moneyreportbyoffice set idoffice='" + id_office$ + "', idmanager='" + ID_manager$ + "', datecreated=convert(datetime, '" + f1$ + "'), " & _
    "year='" + Mid$(f1$, 7, 4) + "', month='" + Left(f1$, 2) + "', datereport=convert(datetime, '" + f2$ + "'), Dunbarnet='" + txtdinero(0).Text + "', " & _
    "coins='" + txtdinero(1).Text + "', moneyorder='" + txtdinero(2).Text + "', debit='" + txtdebit_manager.Text + "', credit='" + txtcredit_manager.Text + "', " & _
    "laetotal='" + lbltotal_LAE_oficina.Caption + "' where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"
    
    Rs.Open sSelect, base, adOpenUnspecified
    Rs.Close
    
  
    
    
     r$ = LTrim(RTrim(Right(cbooficina1(0).List(cbooficina1(0).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice0$ = Rs(0)
    Rs.Close
    
    r$ = LTrim(RTrim(Right(cbooficina1(1).List(cbooficina1(1).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice1$ = Rs(0)
    Rs.Close
    
    r$ = LTrim(RTrim(Right(cbooficina1(2).List(cbooficina1(2).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice2$ = Rs(0)
    Rs.Close
    
    r$ = LTrim(RTrim(Right(cbooficina1(3).List(cbooficina1(3).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice3$ = Rs(0)
    Rs.Close
    
    r$ = LTrim(RTrim(Right(cbooficina1(4).List(cbooficina1(4).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice4$ = Rs(0)
    Rs.Close
    
    r$ = LTrim(RTrim(Right(cbooficina1(5).List(cbooficina1(5).ListIndex), 20)))
    sSelect = "select idoffice from officescatalog where office='" + r$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    IDOffice5$ = Rs(0)
    Rs.Close
    
    
    
    part2 = "agentto3='" + LTrim(RTrim(Right(cbo_employees(2).List(cbo_employees(2).ListIndex), 10))) + "', officeto3='" + IDOffice2$ + "', amountto3='" + txtcant_ida(2).Text & _
    "', agentfrom1='" + LTrim(RTrim(Right(cbo_employees(3).List(cbo_employees(3).ListIndex), 10))) + "', officefrom1='" + IDOffice3$ + "', amountfrom1='" + txtcant_venida(0).Text & _
    "', agentfrom2='" + LTrim(RTrim(Right(cbo_employees(4).List(cbo_employees(4).ListIndex), 10))) + "', officefrom2='" + IDOffice4$ + "', amountfrom2='" + txtcant_venida(1).Text + "', "
    
    part3 = "agentfrom3='" + LTrim(RTrim(Right(cbo_employees(5).List(cbo_employees(5).ListIndex), 10))) + "', officefrom3='" + IDOffice5$ + "', amountfrom3='" + txtcant_venida(2).Text & _
    "', reviewbymanager='" + Format(verificado_por_manager, "0") + "', active='1', submitted='0' " & _
    "where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"
    
    ' *************************************************************************************************************************
    ' reviewbyaccounting='" + Format(verificado_por_Accounting, "0")  debe asignarlo JOSELIN
    ' *************************************************************************************************************************
    
    
    
    
    sSelect = "update moneyreportbyoffice set agentto1='" + LTrim(RTrim(Right(cbo_employees(0).List(cbo_employees(0).ListIndex), 10))) + "', officeto1='" + IDOffice0$ + "', amountto1='" + txtcant_ida(0).Text & _
    "', agentto2='" + LTrim(RTrim(Right(cbo_employees(1).List(cbo_employees(1).ListIndex), 10))) + "', officeto2='" + IDOffice1$ + "', amountto2='" + txtcant_ida(1).Text + "'," + part2 + part3
    
     
    Rs.Open sSelect, base, adOpenUnspecified
    Rs.Close
    
        
    
    sSelect = "update moneyreportbyoffice set notes='" + txtnotas_manager.Text + "', dunbarnet='" + txtdinero(0).Text + "', coins='" + txtdinero(1).Text + "', moneyorder='" + txtdinero(2).Text + "'," & _
    "debit='" + txtdebit_manager.Text + "', credit='" + txtcredit_manager.Text + "', laetotal='" + lbltotal_LAE_oficina.Caption + "' where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"
    
    Rs.Open sSelect, base, adOpenUnspecified
    Rs.Close
       
       
       
  End If
  
                      
    
   ' ********************************************************************************************************************************************************
   ' ********************************************************************************************************************************************************
   ' ********************************************************************************************************************************************************
   
   
   
   
   
 
 End If
   
   
   
salta:
   
   
   
   ' hasta el ultimo tiene que grabar los archivos
' *************************************************************************************

If tipo_guardado = 1 Then

' crea ruta donde grabara las imagenes por agente
MkDir "\\192.168.84.215\moneyreport\" + id_moneyreport$

ruta_archivos$ = "\\192.168.84.215\moneyreport\" + id_moneyreport$ + "\"



     sSelect = "update MoneyReport set Pathfiles='" + ruta_archivos$ + "' where idmoneyreport='" + id_moneyreport$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     a$ = Rs(0)
     Rs.Close


For t = 1 To ListView1.ListItems.Count

   n$ = ListView1.ListItems(t).Text
   ' idmoneyreport
   
   Source = "C:\money\" + n$
   Target = ruta_archivos$ + n$


'Copy File
a = CopyFile(Trim$(Source), Trim(Target), False)

' verifica si se copió el archivo en la ruta
  If Dir$(Target) <> "" Then
     Kill Source
     
  Else
     a = CopyFile(Trim$(Source), Trim(Target), False)
     If Dir$(Target) <> "" Then
        Kill Source
     Else
        Kill Source
        MsgBox "The file named " + n$ + " could not be copied. Try again to load the file and then save it again.", 64, "ERROR"
        
        ListView1.ListItems.Remove (ListView1.SelectedItem.Index)
        archivo_selecto$ = ""
        img1.Picture = LoadPicture()
        pdf1.src = "c:\"

     End If
  
  End If
  
Next t



End If





' ------------------------------------------------------------------------------

If tipo_guardado = 2 Then


 sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idoffice='" + id_office$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreportoffice$ = Rs(0)
     Rs.Close
     

' crea ruta donde grabara las imagenes por manager de office
 MkDir "\\192.168.84.215\moneyreport\O-" + id_moneyreportoffice$

 ruta_archivos$ = "\\192.168.84.215\moneyreport\O-" + id_moneyreportoffice$ + "\"


     sSelect = "update moneyreportbyoffice set Pathfiles='" + ruta_archivos$ + "' where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     a$ = Rs(0)
     Rs.Close


For t = 1 To ListView2.ListItems.Count

   n$ = ListView2.ListItems(t).Text
   ' idmoneyreport
   
   Source = "C:\money\" + n$
   Target = ruta_archivos$ + n$


'Copy File
a = CopyFile(Trim$(Source), Trim(Target), False)

' verifica si se copió el archivo en la ruta
  If Dir$(Target) <> "" Then
     Kill Source
  Else
     a = CopyFile(Trim$(Source), Trim(Target), False)
     If Dir$(Target) <> "" Then
        Kill Source
     Else
        Kill Source
        MsgBox "The file named " + n$ + " could not be copied. Try again to load the file and then save it again.", 64, "ERROR"
        
        ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
        archivo_selecto$ = ""
        img1.Picture = LoadPicture()
        pdf1.src = "c:\"

     End If

  End If
  
Next t


End If
   
   
If tipo_guardado = 1 Then
     
     
     ' verifica si ya existe el reporte
     
     sSelect = "select idmoneyreport from moneyreport where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + Format(id_employee, "###0") + "' and idoffice='" + id_oficina$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreport$ = Rs(0)
     Rs.Close
        
     lblnum.Caption = id_moneyreport$
    
Else
  
       
    
     sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idoffice='" + id_oficina$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreportoffice$ = Rs(0)
     Rs.Close
    
     lblnum.Caption = "O-" + id_moneyreportoffice$
     
     
End If
   
   
End Sub

Public Sub obtener_fecha_real()
On Error Resume Next


Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
 msg.Visible = True
    
    
    Set Rs = New ADODB.Recordset
    
     sSelect = "SELECT GETDATE()"
       Rs.Open sSelect, base, adOpenUnspecified
       fecha_actual$ = Format(Rs(0), "mm/dd/yyyy")
       
       Rs.Close
       
       fecha_computadora$ = Format(Now, "mm/dd/yyyy")
       
       'lbldate_agente.Caption = "03/30/2021"
       
       If fecha_computadora$ <> fecha_actual$ Then
           
      
       End If
    
    
    
       btnerase_archivo.Enabled = True
       chk_dayoff.Enabled = True
       btnsave.Enabled = True
       
         
    
    Set Rs = New ADODB.Recordset
    
   
    oficina$ = LTrim(Right(UCase(RTrim(cbo_oficina.List(cbo_oficina.ListIndex))), 30))
    sSelect = "SELECT idoffice From officescatalog where office='" + oficina$ + "'"  ' and active='1'"
    
    ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    id_office$ = Rs(0)
    Rs.Close
    
    
    
    
    
    
     ' verifica si ya existe el reporte
     
  
   
  If tipo_guardado = 1 Then
   
        UserName$ = RTrim(Left(cbo_agentes.List(cbo_agentes.ListIndex), Len(cbo_agentes.List(cbo_agentes.ListIndex)) - 5))
        sSelect = "SELECT idemployee From employeeinfo where username='" + UserName$ + "'"  ' and active='1'"
        ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
        Rs.Open sSelect, base, adOpenUnspecified
        id_employee$ = Rs(0)
        Rs.Close
    
    
       id_moneyreport$ = ""
     
       sSelect = "select idmoneyreport from moneyreport where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + id_employee$ + "' and idoffice='" + id_office$ + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       id_moneyreport$ = Rs(0)
       Rs.Close
     
    
       If id_moneyreport$ = "" Then
            btnerase_archivo.Enabled = True
            chk_dayoff.Enabled = True
            btnsave.Enabled = True
            btnsave.Picture = img_disk_up.Picture

            btnsend.Enabled = True
            btnsend.Picture = img_send_up.Picture
     
            msg.Visible = False
            Exit Sub
   
       End If
    
   
    
       sSelect = "SELECT submitted From moneyreport where idmoneyreport='" + id_moneyreport$ + "'"
       ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
       Rs.Open sSelect, base, adOpenUnspecified
       submitido = Rs(0)
       Rs.Close
    
       
       If submitido = "True" Then
            btnerase_archivo.Enabled = False
            chk_dayoff.Enabled = False
            btnsave.Enabled = False
            btnsave.Picture = img_caja.Picture
            btnsend.Enabled = False
            btnsend.Picture = img_caja.Picture
         

       Else
           btnerase_archivo.Enabled = True
           btnsave.Enabled = True
           chk_dayoff.Enabled = True
           btnsave.Picture = img_disk_up.Picture
           btnsend.Enabled = True
           btnsend.Picture = img_send_up.Picture
       
       End If
     
     
     
  Else
   
    
       UserName$ = RTrim(Left(cbo_managers.List(cbo_managers.ListIndex), Len(cbo_managers.List(cbo_managers.ListIndex)) - 5))
       sSelect = "SELECT idemployee From employeeinfo where username='" + UserName$ + "'"  ' and active='1'"
      ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
       Rs.Open sSelect, base, adOpenUnspecified
       ID_manager$ = Rs(0)
       Rs.Close
       
   
       id_moneyreportoffice$ = ""
     
       sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idmanager='" + ID_manager$ + "' and idoffice='" + id_office$ + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       id_moneyreportoffice$ = Rs(0)
       Rs.Close
     
    
       If id_moneyreportoffice$ = "" Then
           ' btnerase_archivo.Enabled = True
           btnborrar_archivo2.Enabled = True
           btnsave.Enabled = True
           chk_dayoff.Enabled = True
           btnsave.Picture = img_disk_up.Picture
           btnsend.Enabled = True
           btnsend.Picture = img_send_up.Picture
           msg.Visible = False
           Exit Sub
   
       End If
    
   
   
       sSelect = "SELECT submitted From moneyreportbyoffice where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"
       ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
       Rs.Open sSelect, base, adOpenUnspecified
       submitido = Rs(0)
       Rs.Close
          
    
       If submitido = "True" Then
            ' btnerase_archivo.Enabled = False
            btnborrar_archivo2.Enabled = False
            chk_dayoff.Enabled = False
            btnsave.Enabled = False
            btnsave.Picture = img_caja.Picture
            btnsend.Enabled = False
            btnsend.Picture = img_caja.Picture
      
      Else
            ' btnerase_archivo.Enabled = True
            btnborrar_archivo2.Enabled = True
            chk_dayoff.Enabled = True
            btnsave.Enabled = True
            btnsave.Picture = img_disk_up.Picture
            btnsend.Enabled = True
            btnsend.Picture = img_send_up.Picture
       
      End If
    
   
   
  End If


 msg.Visible = False

End Sub

Public Sub enca3()

On Error Resume Next


grid3.cols = grid3.cols + 1

grid3.ColWidth(0) = 600
grid3.ColAlignment(0) = flexAlignLeftCenter


grid3.ColWidth(1) = 1000 'idreceiptHDR
grid3.ColAlignment(1) = flexAlignRightCenter

grid3.ColWidth(2) = 2000   ' Date
grid3.ColAlignment(2) = flexAlignLeftCenter

grid3.ColWidth(3) = 900   ' Idcustomer
grid3.ColAlignment(3) = flexAlignCenterCenter

grid3.ColWidth(4) = 3200   ' Name
grid3.ColAlignment(4) = flexAlignLeftCenter

grid3.ColWidth(5) = 2200   'Policynumber
grid3.ColAlignment(5) = flexAlignLeftCenter

grid3.ColWidth(6) = 1200   ' idcompany
grid3.ColAlignment(6) = flexAlignCenterCenter

grid3.ColWidth(7) = 2000   ' companyname
grid3.ColAlignment(7) = flexAlignLeftCenter

grid3.ColWidth(8) = 1200   ' idemployee
grid3.ColAlignment(8) = flexAlignCenterCenter

grid3.ColWidth(9) = 1600   ' USR
grid3.ColAlignment(9) = flexAlignLeftCenter

grid3.ColWidth(10) = 1600   ' CSR
grid3.ColAlignment(10) = flexAlignLeftCenter

grid3.ColWidth(11) = 800   ' IdOffice
grid3.ColAlignment(11) = flexAlignCenterCenter



  grid3.ColWidth(12) = 1840   ' Office
  grid3.ColAlignment(12) = flexAlignLeftCenter
  grid3.ColWidth(13) = 1150  ' Fiduciary  1200
  grid3.ColAlignment(13) = flexAlignRightCenter
  
  

  
  grid3.ColWidth(14) = 1300   ' total amount receipt
  grid3.ColAlignment(14) = flexAlignRightCenter
  
  grid3.ColWidth(15) = 1300  ' Amount paid
  grid3.ColAlignment(15) = flexAlignRightCenter
  
  grid3.ColWidth(16) = 1500   ' Payment Method
  grid3.ColAlignment(16) = flexAlignRightCenter

  
    grid3.ColWidth(17) = 1300   ' balance due
  grid3.ColAlignment(17) = flexAlignRightCenter

  grid3.ColWidth(18) = 1200   ' balance due date
  grid3.ColAlignment(18) = flexAlignLeftCenter
  
  grid3.ColWidth(19) = 10   '
  grid3.ColAlignment(19) = flexAlignCenterCenter

 grid3.ColWidth(20) = 10   '
  grid3.ColAlignment(19) = flexAlignCenterCenter

grid3.Row = 0

grid3.Col = 1
grid3.Text = "Receipt#"

grid3.Col = 2
grid3.Text = "Date"

grid3.Col = 3
grid3.Text = "IdCust"

grid3.Col = 4
grid3.Text = "Name"

grid3.Col = 5
grid3.Text = "Policy#"

grid3.Col = 6
grid3.Text = "IdCompany"

grid3.Col = 7
grid3.Text = "Company"


grid3.Col = 8
grid3.Text = "IdEmployee"


grid3.Col = 9
grid3.Text = "USR"


grid3.Col = 10
grid3.Text = "CSR"

grid3.Col = 11
grid3.Text = "IdOffice"

  grid3.Col = 12
  grid3.Text = "Office"
  
  grid3.Col = 13
  grid3.Text = "Fiduciary"
  
  grid3.Col = 14
  grid3.Text = "Total Receipt"
  grid3.Col = 15
  grid3.Text = "Amount Paid"
  
  grid3.Col = 16
  grid3.Text = "PYMT Method"
  
  
  grid3.Col = 17
  grid3.Text = "Balance Due"
  grid3.Col = 18
  grid3.Text = "BD Date"
    
    
  



grid3.FixedRows = 1
grid3.FixedCols = 1

grid3.Row = 1
grid31.Col = 1
End Sub

Public Sub calcula_total_LAE()
On Error Resume Next
Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset

  id_employee = Val(Right(Form1.cbo_agentes.List(Form1.cbo_agentes.ListIndex), 20))
  ID_manager = Val(Right(Form1.cbo_managers.List(Form1.cbo_managers.ListIndex), 20))
  lae_office$ = RTrim(LTrim(Right(Form1.cbo_oficina.List(Form1.cbo_oficina.ListIndex), 25)))
 
  lae_office$ = lbloficina_agente.Caption
  
  
  
  If (agente$ = manager$ And agente$ <> "") Or (id_employee = ID_manager And id_employee > 0) Then
    estado_carga = 1
  Else
    estado_carga = 0
  End If
 
 
    
    
    sSelect = "select idoffice from officescatalog where office='" + lae_office$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_oficina$ = Rs(0)
    Rs.Close
    
    
    
    
    
    


     fecha_de_revision$ = lbldate_agente.Caption
 
     sSelect = "SELECT rechdr.[IdReceiptHDR],rechdr.Date ,rechdr.[IdCustomer], CONCAT(cus.FirstName+' ',cus.MiddleName,+' '+cus.LastName1,+' '+cus.LastName2) as [Name] " & _
                          ",polhdr.PolicyNumber ,polhdr.IdCompany ,ins.CompanyName ,emp.IDEmployee, emp.Username as USR, csr.Username as CSR, rechdr.IdOffice, ofc.Office " & _
                          ",rechdr.Fiduciary, rechdr.TotalAmntReceipt, rechdr.AmountPaid, PaymentMethod=STUFF((SELECT DISTINCT ', ' + CAST(t3.PayMethodName AS VARCHAR(MAX)) " & _
                          "FROM ReceiptsPayments t2 " & _
                          "join PayMethodCatalog t3 on t3.IdPayMethod=t2.IdPayMethod " & _
                          "Where t2.IdReceiptHDR = rechdr.IdReceiptHDR " & _
                          "FOR XML PATH('') " & _
                          "),1,1,'') ,rechdr.BalanceDue ,rechdr.BalanceDueDate " & _
                          "from  [ReceiptsHDR] rechdr " & _
                          "inner join PoliciesHDR polhdr on polhdr.IdPoliciesHDR=rechdr.IdPoliciesHDR " & _
                          "inner join EmployeeInfo emp on emp.IDEmployee=rechdr.IdEmployeeUSR " & _
                          "inner join OfficesCatalog ofc on ofc.IdOffice=rechdr.IdOffice " & _
                          "inner join EmployeeInfo csr on csr.IDEmployee=rechdr.IdEmployeeCSR1 " & _
                          "inner join InsuranceCatalog ins on ins.IdCompany=polhdr.IdCompany " & _
                          "inner join Customers cus on cus.IdCustomer=polhdr.IdCustomer " & _
                          "where rechdr.Active=1 and cast(rechdr.Date as Date) >= '" + fecha_de_revision$ + "' " & _
                          "AND cast( rechdr.DATE as Date) <= '" + fecha_de_revision$ + "' and emp.IDEmployee='" + Format(id_employee, "####0") + "' and rechdr.IdOffice='" + id_oficina$ + "'"
                          'order by Office
 
 
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    suma_total_lae = 0
    For t = 1 To Grid2.Rows - 1
      Grid2.Row = t
      Grid2.Col = 15
      valor = Val(Grid2.Text)
      suma_total_lae = suma_total_lae + valor
    Next t
    
    txttotal_LAE_agente.Text = Format(suma_total_lae, "$###,##0.00")
    
    
    
End Sub


Public Sub limpia_tarjetas()
On Error Resume Next
For t = 0 To 9
  img_credit(t).Picture = LoadPicture()
  img_debito(t).Picture = LoadPicture()
Next t

End Sub

Public Sub carga_datos()
On Error Resume Next
msg.Visible = True
msg.Refresh

Dim sSelect As String
    
Dim Rs As ADODB.Recordset

Set Rs = New ADODB.Recordset
           
  
  id_employee = Val(Right(Form1.cbo_agentes.List(Form1.cbo_agentes.ListIndex), 20))
  
  If Form1.cbo_managers.ListIndex = -1 Then Form1.cbo_managers.ListIndex = 0
  
  ID_manager = Val(Right(Form1.cbo_managers.List(Form1.cbo_managers.ListIndex), 20))
  lae_office$ = RTrim(LTrim(Right(Form1.cbo_oficina.List(Form1.cbo_oficina.ListIndex), 25)))
 
  
  If (agente$ = manager$ And agente$ <> "") Or (id_employee = ID_manager And id_employee > 0) Then
    estado_carga = 1
  Else
    estado_carga = 0
  End If
 
  Grid2.Visible = False
  
  Grid2.Clear
    
    
    sSelect = "select idoffice from officescatalog where office='" + lae_office$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_oficina$ = Rs(0)
    Rs.Close
    
    
    
    
     
    fecha_de_revision$ = lbldate_agente.Caption
        
     
    ' obtiene el ID_moneyreport
    
    
    id_moneyreport$ = ""
    sSelect = "select idmoneyreport from moneyreport where idoffice='" + id_oficina$ + "' and idemployee='" + Format(id_employee, "###0") + "' and datereport=convert(datetime, '" + lbldate_agente.Caption + " ')"
    'sSelect = "select idmoneyreport, idoffice from moneyreport where idemployee='" + Format(id_employee, "###0") + "' and datereport=convert(datetime, '" + lbldate_agente.Caption + " ')"
    Rs.Open sSelect, base, adOpenUnspecified
    id_moneyreport$ = Rs(0)
    Rs.Close
    
    
    ' checa la oficina y la actualiza en la barra
    sSelect = "select office from officescatalog where idoffice='" + id_oficina$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    oficina$ = Rs(0)
    Rs.Close
    lbloficina_agente.Caption = oficina$
    
    lbloficina_agente.Caption = oficina_trabajada$
    
    
    ' obtiene la firma
    sSelect = "select submitted from moneyreport where idoffice='" + id_oficina$ + "' and idemployee='" + Format(id_employee, "###0") + "' and datereport=convert(datetime, '" + lbldate_agente.Caption + " ')"
    ' sSelect = "select submitted from moneyreport where idemployee='" + Format(id_employee, "###0") + "' and datereport=convert(datetime, '" + lbldate_agente.Caption + " ')"
    Rs.Open sSelect, base, adOpenUnspecified
    submitido = Rs(0)
    Rs.Close
    
    
    
    chk_dayoff.Visible = Not submitido
    
    
    
    
    ' obtiene si descansa
    
    sSelect = "select dayoff from moneyreport where idoffice='" + id_oficina$ + "' and idemployee='" + Format(id_employee, "###0") + "' and datereport=convert(datetime, '" + lbldate_agente.Caption + " ')"
    ' sSelect = "select dayoff from moneyreport where idemployee='" + Format(id_employee, "###0") + "' and datereport=convert(datetime, '" + lbldate_agente.Caption + " ')"
    Rs.Open sSelect, base, adOpenUnspecified
    dialibre = Rs(0)
    Rs.Close
    
    
    If dialibre = 0 Then
       msgdescanso.Visible = False
       chk_dayoff.Value = False
    Else
       msgdescanso.Visible = True
       chk_dayoff.Value = 1
    End If
    
    
    
    
    
    
    
    sSelect = "select firstname, lastname1 from employeeinfo where username='" + LTrim(RTrim(Left(cbo_agentes.List(cbo_agentes.ListIndex), Len(cbo_agentes.List(cbo_agentes.ListIndex)) - 15))) + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    nombre$ = Rs(0)
    apellido$ = Rs(1)
    Rs.Close
           
    full_name_agente$ = nombre$ + " " + apellido$
   
   
    
    If submitido = 1 Or submitido = True Then
      chk_firma_agente.Value = 1
      Firma_agente.Caption = full_name_agente$
    Else
      chk_firma_agente.Value = 0
      Firma_agente.Caption = ""
    End If
    
    
    
    ' obtiene si fue revisado
    
    
    
    
    
    ' carga las notas
    
    sSelect = "select notes from moneyreport where idmoneyreport='" + id_moneyreport$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    txtnotas_agente.Text = RTrim(LTrim(Rs(0)))
    Rs.Close
    
    
    
    
    sSelect = "select * from moneyreportpaymetrel where idmoneyreport='" + id_moneyreport$ + "'"
     
    
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    ' acomoda grid y encabezados
    
    'grid1.cols = grid1.cols - 1
    
    
   If Grid2.Rows <= 1 Then
     
   End If
     
    
     Dim total_dinero(10) As Single
     
      
      contador = 0
      contador2 = 0
      contador3 = 1
      contador4 = 3
      
      conta_cash = 0
      conta_check = 0
      conta_coins = 0
      conta_money = 0
      
      
      valido1 = 66
      
     For t = 1 To Grid2.Rows - 1
       Grid2.Row = t
       
       Grid2.Col = 4
       cantidad = Val(Grid2.Text)
       
       Grid2.Col = 3
       metodo$ = LTrim(RTrim(UCase(Grid2.Text)))
       
       Select Case metodo$
       Case "4"  '"CREDIT CARD"
                    
          txtcredit_agente(contador).Text = Str(cantidad)
          contador = contador + 1
          
       Case "2" ' "CASH"
       
          conta_cash = conta_cash + cantidad
          
       
       
       Case "5"  '"DEBIT"
       
          txtdebit_agente(contador2).Text = Str(cantidad)
          contador2 = contador2 + 1
       
                 
       Case "10" '"CHECK"
       
          txtcash(contador4).Text = Str(cantidad)
          contador4 = contador4 + 1
       
       Case "12"  '"COINS"
       
          conta_coins = conta_coins + cantidad
       
       Case "9"  '"MONEY ORDER"
       
          txtcash(contador3).Text = Str(cantidad)
          contador3 = contador3 + 1
      
          
       End Select
       
       
              
              
       
     Next t
     
      
     txtcash(0).Text = Str(conta_cash)
     txtcash(5).Text = Str(conta_coins)
     
    
     
     valido1 = 0
     
     ' calcula totales
     tot_credit = 0
     tot_debit = 0
     
     For t = 0 To 19
        tot_debit = tot_debit + Val(txtdebit_agente(t).Text)
        tot_credit = tot_credit + Val(txtcredit_agente(t).Text)
        

     Next t
     
     
     
For Y = 0 To 19
     img_credit(Y).Picture = LoadPicture()
      img_debito(Y).Picture = LoadPicture()

For t = 1 To grid1.Rows - 1
   grid1.Row = t
   
   grid1.Col = 15
   cantidad_pagada$ = grid1.Text
   
   
   grid1.Col = 16
   metodo$ = grid1.Text
   
   
   
   If Val(txtcredit_agente(Y).Text) = Val(cantidad_pagada$) Then
      
      
      Select Case UCase(RTrim(LTrim(metodo$)))

      Case "VISA"
           img_credit(Y).Picture = visa.Picture
      Case "MASTERCARD"
        img_credit(Y).Picture = master.Picture
      Case "AMEX"
      img_credit(Y).Picture = american.Picture
      Case "DISCOVER"
      img_credit(Y).Picture = discover.Picture
      Case "DEBIT"
      img_credit(Y).Picture = debito.Picture
      End Select
   End If
   
   
   If Val(txtdebit_agente(Y).Text) = Val(cantidad_pagada$) Then
     
      img_debito(Y).Picture = debito.Picture
     
   End If
   
   
Next t

Next Y


' carga los recibos VOID

     
    Grid2.Clear
    sSelect = "select * from moneyreportrecvoidrel where idmoneyreport='" + id_moneyreport$ + "'"
            
       ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
        
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    f2$ = lbldate_agente.Caption
    
    
    If Grid2.Rows > 1 Then
    
    
        sSelect = "select IDReceiptHDR, IdCustomer, AmountPaid from ReceiptsHDR " & _
        "Where Void = 1 and cast(Date as Date) >= '" + f2$ + "' AND cast( DATE as Date) <= '" + f2$ + "' " & _
        "and IdEmployeeUSR='" + Format(id_employee, "#####0") + "' order by Date"
    
        ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
        Rs.Open sSelect, base, adOpenUnspecified
        
       ' Permitir redimensionar las columnas
        Grid2.AllowUserResizing = flexResizeColumns

      ' Asignar el recordset al FlexGrid
        Set Grid2.DataSource = Rs
                                
        Rs.Close
    
    
       
       For t = 1 To Grid2.Rows - 1
             Grid2.Row = t
             Grid2.Col = 1
             id_receiptHDR_LAE$ = Grid2.Text
             
             Grid2.Col = 2
             id_customer_LAE$ = Grid2.Text
             
             Grid2.Col = 3
             cantidad_pagada_LAE$ = Grid2.Text
             
    
             txtcustomer_agente(t - 1).Text = id_customer_LAE$
             txtrecibos_agente(t - 1).Text = id_receiptHDR_LAE$
             txtamount_agente(t - 1).Text = cantidad_pagada_LAE$
                    
             
    
       Next t
       
    End If
    
    

     
     lbltotal_debit_agent.Caption = Format(tot_debit, "$###,##0.00")
     lbltotal_credit_agent.Caption = Format(tot_credit, "$###,##0.00")
     
     lbltotal_debit_credit_agent.Caption = Format(tot_credit + tot_debit, "$###,##0.00")
     lblgrantotal_debitcredit_agente.Caption = lbltotal_debit_credit_agent.Caption
     

     tot_cash = 0
     For t = 0 To 5
       tot_cash = tot_cash + Val(txtcash(t).Text)
     Next t
     
     lbltotal_cash_agente.Caption = Format(tot_cash, "$###,##0.00")
     lblgrantotal_cash_agente.Caption = lbltotal_cash_agente.Caption
     
     
     
     
     
  
     
     

 msg.Visible = False

End Sub



Public Sub carga_combo_de_empleados()
On Error Resume Next

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
    
    
    Set Rs = New ADODB.Recordset
    
For Y = 0 To 5
cbo_employees(Y).Clear

Next Y



Grid2.Clear

sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (3,6,16,17,18,2,24,37)  and empjob.Active='1'"

    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    

  
' load agents to Combo_agents


For t = 1 To Grid2.Rows - 1
  Grid2.Row = t
  Grid2.Col = 1
  id_agente$ = Grid2.Text
  
  Grid2.Col = 2
  
  agente$ = Grid2.Text
  
' carga el nombre completo

       sSelect = "select firstname, lastname1 from employeeinfo where username='" + agente$ + "'"
       Rs.Open sSelect, base, adOpenUnspecified
       nombre$ = Rs(0)
       apellido$ = Rs(1)
       Rs.Close
       
       ' verifica que no este en la lista
       existe = 0
       n$ = nombre$ + " " + apellido$
       For Y = 0 To cbo_employees(0).ListCount - 1
          n2$ = RTrim(Left(cbo_employees(0).List(Y), Len(cbo_employees(0).List(Y)) - 32))
          If n$ = n2$ Then
             existe = 1
             Exit For
          End If
       Next Y
       
       If existe = 0 Then
          cbo_employees(0).AddItem nombre$ + " " + apellido$ + Space(30) + Format(id_agente$, "00")
       End If
  
Next t


' carga



For Y = 0 To cbo_employees(0).ListCount - 1
  For z = 1 To 5
     cbo_employees(z).AddItem cbo_employees(0).List(Y)
  Next z
Next Y





End Sub


Public Sub calcula_total_oficina_manager()
        On Error Resume Next
        btnupdate_LAE_Click
        
  ' calcula total txtcant_ida
        total_txtcant_ida = 0
        For Y = 0 To 2
           total_txtcant_ida = total_txtcant_ida + Val(txtcant_ida(Y).Text)
        Next Y
        
        lbltotal_agentes_idos.Caption = Format(total_txtcant_ida, "$###,##0.00")
        
        
        
    ' calcula total txtcant_venida
        total_txtcant_venida = 0
        For Y = 0 To 2
           total_txtcant_venida = total_txtcant_venida + Val(txtcant_venida(Y).Text)
        Next Y
        
        lbltotal_agentes_que_vinieron.Caption = Format(total_txtcant_venida, "$###,##0.00")
        
        
        
        ' calcula total cash_manager
        total_cash_manager = 0
        For Y = 0 To 2
           total_cash_manager = total_cash_manager + Val(txtdinero(Y).Text)
        Next Y
        
        lbltotal_cash_manager.Caption = Format(total_cash_manager, "$###,##0.00")
        lbltotal_needed_oficina.Caption = lbltotal_cash_manager.Caption
        
        
        
        ' calcula total tarjetas de oficina
        total_debito_credito_manager = Val(txtdebit_manager.Text) + Val(txtcredit_manager.Text)
        lbltotal_debito_credito_manager.Caption = Format(total_debito_credito_manager, "$###,##0.00")
                
        lbltotal_debit_and_credit_oficina.Caption = lbltotal_debito_credito_manager.Caption
        
        
        
        
        lbltotal_dejado_por_agentes_de_oficina.Caption = lbltotal_agentes_idos.Caption  ' -
        lbltotal_dejado_por_agentes_que_vinieron.Caption = lbltotal_agentes_que_vinieron.Caption   ' +
        
        
        
        ' empiezan los calculos aqui
        ' ***************************
        dinero_en_efectivo = Val(Format(lbltotal_cash_manager.Caption, "00000.00"))           ' +
        dinero_en_tarjetas = Val(Format(lbltotal_debito_credito_manager.Caption, "00000.00")) ' +
        dinero_que_se_va = Val(Format(lbltotal_agentes_idos.Caption, "00000.00"))             ' +
        dinero_que_llega = Val(Format(lbltotal_agentes_que_vinieron.Caption, "00000.00"))     ' -
        
        total_recaudado = dinero_en_efectivo + dinero_en_tarjetas - dinero_que_llega + dinero_que_se_va
        
        
        
        grand_Total_Oficina = total_recaudado - Val(txttotal_venta_manager.Text)
        
        
        lblover_short_oficina.Caption = Format(grand_Total_Oficina, "$###,##0.00")
        
        
        
        
        If total_recaudado > Val(txttotal_venta_manager.Text) Then
           
           arrow(0).Visible = True
           arrow(1).Visible = False
        ElseIf total_recaudado < Val(txttotal_venta_manager.Text) Then
           
           arrow(0).Visible = False
           arrow(1).Visible = True
        Else
           
           arrow(0).Visible = False
           arrow(1).Visible = False
        End If
        
        
       
        
        
End Sub

Public Sub carga_totales_oficina()
On Error Resume Next

Grid2.Visible = True
Dim sSelect As String
    
Dim Rs As ADODB.Recordset

Set Rs = New ADODB.Recordset
 ' verifica si ya existe el reporte de la oficina
     
     
       id_employee = Val(Right(Form1.cbo_agentes.List(Form1.cbo_agentes.ListIndex), 20))
  ID_manager = Val(Right(Form1.cbo_managers.List(Form1.cbo_managers.ListIndex), 20))
  lae_office$ = RTrim(LTrim(Right(Form1.cbo_oficina.List(Form1.cbo_oficina.ListIndex), 25)))


     
      sSelect = "select idoffice from officescatalog where office='" + lae_office$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_oficina$ = Rs(0)
    Rs.Close
    
     
     sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idoffice='" + id_oficina$ + "'"
     Rs.Open sSelect, base, adOpenUnspecified
     id_moneyreportoffice$ = Rs(0)
     Rs.Close
     
     
    Grid2.Clear
    sSelect = "select * from MoneyReportByOffice where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"
            
       ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
        
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    
     
     ' carga la hoja de MANAGER
     ' ***********************************************************************************************************************************************
     ' ***********************************************************************************************************************************************
     

       
    
    If Grid2.Rows > 1 Then
        
        For t = 1 To Grid2.Rows - 1
            Grid2.Row = t
            valido1 = 66
            Grid2.Col = 8
            txtnotas_manager.Text = Grid2.Text
            
            Grid2.Col = 9
            txtdinero(0).Text = Grid2.Text
            
            Grid2.Col = 10
            txtdinero(1).Text = Grid2.Text
            
            Grid2.Col = 11
            txtdinero(2).Text = Grid2.Text
            
            Grid2.Col = 12
            txtdebit_manager.Text = Grid2.Text
            
            Grid2.Col = 13
           txtcredit_manager.Text = Grid2.Text
            
            ' EMPLEADO1
            Grid2.Col = 15
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
            
                       
            cbo_employees(0).ListIndex = -1
           
               For Y = 0 To cbo_employees(0).ListCount - 1
                    empid = Val(Right(cbo_employees(0).List(Y), 10))
                    If empid = Val(r$) Then
                        cbo_employees(0).ListIndex = Y
                        Exit For
                    End If
               
               Next Y
            
            
            
            ' OFICINA
            Grid2.Col = 16
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
                      
            sSelect = "select Office from officescatalog where idoffice='" + r$ + "'"
            Rs.Open sSelect, base, adOpenUnspecified
            Office0$ = UCase(Rs(0))
            Rs.Close
            
            cbooficina1(0).ListIndex = -1
            If LTrim(RTrim(Office0$)) <> "" Then
               
               
               For Y = 0 To cbooficina1(0).ListCount - 1
                    n$ = RTrim(LTrim(UCase(Right(cbooficina1(0).List(Y), 20))))
                    
                    If n$ = Office0$ Then
                        cbooficina1(0).ListIndex = Y
                        Exit For
                    End If
               Next Y
            End If
            
            
              
            Grid2.Col = 17
            txtcant_ida(0).Text = LTrim(RTrim(UCase(Grid2.Text)))
            
            ' =================================
            
            
             ' EMPLEADO2
            Grid2.Col = 18
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
            
                       
            cbo_employees(1).ListIndex = -1
           
               For Y = 0 To cbo_employees(1).ListCount - 1
                    empid = Val(Right(cbo_employees(1).List(Y), 10))
                    If empid = Val(r$) Then
                        cbo_employees(1).ListIndex = Y
                        Exit For
                    End If
               
               Next Y
            
            
            
            ' OFICINA
            Grid2.Col = 19
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
                      
            sSelect = "select Office from officescatalog where idoffice='" + r$ + "'"
            Rs.Open sSelect, base, adOpenUnspecified
            Office1$ = UCase(Rs(0))
            Rs.Close
            
            cbooficina1(1).ListIndex = -1
            If LTrim(RTrim(Office1$)) <> "" Then
               
               
               For Y = 0 To cbooficina1(1).ListCount - 1
                    n$ = RTrim(LTrim(UCase(Right(cbooficina1(1).List(Y), 20))))
                    
                    If n$ = Office1$ Then
                        cbooficina1(1).ListIndex = Y
                        Exit For
                    End If
               Next Y
            End If
            
            
              
            Grid2.Col = 20
            txtcant_ida(1).Text = LTrim(RTrim(UCase(Grid2.Text)))
            
          
            ' =================================
            
            
             ' EMPLEADO3
            Grid2.Col = 21
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
            
                       
            cbo_employees(2).ListIndex = -1
           
               For Y = 0 To cbo_employees(2).ListCount - 1
                    empid = Val(Right(cbo_employees(2).List(Y), 10))
                    If empid = Val(r$) Then
                        cbo_employees(2).ListIndex = Y
                        Exit For
                    End If
               
               Next Y
            
            
            
            ' OFICINA
            Grid2.Col = 22
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
                      
            sSelect = "select Office from officescatalog where idoffice='" + r$ + "'"
            Rs.Open sSelect, base, adOpenUnspecified
            Office2$ = UCase(Rs(0))
            Rs.Close
            
            cbooficina1(2).ListIndex = -1
            If LTrim(RTrim(Office2$)) <> "" Then
               
               
               For Y = 0 To cbooficina1(2).ListCount - 1
                    n$ = RTrim(LTrim(UCase(Right(cbooficina1(2).List(Y), 20))))
                    
                    If n$ = Office2$ Then
                        cbooficina1(2).ListIndex = Y
                        Exit For
                    End If
               Next Y
            End If
            
            
              
            Grid2.Col = 23
            txtcant_ida(2).Text = LTrim(RTrim(UCase(Grid2.Text)))
            
            
             ' =================================
             ' SEGUNDA PARTE
             
            
             ' EMPLEADO4
            Grid2.Col = 24
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
            
                       
            cbo_employees(3).ListIndex = -1
           
               For Y = 0 To cbo_employees(3).ListCount - 1
                    empid = Val(Right(cbo_employees(3).List(Y), 10))
                    If empid = Val(r$) Then
                        cbo_employees(3).ListIndex = Y
                        Exit For
                    End If
               
               Next Y
            
            
            
            ' OFICINA
            Grid2.Col = 25
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
                      
            office$ = ""
            sSelect = "select Office from officescatalog where idoffice='" + r$ + "'"
            Rs.Open sSelect, base, adOpenUnspecified
            office$ = UCase(Rs(0))
            Rs.Close
            
            cbooficina1(3).ListIndex = -1
            If LTrim(RTrim(office$)) <> "" Then
               
               
               For Y = 0 To cbooficina1(3).ListCount - 1
                    n$ = RTrim(LTrim(UCase(Right(cbooficina1(3).List(Y), 20))))
                    
                    If n$ = office$ Then
                        cbooficina1(3).ListIndex = Y
                        Exit For
                    End If
               Next Y
            End If
            
            
              
            Grid2.Col = 26
            txtcant_venida(0).Text = LTrim(RTrim(UCase(Grid2.Text)))
            
            ' =================================
            
            
           ' EMPLEADO5
            Grid2.Col = 27
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
            
                       
            cbo_employees(4).ListIndex = -1
           
               For Y = 0 To cbo_employees(4).ListCount - 1
                    empid = Val(Right(cbo_employees(4).List(Y), 10))
                    If empid = Val(r$) Then
                        cbo_employees(4).ListIndex = Y
                        Exit For
                    End If
               
               Next Y
            
            
            
            ' OFICINA
            Grid2.Col = 28
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
                      
            office$ = ""
            sSelect = "select Office from officescatalog where idoffice='" + r$ + "'"
            Rs.Open sSelect, base, adOpenUnspecified
            office$ = UCase(Rs(0))
            Rs.Close
            
            cbooficina1(4).ListIndex = -1
            If LTrim(RTrim(office$)) <> "" Then
               
               
               For Y = 0 To cbooficina1(4).ListCount - 1
                    n$ = RTrim(LTrim(UCase(Right(cbooficina1(4).List(Y), 20))))
                    
                    If n$ = office$ Then
                        cbooficina1(4).ListIndex = Y
                        Exit For
                    End If
               Next Y
            End If
            
            
              
            Grid2.Col = 29
            txtcant_venida(1).Text = LTrim(RTrim(UCase(Grid2.Text)))
            
            ' =================================
            
            
             ' EMPLEADO6
             
            Grid2.Col = 30
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
            
                       
            cbo_employees(5).ListIndex = -1
           
               For Y = 0 To cbo_employees(5).ListCount - 1
                    empid = Val(Right(cbo_employees(5).List(Y), 10))
                    If empid = Val(r$) Then
                        cbo_employees(5).ListIndex = Y
                        Exit For
                    End If
               
               Next Y
            
            
            
            ' OFICINA
            Grid2.Col = 31
            r$ = LTrim(RTrim(UCase(Grid2.Text)))
            office$ = ""
                      
            sSelect = "select Office from officescatalog where idoffice='" + r$ + "'"
            Rs.Open sSelect, base, adOpenUnspecified
            office$ = UCase(Rs(0))
            Rs.Close
            
            cbooficina1(5).ListIndex = -1
            If LTrim(RTrim(office$)) <> "" Then
               
               
               For Y = 0 To cbooficina1(5).ListCount - 1
                    n$ = RTrim(LTrim(UCase(Right(cbooficina1(5).List(Y), 20))))
                    
                    If n$ = office$ Then
                        cbooficina1(5).ListIndex = Y
                        Exit For
                    End If
               Next Y
            End If
            
            
              
            Grid2.Col = 32
            txtcant_venida(2).Text = LTrim(RTrim(UCase(Grid2.Text)))
            
            ' =================================
            
            
            Grid2.Col = 33
            If Grid2.Text = "True" Then Grid2.Text = "1"
            chkmanager.Value = Grid2.Text
            
        Next t
        
        
        
        
        calcula_total_oficina_manager
        
      
        
    
    End If
    
    
    
     ' obtiene el ID_moneyreport
    
    
    id_moneyreportoffice$ = ""
    sSelect = "select idmoneyreportoffice from moneyreportbyoffice where idoffice='" + id_oficina$ + "' and datereport=convert(datetime, '" + lbldate_agente.Caption + " ')"
    Rs.Open sSelect, base, adOpenUnspecified
    id_moneyreportoffice$ = Rs(0)
    Rs.Close
    
    ' carga los archivos
    
    ListView2.ListItems.Clear
    
    If id_moneyreportoffice$ = "" Then Exit Sub
    
    ruta_archivos$ = "\\192.168.84.215\moneyreport\O-" + id_moneyreportoffice$ + "\"

   


File1.Path = "c:\"
File1.Path = ruta_archivos$
 For t = 0 To File1.ListCount - 1
    
    If UCase(Right(File1.List(t), 3)) = "PDF" Then
       Set list_item = ListView2.ListItems.Add(, , File1.List(t))
       list_item.Icon = 1
       list_item.SmallIcon = 1
       list_item.SubItems(1) = File1.List(t)
    Else
       Set list_item = ListView2.ListItems.Add(, , File1.List(t))
       list_item.Icon = 2
       list_item.SmallIcon = 2
       list_item.SubItems(1) = File1.List(t)
    End If
                
         

    
 Next t
    
    
    valido1 = 0
End Sub


Public Sub arrastra_archivo2(nombre_archivo As String)
On Error Resume Next

 ' See if we know what to do with the data.

   

   
        ' See if this is a file name ending in

        ' bmp, gif, jpg, pdf or jpeg.

        extension1$ = LCase$(Right$(nombre_archivo, 4))
        
        guarda$ = ""
        For Y = Len(nombre_archivo) To 1 Step -1
           If Mid$(nombre_archivo, Y, 1) = "\" Then
               Exit For
           Else
              guarda$ = Mid$(nombre_archivo, Y, 1) + guarda$
           End If
        Next Y
        
        nombre_archivo$ = guarda$
        
        Kill "c:\money\" + nombre_archivo$
         
        Select Case extension1$
        Case ".pdf"
        
                If Dir$("c:\money\" + guarda$) <> "" Then
                   GoTo brinca
                End If
    
               For z = 1 To ListView2.ListItems.Count
                                 
                   a$ = ListView2.ListItems.Item(z)
                   If a$ = guarda$ Then
                       ListView2.ListItems.Remove (z)
                       Exit For
                  End If
                
                Next z
                
                Set list_item = ListView2.ListItems.Add(, , guarda$)
                list_item.Icon = 1
                list_item.SmallIcon = 1
                list_item.SubItems(1) = nombre_archivo
                'list_item.SubItems(2) = "0-471-24267-5"
                
            Effect = vbDropEffectCopy
            
            FileCopy nombre_archivo, "c:\money\" + guarda$
            
            
        Case ".bmp", ".png", ".jpg", "jpeg"

            ' Load the file.

        '    picDragTo(Index).Picture = LoadPicture(Data.Files(1))
        
                If Dir$("c:\money\" + guarda$) <> "" Then
                   GoTo brinca
                End If
    
        
           Set list_item = ListView2.ListItems.Add(, , guarda$)
                list_item.Icon = 2
                list_item.SmallIcon = 2
                list_item.SubItems(1) = nombre_archivo
                'list_item.SubItems(2) = "0-471-24267-5"
                
         

            Effect = vbDropEffectCopy
            
            FileCopy nombre_archivo, "c:\money\" + guarda$
         
            
       
        Case Else

            ' Tell the source we did nothing.

            Effect = vbDropEffectNone

        End Select
    


    
brinca:
End Sub

Public Sub carga_archivos2()
Dim column_header As ColumnHeader
Dim list_item As ListItem

msg.Visible = True



    ' Create the column headers.
    Set column_header = ListView2. _
        ColumnHeaders.Add(, , "Abbrev", _
        TextWidth("Abbrev"))
    Set column_header = ListView2. _
        ColumnHeaders.Add(, , "Title", _
        TextWidth("Ready-to-Run Visual Basic Algorithms"))
    Set column_header = ListView2. _
        ColumnHeaders.Add(, , "ISBN", _
        TextWidth("0-000-00000-0"))

    ' Start with report view.
    ' mnuViewChoice_Click lvwReport
     ListView2.View = 0

    ' Associate the ImageLists with the
    ' ListView's Icons and SmallIcons properties.
    ListView2.Icons = imgLarge
    ListView2.SmallIcons = imgSmall
    
     msg.Visible = False

    
End Sub

Public Sub carga_revisado()
On Error Resume Next

Dim sSelect As String
    
Dim Rs As ADODB.Recordset

Set Rs = New ADODB.Recordset
           
  
  id_employee = Val(Right(Form1.cbo_agentes.List(Form1.cbo_agentes.ListIndex), 20))
  ID_manager = Val(Right(Form1.cbo_managers.List(Form1.cbo_managers.ListIndex), 20))
  lae_office$ = RTrim(LTrim(Right(Form1.cbo_oficina.List(Form1.cbo_oficina.ListIndex), 25)))
 
    
    
    sSelect = "select idoffice from officescatalog where office='" + lae_office$ + "'"
    Rs.Open sSelect, base, adOpenUnspecified
    id_oficina$ = Rs(0)
    Rs.Close
    
     
    fecha_de_revision$ = lbldate_agente.Caption
        
     
    ' obtiene el ID_moneyreport
    
    
    id_moneyreport$ = ""
      
    

 Dim revisado As Boolean
    
    
    
If tipo_guardado = 1 Then
    
         
        ' verifica si ya existe el reporte
     
        sSelect = "select idmoneyreport from moneyreport where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idemployee='" + Str(id_employee) + "' and idoffice='" + id_oficina$ + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        id_moneyreport$ = Rs(0)
        Rs.Close
     
        If id_moneyreport$ = "" Then GoTo saltado
        
        
        sSelect = "select reviewed from moneyreport where idmoneyreport='" + id_moneyreport$ + "'"
                 
        Rs.Open sSelect, base, adOpenUnspecified
        revisado = Rs(0)
        Rs.Close
   
        If revisado = "False" Then
             chk_revisado.Value = 0
        Else
             chk_revisado.Value = 1
        End If
        
        
               
         
     
     
     
Else
     
        sSelect = "select idmoneyreportoffice from moneyreportbyoffice where datereport=convert(datetime, '" + lbldate_agente.Caption + "') and idmanager='" + Str(ID_manager) + "' and idoffice='" + id_oficina$ + "'"
        Rs.Open sSelect, base, adOpenUnspecified
        id_moneyreportoffice$ = Rs(0)
        Rs.Close
     
        If id_moneyreportoffice$ = "" Then GoTo saltado
     
                        
         sSelect = "select reviewbyaccounting from moneyreportbyoffice where idmoneyreportoffice='" + id_moneyreportoffice$ + "'"
                
         Rs.Open sSelect, base, adOpenUnspecified
         revisado = Rs(0)
         Rs.Close
   
         If revisado = "False" Then
             chk_revisado.Value = 0
         Else
             chk_revisado.Value = 1
         End If
        
     
     
End If
     
     
     
     
saltado:
    
End Sub


Public Sub carga_todo()
On Error Resume Next


obtener_fecha_real
limpia_datos
img_tab1_Click (0)
carga_registros
carga_datos
calcula_total_LAE
btnupdate_LAE_Click
carga_totales_oficina
calcula_total_oficina_manager

If administrador = 1 Then
  'carga_agentes
End If

End Sub

Public Sub carga_manager()
On Error Resume Next

If valido1 = 777 Then
  Exit Sub
End If

Dim sSelect As String
    
    Dim Rs As ADODB.Recordset
    
 
    
    Set Rs = New ADODB.Recordset
           
  
   
' carga manager de la oficina


 Grid2.Clear
 
   If oficina_trabajada$ = "" Then
      oficina$ = LTrim(RTrim(Right(cbo_oficina.List(cbo_oficina.ListIndex), 30)))
   Else
      oficina$ = oficina_trabajada$
   End If
  
   If chkagentes(1).Value = 0 Then
  

  
  
  sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where emp.Active=1 and empofc.active=1 and IdJobTitle in (17, 29, 24) and ofc.office='" + oficina$ + "' and empjob.Active='1'"


    
  
  Else
  
 
  
  sSelect = "select emp.IdEmployee, Username, Office,  ciarel.IdJobTitle from EmployeeInfo emp " & _
  "join EmplDeptOfcRel empofc on empofc.IdEmployee= emp.IDEmployee " & _
  "join DeptOfcRel     depofc on depofc.IdDeptOfcRel = empofc.IdDeptOfcRel " & _
  "join OfficesCatalog ofc    on ofc.IdOffice = depofc.IdOffice " & _
  "join EmplJobTRel empjob on empjob.IDEmployee = emp.IDEmployee " & _
  "join CiaRegOfcDepJobTRel ciarel on ciarel.IdCiaRegOfcDepJobTRel= empjob.IdCiaRegOfcDepJobTRel " & _
  "where empofc.active=1 and IdJobTitle in (17, 29,24) and ofc.office='" + oficina$ + "' and empjob.Active='1'"
  
  
  
  
  End If
  
    
     ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    Grid2.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set Grid2.DataSource = Rs
                         
    Rs.Close
    
    
    
  cbo_managers.Clear
 
  
  
  For t = 1 To Grid2.Rows - 1
  Grid2.Row = t
  Grid2.Col = 1
  ID_manager$ = Grid2.Text
  
  Grid2.Col = 2
  manager$ = Grid2.Text
  
  
  If UCase(manager$) = "GJIMENEZ" Then
     GoTo saltado
  End If
  
  
  
  
  If oficina$ = "JA - HAVEN" Then
    Select Case UCase(manager$)
    Case "GJIMENEZ"
    
    Case "HNAVARRO", "DLOPEZ", "CCADENA", "MONEYREPORTS", "RECEIPTSCORRECTIONS", "BETZY", "BMARQUEZ"
    
    Case Else
       cbo_managers.AddItem manager$ + Space(20) + ID_manager$
    End Select
    
  Else
    
    existe = 0
    For Y = 0 To cbo_managers.ListCount - 1
       nombre_manager$ = Left(cbo_managers.List(Y), Len(cbo_managers.List(Y)) - 10)
       If RTrim(UCase(nombre_manager$)) = RTrim(UCase(manager$)) Then
           existe = 1
           Exit For
       End If
    Next Y
    
    If existe = 0 Then
         cbo_managers.AddItem manager$ + Space(20) + ID_manager$
    End If
    
    
  End If
  
saltado:
  
  
  Next t
   
   
   
valido1 = 777

If cbo_managers.ListCount = 1 Then
    cbo_managers.ListIndex = 0
    
ElseIf cbo_managers.ListCount = 0 Then
    cbo_managers.AddItem "JOSELIN" + Space(20) + "119"

End If
   
   
   
   
End Sub

Public Sub Carga_todas_las_oficinas()
On Error Resume Next

Dim sSelect As String
Dim Rs As ADODB.Recordset
    
Set Rs = New ADODB.Recordset
    
grid4.Clear

sSelect = "select idoffice, office, shortname from OfficesCatalog where active=1"
    
  ' Abrir el recordset de forma sin especificar, porque vamos a cambiar datos
    Rs.Open sSelect, base, adOpenUnspecified
    
    
     ' Permitir redimensionar las columnas
    grid4.AllowUserResizing = flexResizeColumns

    ' Asignar el recordset al FlexGrid
    Set grid4.DataSource = Rs
                         
    Rs.Close
    
    
    
    
    
    
    
End Sub

Public Sub actualiza_google()
On Error Resume Next


' actualiza Google
Call OReg.CrearNuevaClave(HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Google\Update")
  
Call OReg.EstablecerValor(HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Google\Update", "AutoUpdateCheckPeriodMinutesd", "1", REG_DWORD)
Call OReg.EstablecerValor(HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Google\Update", "UpdateDefault", "1", REG_DWORD)
  
' Exit Sub
  
  
n$ = "REG add " + Chr$(34) + "HKLM\SOFTWARE\Policies\Google\Update" + Chr$(34) + " /v AutoUpdateCheckPeriodMinutes /t REG_DWORD /d 1 /f /reg:64"
n2$ = "REG add " + Chr$(34) + "HKLM\SOFTWARE\Policies\Google\Update" + Chr$(34) + " /v UpdateDefault /t REG_DWORD /d 1 /f /reg:64"



 nf = FreeFile
Open "c:\iconos\google.bat" For Output Shared As #nf
Lock #nf
Print #nf, "@echo off"
Print #nf, "Cls"
Print #nf, "ECHO."
Print #nf, "ECHO ==================================="
Print #nf, "ECHO Actualizando Google"
Print #nf, "ECHO ==================================="
Print #nf, ""
Print #nf, ":init"
Print #nf, "setlocal DisableDelayedExpansion"
Print #nf, "Set cmdInvoke=1"
Print #nf, "Set winSysFolder=System32"
Print #nf, "set " + Chr$(34) + "batchPath=%~0" + Chr$(34)
Print #nf, "for %%k in (%0) do set batchName=%%~nk"
Print #nf, "set " + Chr$(34) + "vbsGetPrivileges=%temp%\OEgetPriv_%batchName%.vbs" + Chr$(34)
Print #nf, "setlocal EnableDelayedExpansion"
Print #nf, ""
Print #nf, ":checkPrivileges"
Print #nf, "NET FILE 1>NUL 2>NUL"
Print #nf, "if '%errorlevel%' == '0' ( goto gotPrivileges ) else ( goto getPrivileges )"
Print #nf, ""
Print #nf, ":getPrivileges"
Print #nf, "if '%1'=='ELEV' (echo ELEV & shift /1 & goto gotPrivileges)"
Print #nf, "ECHO."
Print #nf, "ECHO **************************************"
Print #nf, "ECHO Invocando UAC para realizar el cambio "
Print #nf, "ECHO **************************************"
Print #nf, ""
Print #nf, "ECHO Set UAC = CreateObject^(" + Chr$(34) + "Shell.Application" + Chr$(34) + "^) > " + Chr$(34) + "%vbsGetPrivileges%" + Chr$(34)
Print #nf, "ECHO args = " + Chr$(34) + "ELEV " + Chr$(34) + " >> " + Chr$(34) + "%vbsGetPrivileges%" + Chr$(34)
Print #nf, "ECHO For Each strArg in WScript.Arguments >> " + Chr$(34) + "%vbsGetPrivileges%" + Chr$(34)
Print #nf, "ECHO args = args ^& strArg ^& " + Chr$(34) + " " + Chr$(34) + "  >> " + Chr$(34) + "%vbsGetPrivileges%" + Chr$(34)
Print #nf, "ECHO Next >> " + Chr$(34) + "%vbsGetPrivileges%" + Chr$(34)
Print #nf, ""
Print #nf, "if '%cmdInvoke%'=='1' goto InvokeCmd"
Print #nf, ""
Print #nf, "ECHO UAC.ShellExecute " + Chr$(34) + "!batchPath!" + Chr$(34) + ", args, " + Chr$(34) + Chr$(34) + ", " + Chr$(34) + "runas" + Chr$(34) + ", 1 >> " + Chr$(34) + "%vbsGetPrivileges%" + Chr$(34)
Print #nf, "GoTo ExecElevation"
Print #nf, ""
Print #nf, ":InvokeCmd"
Print #nf, "ECHO args = " + Chr$(34) + "/c " + Chr$(34) + Chr$(34) + Chr$(34) + " + " + Chr$(34) + "!batchPath!" + Chr$(34) + " + " + Chr$(34) + Chr$(34) + Chr$(34) + " " + Chr$(34) + " + args >> " + Chr$(34) + "%vbsGetPrivileges%" + Chr$(34)
Print #nf, "ECHO UAC.ShellExecute " + Chr$(34) + "%SystemRoot%\%winSysFolder%\cmd.exe" + Chr$(34) + ", args, " + Chr$(34) + Chr$(34) + ", " + Chr$(34) + "runas" + Chr$(34) + ", 1 >> " + Chr$(34) + "%vbsGetPrivileges%" + Chr$(34)
Print #nf, ""
Print #nf, ":ExecElevation"
Print #nf, Chr$(34) + "%SystemRoot%\%winSysFolder%\WScript.exe" + Chr$(34) + " " + Chr$(34) + "%vbsGetPrivileges%" + Chr$(34) + " %*"
Print #nf, "exit /B"
Print #nf, ""
Print #nf, ":gotPrivileges"
Print #nf, "setlocal & cd /d %~dp0"
Print #nf, "if '%1'=='ELEV' (del " + Chr$(34) + "%vbsGetPrivileges%" + Chr$(34) + " 1>nul 2>nul  &  shift /1)"
Print #nf, ""
Print #nf, "::::::::::::::::::::::::::::"
Print #nf, "::start"
Print #nf, "::::::::::::::::::::::::::::"
Print #nf, "Rem Run shell as admin (example) - put here code as you like"
Print #nf, n$
Print #nf, n2$

Unlock #nf
Close nf



' r$ = Shell("cmd /c c:\transfer\google.bat", vbHide)
 






 



End Sub
