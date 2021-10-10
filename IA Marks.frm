VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form7 
   Caption         =   "IA Marks Update"
   ClientHeight    =   10755
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17040
   Icon            =   "IA Marks.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "IA Marks.frx":2B3FB
   ScaleHeight     =   10755
   ScaleWidth      =   17040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton updatebtn2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9600
      Width           =   1455
   End
   Begin VB.CommandButton updatebtn1 
      BackColor       =   &H0000FFFF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9600
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1680
      Top             =   10560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ricchi\Desktop\CMS\CMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ricchi\Desktop\CMS\CMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Admission"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   7800
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   10560
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ricchi\Desktop\CMS\CMS.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Ricchi\Desktop\CMS\CMS.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Internal"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton getstudbtn 
      BackColor       =   &H00FFFF80&
      Caption         =   "Get Student Data"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox Text11 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   27
      Top             =   9000
      Width           =   3135
   End
   Begin VB.TextBox Text10 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   26
      Top             =   8400
      Width           =   3135
   End
   Begin VB.TextBox Text9 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   23
      Top             =   7800
      Width           =   3135
   End
   Begin VB.TextBox Text8 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   22
      Top             =   7200
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   19
      Top             =   6600
      Width           =   3135
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   18
      Top             =   6000
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   13
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   12
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   10
      Top             =   3600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      MaxLength       =   10
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image Image1 
      DataSource      =   "Adodc2"
      Height          =   2055
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "CCA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   29
      Top             =   9120
      Width           =   2535
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Sub"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   28
      Top             =   8520
      Width           =   2535
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Core-III Lab"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   7920
      Width           =   2535
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Core-III"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   7320
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Core-II Lab"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   21
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Core-II"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2280
      TabIndex        =   17
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2280
      TabIndex        =   16
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2280
      TabIndex        =   14
      Top             =   600
      Width           =   3135
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Language-II"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Language-I"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Core-I"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Core-I Lab"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Semester"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Combination"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reg No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Menu menuhom 
      Caption         =   "Home"
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pic As String

Private Sub Form_Load()
Text5.Enabled = False
Text7.Enabled = False
Text9.Enabled = False
updatebtn1.Visible = False
updatebtn2.Visible = False
End Sub

Private Sub getstudbtn_Click()
Adodc2.RecordSource = "select * from Admission where Reg_No = '" + Text1.Text + "'"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount = 0 Then
MsgBox "No Data Found"
Text1.Text = ""
ElseIf Adodc2.Recordset.RecordCount = 1 Then
Dim IAEdit As String
IAEdit = MsgBox("Data Already Updated Do you want to Edit?", vbYesNoCancel + vbInformation)
If IAEdit = vbYes Then
updatebtn2.Visible = True
Adodc1.RecordSource = "select * from Internal where Reg_No = '" + Text1.Text + "'"
Adodc1.Refresh
pic = Adodc2.Recordset.Fields("Photo").Value
Image1.Picture = LoadPicture(pic)
Label12 = Adodc1.Recordset("Name")
Label13 = Adodc1.Recordset("Gender")
Label14 = Adodc1.Recordset("Combination")
Label15 = Adodc1.Recordset("Semester")
If Label14.Caption = "PCM" Or Label14.Caption = "PMCs" Or Label14.Caption = "CBZ" Or Label14.Caption = "Bio-Technology" Or Label14.Caption = "BCA" Then
Text5.Enabled = True
Text7.Enabled = True
Text9.Enabled = True
End If
pic = Adodc2.Recordset.Fields("Photo").Value
Image1.Picture = LoadPicture(pic)
Text2 = Adodc1.Recordset("Language_I")
Text3 = Adodc1.Recordset("Language_II")
Text4 = Adodc1.Recordset("Core_I")
Text5 = Adodc1.Recordset("Core_I_Lab")
Text6 = Adodc1.Recordset("Core_II")
Text7 = Adodc1.Recordset("Core_II_Lab")
Text8 = Adodc1.Recordset("Core_III")
Text9 = Adodc1.Recordset("Core_III_Lab")
Text10 = Adodc1.Recordset("AdditionalSub")
Text11 = Adodc1.Recordset("CCA")
ElseIf IAEdit = vbNo Then
Text1.Text = ""
End If
Else
Label12 = Adodc2.Recordset("Name")
Label13 = Adodc2.Recordset("Gender")
Label14 = Adodc2.Recordset("Combination")
Label15 = Adodc2.Recordset("Semester")
If Label14.Caption = "PCM" Or Label14.Caption = "PMCs" Or Label14.Caption = "CBZ" Or Label14.Caption = "Bio-Technology" Or Label14.Caption = "BCA" Then
Text5.Enabled = True
Text7.Enabled = True
Text9.Enabled = True
End If
updatebtn1.Visible = True
End If
End Sub

Private Sub menuhom_Click()
Form3.Show
Unload Me
End Sub

Private Sub updatebtn1_Click()
If Text1.Text = "" Or Label12.Caption = "" Or Label13.Caption = "" Or Label14.Caption = "" Or Label15.Caption = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text6.Text = "" Or Text8.Text = "" Or Text10.Text = "" Or Text11.Text = "" Then
MsgBox "Enter Valid Marks"
Else
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields("Reg_No").Value = Text1.Text
Adodc1.Recordset.Fields("Name").Value = Label12.Caption
Adodc1.Recordset.Fields("Gender").Value = Label13.Caption
Adodc1.Recordset.Fields("Combination").Value = Label14.Caption
Adodc1.Recordset.Fields("Semester").Value = Label15.Caption
Adodc1.Recordset.Fields("Language_I").Value = Text2.Text
Adodc1.Recordset.Fields("Language_II").Value = Text3.Text
Adodc1.Recordset.Fields("Core_I").Value = Text4.Text
Adodc1.Recordset.Fields("Core_I_Lab").Value = Text5.Text
Adodc1.Recordset.Fields("Core_II").Value = Text6.Text
Adodc1.Recordset.Fields("Core_II_Lab").Value = Text7.Text
Adodc1.Recordset.Fields("Core_III").Value = Text8.Text
Adodc1.Recordset.Fields("Core_III_Lab").Value = Text9.Text
Adodc1.Recordset.Fields("AdditionalSub").Value = Text10.Text
Adodc1.Recordset.Fields("CCA").Value = Text11.Text
Adodc1.Recordset.Update
MsgBox "Updated"
Text1.Text = ""
Label12.Caption = ""
Label13.Caption = ""
Label14.Caption = ""
Label15.Caption = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
End If
End Sub

Private Sub updatebtn2_Click()
If Text1.Text = "" Or Label12.Caption = "" Or Label13.Caption = "" Or Label14.Caption = "" Or Label15.Caption = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text6.Text = "" Or Text8.Text = "" Or Text10.Text = "" Or Text11.Text = "" Then
MsgBox "Enter Valid Marks"
Else
Adodc1.Recordset.Fields("Reg_No").Value = Text1.Text
Adodc1.Recordset.Fields("Name").Value = Label12.Caption
Adodc1.Recordset.Fields("Gender").Value = Label13.Caption
Adodc1.Recordset.Fields("Combination").Value = Label14.Caption
Adodc1.Recordset.Fields("Semester").Value = Label15.Caption
Adodc1.Recordset.Fields("Language_I").Value = Text2.Text
Adodc1.Recordset.Fields("Language_II").Value = Text3.Text
Adodc1.Recordset.Fields("Core_I").Value = Text4.Text
Adodc1.Recordset.Fields("Core_I_Lab").Value = Text5.Text
Adodc1.Recordset.Fields("Core_II").Value = Text6.Text
Adodc1.Recordset.Fields("Core_II_Lab").Value = Text7.Text
Adodc1.Recordset.Fields("Core_III").Value = Text8.Text
Adodc1.Recordset.Fields("Core_III_Lab").Value = Text9.Text
Adodc1.Recordset.Fields("AdditionalSub").Value = Text10.Text
Adodc1.Recordset.Fields("CCA").Value = Text11.Text
Adodc1.Recordset.Update
MsgBox "Updated"
Text1.Text = ""
Label12.Caption = ""
Label13.Caption = ""
Label14.Caption = ""
Label15.Caption = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
End If
End Sub
