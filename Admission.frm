VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form5 
   Caption         =   "Admission"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17040
   Icon            =   "Admission.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Admission.frx":2B3FB
   ScaleHeight     =   10635
   ScaleWidth      =   17040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo6 
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
      Height          =   405
      ItemData        =   "Admission.frx":37D11
      Left            =   2640
      List            =   "Admission.frx":37D1E
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   2280
      Width           =   3375
   End
   Begin VB.ComboBox Combo5 
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
      Height          =   405
      ItemData        =   "Admission.frx":37D3D
      Left            =   2640
      List            =   "Admission.frx":37D4D
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   3000
      Width           =   3375
   End
   Begin VB.CommandButton ADMITBTN 
      BackColor       =   &H0000FFFF&
      Caption         =   "Admitt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   9480
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   8760
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo4 
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
      Height          =   405
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   5880
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2640
      TabIndex        =   23
      Top             =   1560
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   113901569
      CurrentDate     =   43561
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   10200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
   Begin VB.CommandButton newbtn 
      BackColor       =   &H0000FF00&
      Caption         =   "NEW ADMIT"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Checkbtn 
      BackColor       =   &H0000FF00&
      Caption         =   "CHECK"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton uploadbtn 
      BackColor       =   &H000040C0&
      Caption         =   "Upload Photo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
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
      Height          =   405
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   5160
      Width           =   3375
   End
   Begin VB.ComboBox Combo2 
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
      Height          =   405
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   4440
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
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
      Height          =   405
      Left            =   2640
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   2640
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   9480
      Width           =   3375
   End
   Begin VB.TextBox Text6 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2640
      TabIndex        =   8
      Top             =   8760
      Width           =   3375
   End
   Begin VB.TextBox Text5 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2640
      TabIndex        =   7
      Top             =   8040
      Width           =   3375
   End
   Begin VB.TextBox Text4 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2640
      TabIndex        =   6
      Top             =   7320
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      TabIndex        =   4
      Top             =   6600
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   31
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   30
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Image Image4 
      Height          =   2295
      Left            =   9360
      Picture         =   "Admission.frx":37D67
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   4695
   End
   Begin VB.Image Image3 
      Height          =   2775
      Left            =   16920
      Picture         =   "Admission.frx":389BF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3615
   End
   Begin VB.Image Image2 
      Height          =   4575
      Left            =   15360
      Picture         =   "Admission.frx":3F173
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   5055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Combination"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   26
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Image Image1 
      DataSource      =   "Adodc1"
      Height          =   2055
      Left            =   6360
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Prev. Year %"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Cantact No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   8760
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   9600
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Prev. Comb."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   13
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Semester"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Mother Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   8040
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Admission To"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Register No."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
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
   Begin VB.Menu admit1 
      Caption         =   "New Admission"
   End
   Begin VB.Menu admit2 
      Caption         =   "Modify Admission"
   End
   Begin VB.Menu homemenu 
      Caption         =   "Home"
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim reg As Integer
Dim pic As String

Private Sub admit1_Click()
Image1.Visible = True
Checkbtn.Visible = False
newbtn.Visible = True
Text1.Enabled = False
End Sub

Private Sub admit2_Click()
Image1.Visible = True
Checkbtn.Visible = True
newbtn.Visible = False
Text1.Enabled = True
uploadbtn.Visible = False
End Sub

Private Sub ADMITBTN_Click()
If Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Combo4.Text = "" Or Combo5.Text = "" Or Combo6.Text = "" Or Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Then
MsgBox "Enter Correct Details", , "CMS"
Else
Adodc1.Recordset.Fields("Reg_No").Value = Text1.Text
Adodc1.Recordset.Fields("Name").Value = Text2.Text
Adodc1.Recordset.Fields("DOB").Value = DTPicker1.Value
Adodc1.Recordset.Fields("Course").Value = Combo1.Text
Adodc1.Recordset.Fields("Combination").Value = Combo2.Text
Adodc1.Recordset.Fields("Semester").Value = Combo3.Text
Adodc1.Recordset.Fields("Previous_Stream").Value = Combo4.Text
Adodc1.Recordset.Fields("Gender").Value = Combo5.Text
Adodc1.Recordset.Fields("Category").Value = Combo6.Text
Adodc1.Recordset.Fields("Previous_Perc").Value = Text3.Text
Adodc1.Recordset.Fields("Fathername").Value = Text4.Text
Adodc1.Recordset.Fields("Mothername").Value = Text5.Text
Adodc1.Recordset.Fields("Cantact_No").Value = Text6.Text
Adodc1.Recordset.Fields("Address").Value = Text7.Text
Adodc1.Recordset.Fields("Photo").Value = pic
Adodc1.Recordset.Fields("AdmitDate").Value = Format(Now, "dd/mm/yyyy")
Adodc1.Recordset.Update
MsgBox "successfully admitted", , "CMS"
Text2.Text = ""
DTPicker1 = Format(Now, "dd/mm/yyyy")
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End If
End Sub

Private Sub Checkbtn_Click()
Adodc1.RecordSource = "select * from Admission where Reg_No = '" + Text1.Text + "'"
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
MsgBox "No Data Found"
Text1.Text = ""
Else
Text2 = Adodc1.Recordset("Name")
DTPicker1 = Adodc1.Recordset("DOB")
Combo1 = Adodc1.Recordset("Course")
Combo2 = Adodc1.Recordset("Combination")
Combo3 = Adodc1.Recordset("Semester")
Combo4 = Adodc1.Recordset("Previous_Stream")
Combo5 = Adodc1.Recordset("Gender")
Combo6 = Adodc1.Recordset("Category")
Text3 = Adodc1.Recordset("Previous_Perc")
Text4 = Adodc1.Recordset("Fathername")
Text5 = Adodc1.Recordset("Mothername")
Text6 = Adodc1.Recordset("Cantact_No")
Text7 = Adodc1.Recordset("Address")
pic = Adodc1.Recordset("Photo")
Image1.Picture = LoadPicture(pic)
ADMITBTN.Visible = True
End If
End Sub

Private Sub Combo1_Click()
Combo2.Clear
Combo4.Clear
If Combo1.Text = "BA" Then
Combo2.AddItem "HEP"
Combo2.AddItem "HES"
Combo2.AddItem "HEJ"
Combo2.AddItem "HJE"
Combo4.AddItem "HEPS"
Combo4.AddItem "HEGP"
Combo4.AddItem "HEPyS"
Combo4.AddItem "HELP"
Combo4.AddItem "HEPK"
Combo4.AddItem "HESK"
Combo4.AddItem "HELS"
Combo4.AddItem "HEKM"
Combo4.AddItem "HEKK"
Combo4.AddItem "ELSP"
Combo4.AddItem "HEKH"
ElseIf Combo1.Text = "BSC" Then
Combo2.AddItem "PCM"
Combo2.AddItem "PMCs"
Combo2.AddItem "CBZ"
Combo2.AddItem "Bio-Technology"
Combo4.AddItem "PCMB"
Combo4.AddItem "PCMC"
Combo4.AddItem "PCME"
Combo4.AddItem "PCMG"
Combo4.AddItem "PCBH"
Combo4.AddItem "PCBE"
ElseIf Combo1.Text = "BCOM" Then
Combo2.AddItem "BCOM"
Combo4.AddItem "SEBA"
Combo4.AddItem "ABEM"
Combo4.AddItem "BSAM"
Combo4.AddItem "EABC"
Combo4.AddItem "HEBA"
Combo4.AddItem "GEBA"
Combo4.AddItem "BACS"
Combo4.AddItem "ABPE"
ElseIf Combo1.Text = "BBA" Then
Combo2.AddItem "BBA"
Combo4.AddItem "SEBA"
Combo4.AddItem "ABEM"
Combo4.AddItem "BSAM"
Combo4.AddItem "EABC"
Combo4.AddItem "HEBA"
Combo4.AddItem "GEBA"
Combo4.AddItem "BACS"
Combo4.AddItem "ABPE"
ElseIf Combo1.Text = "BCA" Then
Combo2.AddItem "BCA"
Combo4.AddItem "PCMB"
Combo4.AddItem "PCMC"
Combo4.AddItem "PCME"
Combo4.AddItem "PCMG"
Else
End If
End Sub


Private Sub Form_Load()
Combo1.AddItem "BA"
Combo1.AddItem "BSC"
Combo1.AddItem "BCOM"
Combo1.AddItem "BBA"
Combo1.AddItem "BCA"
Combo3.AddItem "Semester-I"
Combo3.AddItem "Semester-II"
Combo3.AddItem "Semester-III"
Combo3.AddItem "Semester-IV"
Combo3.AddItem "Semester-V"
Combo3.AddItem "Semester-VI"
Checkbtn.Visible = False
newbtn.Visible = False
uploadbtn.Visible = False
ADMITBTN.Visible = False
End Sub

Private Sub homemenu_Click()
Form3.Show
Form5.Hide
End Sub

Private Sub newbtn_Click()
Call AutoReg
Adodc1.Recordset.AddNew
Text1 = Format(reg, "17P5V85000")
uploadbtn.Visible = True
Text2.Text = ""
DTPicker1 = Format(Now, "dd/mm/yyyy")
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

Public Sub AutoReg()
On Error GoTo Err_id
Adodc1.Refresh
If Adodc1.Recordset.RecordCount = 0 Then
reg = 1
Else
Adodc1.Recordset.MoveLast
reg = Mid(Adodc1.Recordset("Reg_No"), 9, 10) + 1
Adodc1.Refresh
End If
Exit Sub
Err_id:
reg = 1
MsgBox "Register is not generated", vbCritical, "CMS"
End Sub

Private Sub uploadbtn_Click()
Cd1.ShowOpen
Cd1.Filter = "Jpeg|*.jpg"
pic = Cd1.FileName
Image1.Picture = LoadPicture(pic)
ADMITBTN.Visible = True
End Sub
