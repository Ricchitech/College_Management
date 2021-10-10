VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "College Management"
   ClientHeight    =   10635
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17040
   Icon            =   "Home.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Home.frx":2B3FB
   ScaleHeight     =   10635
   ScaleWidth      =   17040
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Image Image1 
      Height          =   3255
      Left            =   12960
      Picture         =   "Home.frx":76C10
      Stretch         =   -1  'True
      Top             =   720
      Width           =   3735
   End
   Begin VB.Menu mnuusr 
      Caption         =   "User"
   End
   Begin VB.Menu admitmenu 
      Caption         =   "Admission"
   End
   Begin VB.Menu feemenu 
      Caption         =   "Fee Payment"
   End
   Begin VB.Menu internalsmenu 
      Caption         =   "Internals"
   End
   Begin VB.Menu reports 
      Caption         =   "Reports"
      Begin VB.Menu report1 
         Caption         =   "Admission"
      End
      Begin VB.Menu report2 
         Caption         =   "Fee"
      End
      Begin VB.Menu report3 
         Caption         =   "Internals"
      End
   End
   Begin VB.Menu mnulog 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub admitmenu_Click()
Form5.Show
Form3.Hide
End Sub

Private Sub feemenu_Click()
Form6.Show
Form3.Hide
End Sub

Private Sub internalsmenu_Click()
Form7.Show
Form3.Hide
End Sub

Private Sub mnulog_Click()
Form2.Show
Unload Me
End Sub

Private Sub mnuusr_Click()
Form4.Show
Form3.Hide
End Sub

Private Sub report1_Click()
Form8.Show
Form3.Hide
End Sub

Private Sub report2_Click()
Form9.Show
Form3.Hide
End Sub

Private Sub report3_Click()
Form10.Show
Form3.Hide
End Sub
