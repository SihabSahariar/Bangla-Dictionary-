VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5415
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "http://teamerror.org/"
      ForeColor       =   &H00808080&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Form2.frx":3503A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "KwcivBU 2015,wUg Bii"
      BeginProperty Font 
         Name            =   "Siyam Rupali ANSI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "wUg Bii"
      BeginProperty Font 
         Name            =   "Siyam Rupali ANSI"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   2160
      MouseIcon       =   "Form2.frx":3518C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "†cÖvMÖvwgs I wWRvBwbs"
      BeginProperty Font 
         Name            =   "Siyam Rupali ANSI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   5040
      MouseIcon       =   "Form2.frx":352DE
      MousePointer    =   99  'Custom
      Picture         =   "Form2.frx":35430
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "mdUIq¨vi m¤c‡K©"
      BeginProperty Font 
         Name            =   "Siyam Rupali ANSI"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Image5_Click()
Unload Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Label3_Click()
Shell "explorer.exe " & "http://teamerror.org/"
End Sub

Private Sub Label5_Click()
 Shell "explorer.exe " & "http://teamerror.org/"
End Sub
