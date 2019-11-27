VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "VText.dll"
Begin VB.Form Form1 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   -45
   ClientWidth     =   7650
   Icon            =   "t.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "t.frx":6852
   ScaleHeight     =   8490
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HTTSLibCtl.TextToSpeech v 
      Height          =   375
      Left            =   8280
      OleObjectBlob   =   "t.frx":DA3BC
      TabIndex        =   8
      Top             =   3120
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   840
      TabIndex        =   7
      Top             =   4320
      Width           =   3975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Windows\siz.1"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Siz"
      Top             =   7560
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "A_©"
      BeginProperty Font 
         Name            =   "Siyam Rupali ANSI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   840
      TabIndex        =   0
      Top             =   5040
      Width           =   3975
      Begin VB.TextBox Text3 
         DataField       =   "ag"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1080
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1410
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         DataField       =   "add"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         DataField       =   "nam"
         DataSource      =   "Data1"
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
         Left            =   720
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "D”PviY:"
         BeginProperty Font 
            Name            =   "Siyam Rupali ANSI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "A_© :"
         BeginProperty Font 
            Name            =   "Siyam Rupali ANSI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "kã:"
         BeginProperty Font 
            Name            =   "Siyam Rupali ANSI"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   615
      Left            =   8400
      TabIndex        =   9
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Image Image8 
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image7 
      Height          =   615
      Left            =   0
      MouseIcon       =   "t.frx":DA3E0
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   6735
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   6840
      MouseIcon       =   "t.frx":DA532
      MousePointer    =   99  'Custom
      ToolTipText     =   "About"
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   7200
      MouseIcon       =   "t.frx":DA684
      MousePointer    =   99  'Custom
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   6120
      MouseIcon       =   "t.frx":DA7D6
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   0
      MouseIcon       =   "t.frx":DA928
      MousePointer    =   99  'Custom
      Top             =   8160
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   5040
      MouseIcon       =   "t.frx":DAA7A
      MousePointer    =   99  'Custom
      ToolTipText     =   "Listen A Word"
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   5040
      MouseIcon       =   "t.frx":DABCC
      MousePointer    =   99  'Custom
      ToolTipText     =   "Search A Word"
      Top             =   4200
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1

Private Sub Form_Load()
Label1.FontName = "Siyam Rupali ANSI"
Label1.FontSize = 12
Label2.FontName = "Siyam Rupali ANSI"
Label2.FontSize = 12
Label3.FontName = "Siyam Rupali ANSI"
Label3.FontSize = 12
Text3.FontName = "Siyam Rupali ANSI"
Text3.FontSize = 12
Text2.FontName = "Siyam Rupali ANSI"
Text2.FontSize = 12
If Dir(App.Path & "/Resource.tlb") = "" Then 'Database Using Microsoft Access 2003
MsgBox "Sorry Database Not Found.", 0 + vbExclamation, "Error"
End
Else
Data1.DatabaseName = App.Path & "/Resource.tlb"
Data1.RecordSource = "siz"
End If
End Sub
Private Sub Image1_Click()
 Dim content
    content = Trim(Text4.Text) & "*"
    content = "nam like '" & content & "'"
    If Text4.Text <> "" Then
        Data1.Recordset.FindFirst content
    End If
    
   
    
'If You Wanna Add Auto Speaking Function'
    
    'If Text4.Text = "" Then
'MsgBox "Please Enter A Word", vbCritical, "Error"
'Else
'v.Speak (Text4.Text)
'End If
End Sub

Private Sub Image2_Click()
If Text1.Text = "" Then
MsgBox "Please Enter A Word", vbCritical, "Error"
Else
v.Speak (Text1.Text)
End If

End Sub

Private Sub Image3_Click()
 Shell "explorer.exe " & "http:/facebook.com/bd.terror/"
End Sub

Private Sub Image4_Click()
MsgBox "You Can Modify & Develop It But Never Forget To Give Credit", vbInformation, "Warning"
End Sub

Private Sub Image5_Click()
End
End Sub

Private Sub Image6_Click()
'MsgBox "A Freeware & Opensource Prject Of Team Error", vbInformation, "About"
Form2.Show
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

