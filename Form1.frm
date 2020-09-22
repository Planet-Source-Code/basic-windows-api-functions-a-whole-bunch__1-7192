VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000000FF&
   Caption         =   "nKm Basic API Functions - Enjoy!"
   ClientHeight    =   6615
   ClientLeft      =   2370
   ClientTop       =   2625
   ClientWidth     =   3120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   3120
   Begin VB.Frame Frame3 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3855
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2895
      Begin VB.CommandButton Command17 
         BackColor       =   &H000000C0&
         Caption         =   "-nKm iNc- -Website-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H000000C0&
         Caption         =   "MouseUp Command(R)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H000000C0&
         Caption         =   "MouseDown Command(R)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H000000C0&
         Caption         =   "Simulate Click(R)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H000000C0&
         Caption         =   "MouseUp Command(L)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H000000C0&
         Caption         =   "MouseDown Command(L)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H000000C0&
         Caption         =   "Simulate Click(L)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H000000C0&
         Caption         =   "Close Window"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H000000C0&
         Caption         =   "Change Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000C0&
         Caption         =   " Hide  Window"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H000000C0&
         Caption         =   "Show Window"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H000000C0&
         Caption         =   "Minimize Window"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H000000C0&
         Caption         =   "Maximize Window"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H000000C0&
         Caption         =   "Restore Window"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H000000C0&
         Caption         =   "Set on Top of All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H000000C0&
         Caption         =   "Set Not on Top of All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000C0&
      Caption         =   "Get hWnd"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000000FF&
      Caption         =   "Windows:"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.ListBox List2 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
      Begin VB.ListBox List1 
         BackColor       =   &H000000C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By nKm"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   6120
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":0000
      Height          =   1455
      Left            =   3120
      TabIndex        =   22
      Top             =   1560
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ŽŽ    ŽŽ  ŽŽ   ŽŽ  ŽŽ     ŽŽ
'   ŽŽŽ   ŽŽ  ŽŽ  ŽŽ   ŽŽŽ   ŽŽŽ
'   ŽŽŽŽ  ŽŽ  ŽŽ ŽŽ    ŽŽŽŽ ŽŽŽŽ
'   ŽŽ ŽŽ ŽŽ  ŽŽŽŽŽ    ŽŽ ŽŽŽ ŽŽ
'   ŽŽ  ŽŽŽŽ  ŽŽ  ŽŽ   ŽŽ  Ž  ŽŽ
'   ŽŽ   ŽŽŽ  ŽŽ   ŽŽ  ŽŽ     ŽŽ
'   ŽŽ    ŽŽ  ŽŽ   ŽŽ  ŽŽ     ŽŽ
'  ====-====-====-====-====-=====
'      B    A    S    I    C
'  ====-====-====-====-====-=====
'   ŽŽ           ŽŽŽŽŽŽ        ŽŽ
'  ŽŽŽŽ         ŽŽ   ŽŽ        ŽŽ
' ŽŽ  ŽŽ        ŽŽ   ŽŽ        ŽŽ
' ŽŽ  ŽŽ        ŽŽ   ŽŽ        ŽŽ
' ŽŽŽŽŽŽ        ŽŽŽŽŽŽ         ŽŽ
' ŽŽ  ŽŽ        ŽŽ             ŽŽ
' ŽŽ  ŽŽ        ŽŽ             ŽŽ

'This file was made my nKm, of nKm iNc.
'Made for Planet Cource Code.
'API by me.
'All the CONSTs.  Took from OxidApi.bas,
'-HELL NO I WOULDN"T TYPE ALL THOUGHS CONSTS OUT!

'AIM: LxL nkm LxL       (i change aim sn alot)
'AOL: xxnkmxx           (will never change sn)
'EMAIL: xxnkmxx@aol.com (will never chande address)



'--BEGIN STUFF TO DECLARE!!--'
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function WindowFromPointXY Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SendMessageLong& Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Private Declare Function GetCursorPos Lib "user32" (lpPoint As Moo) As Long

'--END STUFF TO DECLARE!!--'




'Now is the nice part.  If you don't like
'the options in this program.
'You can use one (or more) of the CONSTs
'that are down below.
'At the begining of each section I will specify
'how to use them.  Ok?  Here.... there are alot...







'--BEGIN CONSTs FOR SendMessageLong--'
'HOW TO:
'eXample--
'cow = SendMessageLong(Hoover, WM_NCRBUTTONDOWN, 0&, 0&)
'That will make the hWnd that the mouse is over think that
'the right button as just been pressed.  Not released. Juss pressed.
'----------=============------------
'Here are a giant list of consts....
Private Const WM_NCCREATE = &H81
Private Const WM_NCDESTROY = &H82
Private Const WM_NCCALCSIZE = &H83
Private Const WM_NCHITTEST = &H84
Private Const WM_NCPAINT = &H85
Private Const WM_NCACTIVATE = &H86
Private Const WM_GETDLGCODE = &H87
Private Const WM_NCMOUSEMOVE = &HA0
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCLBUTTONUP = &HA2
Private Const WM_NCLBUTTONDBLCLK = &HA3
Private Const WM_NCRBUTTONDOWN = &HA4
Private Const WM_NCRBUTTONUP = &HA5
Private Const WM_NCRBUTTONDBLCLK = &HA6
Private Const WM_NCMBUTTONDOWN = &HA7
Private Const WM_NCMBUTTONUP = &HA8
Private Const WM_NCMBUTTONDBLCLK = &HA9
Private Const WM_KEYFIRST = &H100
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const WM_CHAR = &H102
Private Const WM_DEADCHAR = &H103
Private Const WM_SYSKEYDOWN = &H104
Private Const WM_SYSKEYUP = &H105
Private Const WM_SYSCHAR = &H106
Private Const WM_SYSDEADCHAR = &H107
Private Const WM_KEYLAST = &H108
Private Const WM_INITDIALOG = &H110
Private Const WM_COMMAND = &H111
Private Const WM_SYSCOMMAND = &H112
Private Const WM_TIMER = &H113
Private Const WM_HSCROLL = &H114
Private Const WM_VSCROLL = &H115
Private Const WM_INITMENU = &H116
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_MENUSELECT = &H11F
Private Const WM_MENUCHAR = &H120
Private Const WM_ENTERIDLE = &H121
Private Const WM_CTLCOLORMSGBOX = &H132
Private Const WM_CTLCOLOREDIT = &H133
Private Const WM_CTLCOLORLISTBOX = &H134
Private Const WM_CTLCOLORBTN = &H135
Private Const WM_CTLCOLORDLG = &H136
Private Const WM_CTLCOLORSCROLLBAR = &H137
Private Const WM_CTLCOLORSTATIC = &H138
Private Const WM_MOUSEFIRST = &H200
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MOUSELAST = &H209
Private Const WM_PARENTNOTIFY = &H210
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_EXITMENULOOP = &H212
Private Const WM_MDICREATE = &H220
Private Const WM_MDIDESTROY = &H221
Private Const WM_MDIACTIVATE = &H222
Private Const WM_MDIRESTORE = &H223
Private Const WM_MDINEXT = &H224
Private Const WM_MDIMAXIMIZE = &H225
Private Const WM_MDITILE = &H226
Private Const WM_MDICASCADE = &H227
Private Const WM_MDIICONARRANGE = &H228
Private Const WM_MDIGETACTIVE = &H229
Private Const WM_MDISETMENU = &H230
Private Const WM_DROPFILES = &H233
Private Const WM_MDIREFRESHMENU = &H234
Private Const WM_CUT = &H300
Private Const WM_COPY = &H301
Private Const WM_PASTE = &H302
Private Const WM_CLEAR = &H303
Private Const WM_UNDO = &H304
Private Const WM_RENDERFORMAT = &H305
Private Const WM_RENDERALLFORMATS = &H306
Private Const WM_DESTROYCLIPBOARD = &H307
Private Const WM_DRAWCLIPBOARD = &H308
Private Const WM_PAINTCLIPBOARD = &H309
Private Const WM_VSCROLLCLIPBOARD = &H30A
Private Const WM_SIZECLIPBOARD = &H30B
Private Const WM_ASKCBFORMATNAME = &H30C
Private Const WM_CHANGECBCHAIN = &H30D
Private Const WM_HSCROLLCLIPBOARD = &H30E
Private Const WM_QUERYNEWPALETTE = &H30F
Private Const WM_PALETTEISCHANGING = &H310
Private Const WM_PALETTECHANGED = &H311
Private Const WM_HOTKEY = &H312
Private Const WM_PENWINFIRST = &H380
Private Const WM_PENWINLAST = &H38F
'--END CONSTs FOR SendMessageLong--'





'--BEGIN CONSTs FOR ShowWindow--'
'HOW TO
'eXample:
'cow = ShowWindow(Hoover, SWHIDE)
'That will hide the window that the mouse is over.
'Very mean >=)
'----------=============------------
'Here are a giant list of consts....
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const HWND_TOP = 0
Private Const HWND_BOTTOM = 1
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
'--END CONSTs FOR ShowWindow--'



Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Dim ThehWnd As Long
Private Type Moo
    X As Long
    Y As Long
End Type
Function Capshin(hWnd)
hwndLength% = GetWindowTextLength(hWnd)
hwndTitle$ = String$(hwndLength%, 0)
a% = GetWindowText(hWnd, hwndTitle$, (hwndLength% + 1))

Capshin = hwndTitle$
End Function
Sub Convert()
ThehWnd = List2.text
End Sub
Function Hoover()
      Dim nKm1 As Moo
      Dim nKmX As Long
      Dim nKmY As Long
   
      Call GetCursorPos(nKm1)
      nKmX = nKm1.X
      nKmY = nKm1.Y
      Hoover = WindowFromPointXY(nKmX, nKmY)
End Function






Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.MousePointer = 10
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim IamAcow As String
Dim yo As Long
yo = Hoover
Command1.MousePointer = 0
IamAcow = InputBox("Caption to add for " & yo, "Caption to add...", Capshin(Hoover), 100, 100)
If IamAcow = "" Then GoTo CoWzOwNmE
List2.AddItem yo
List1.AddItem IamAcow
Exit Sub

CoWzOwNmE:
MsgBox "You need to type in a caption!", vbCritical, "Viva La Cows!"
End Sub

Private Sub Command10_Click()
Call Convert
cow = SendMessageLong(ThehWnd, &H10, 0&, 0&)
End Sub

Private Sub Command11_Click()
Call Convert
cow = SendMessageLong(ThehWnd, &H201, 0&, 0&)
cow = SendMessageLong(ThehWnd, &H202, 0&, 0&)
End Sub

Private Sub Command12_Click()
Call Convert
cow = SendMessageLong(ThehWnd, &H201, 0&, 0&)
End Sub

Private Sub Command13_Click()
Call Convert
cow = SendMessageLong(ThehWnd, &H202, 0&, 0&)
End Sub

Private Sub Command14_Click()
Call Convert
cow = SendMessageLong(ThehWnd, &H204, 0&, 0&)
cow = SendMessageLong(ThehWnd, &H205, 0&, 0&)

End Sub

Private Sub Command15_Click()
Call Convert
cow = SendMessageLong(ThehWnd, &H204, 0&, 0&)
End Sub

Private Sub Command16_Click()
Call Convert
cow = SendMessageLong(ThehWnd, &H205, 0&, 0&)
End Sub

Private Sub Command17_Click()
Call Shell("explorer http://www.nkm.cjb.net/", vbMaximizedFocus)
End Sub

Private Sub Command2_Click()
Call Convert
cow = ShowWindow(ThehWnd, 0)
End Sub

Private Sub Command3_Click()
Call Convert
cow = ShowWindow(ThehWnd, 5)
End Sub

Private Sub Command4_Click()
Call Convert
cow = ShowWindow(ThehWnd, 6)
End Sub

Private Sub Command5_Click()
cow = ShowWindow(ThehWnd, 3)
End Sub


Private Sub Command6_Click()
Call Convert
cow = SetWindowPos(ThehWnd, -2, 0, 0, 0, 0, FLAGS)
End Sub

Private Sub Command7_Click()
Call Convert
cow = SetWindowPos(ThehWnd, -2, 0, 0, 0, 0, FLAGS)
cow = ShowWindow(ThehWnd, 1)
cow = ShowWindow(ThehWndt, 5)

End Sub

Private Sub Command8_Click()
Call Convert
cow = SetWindowPos(ThehWnd, -1, 0, 0, 0, 0, FLAGS)
End Sub

Private Sub Command9_Click()
Call Convert
Dim iLIKEcows
iLIKEcows = InputBox("What do you want the new caption to be?", "New Caption Name", "Cows rock!", 100, 100)
cow% = SendMessageByString(ThehWnd, &HC, 0&, iLIKEcows)
End Sub

Private Sub Form_Resize()
If Me.ScaleWidth < 3015 Then Exit Sub
Frame1.Width = Me.ScaleWidth
List1.Width = Frame1.Left + Frame1.Width - List1.Left - 120
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "By nKm " & Chr$(13) & Chr$(13) & "xxnkmxx@aol.com" & Chr$(13) & Chr$(13) & "My aim sn is: LxL nkm LxL" & Chr$(13) & "I CHANGE IT OFTEN.  IF YOU WANT MY SN.  EMAIL ME", vbOKOnly + vbInformation, "nKm Basic API Functions - Goodbye"
End Sub




Private Sub List1_Click()
If List2.text <> "" Then Frame3.Enabled = True
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If List2.text <> "" Then Frame3.Enabled = True
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If List2.text <> "" Then Frame3.Enabled = True
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
If List2.text <> "" Then Frame3.Enabled = True
List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
List1.ListIndex = List2.ListIndex
If List2.text <> "" Then Frame3.Enabled = True

End Sub





Private Sub List2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If List2.text <> "" Then Frame3.Enabled = True
List2.ListIndex = List1.ListIndex
End Sub
