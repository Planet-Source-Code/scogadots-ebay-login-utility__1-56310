VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "eBay Login Utility"
   ClientHeight    =   2685
   ClientLeft      =   150
   ClientTop       =   735
   ClientWidth     =   4065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000003&
      Caption         =   "Remember eBay ID and Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   3015
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   30
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   30
      ExtentX         =   53
      ExtentY         =   53
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&Submit"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   3795
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000003&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000003&
      Caption         =   "eBay ID:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   -360
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Menu mnuPop 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu reset 
         Caption         =   "Reset Fields"
      End
   End
   Begin VB.Menu links 
      Caption         =   "Links"
      Begin VB.Menu myebay 
         Caption         =   "My eBay"
      End
      Begin VB.Menu home 
         Caption         =   "eBay Home"
      End
      Begin VB.Menu time 
         Caption         =   "eBay Offical Time"
      End
      Begin VB.Menu paypal 
         Caption         =   "PayPal"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This code was written by Michael DiVincent
'e-mail address: rnm89@divincent.com
'
'Purpose: It is provided for instructional purposes.
'
'Details: Allows an eBay user to log into there account quickly
'without any typing and saves the username and password into the registry.
'
'Disclaimer: This code is provided as-is, without warranty. The author
'takes no responsibility for problems arising from the use of this code.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'To put the icon in the SysTray
Private Declare Function Shell_NotifyIcon Lib "SHELL32" Alias "Shell_NotifyIconA" _
        (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
    cbsize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Const NIM_ADD = 0
Const NIM_MODIFY = 1
Const NIM_DELETE = 2
Const NIF_MESSAGE = 1
Const NIF_ICON = 2
Const NIF_TIP = 4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_RBUTTONUP = &H205

Dim nid As NOTIFYICONDATA

'so when we return to the form, it will be on top
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

'To determine if the user is trying to close the app
Dim blnShutDown As Boolean

Private Sub about_Click()
'Open form2
Form2.Show
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
    
    nid.hwnd = Me.hwnd
    nid.uID = vbNull
    nid.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    nid.uCallbackMessage = WM_LBUTTONDOWN
    nid.hIcon = Me.Icon
    nid.szTip = Me.Caption & Chr$(0)
    nid.cbsize = Len(nid)
    
    Text1.Text = GetSetting(App.CompanyName, App.Title, "Name", vbNullString)
    Text2.Text = GetSetting(App.CompanyName, App.Title, "Pass", vbNullString)
    If Text1.Text <> vbNullString Or Text2.Text <> vbNullString Then Check1.Value = 1
End Sub

'Hide form, show icon, and stop timer when minimized
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Me.Hide
        Call Shell_NotifyIcon(NIM_ADD, nid)
    End If
End Sub

'Show form, remove icon, when the icon is doubleclicked (Resize won't fire if form is hidden, so code goes here)
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Msg As Long
    Msg = X / Screen.TwipsPerPixelX
    Select Case Msg
        Case WM_LBUTTONDBLCLK
            Call mnuOpen_Click
        Case WM_RBUTTONUP
            Me.PopupMenu mnuPop
    End Select
End Sub

Private Sub home_Click()
'Open Internet Explorer
Dim IE
Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = True
strURL = "http://www.ebay.com"
IE.Navigate strURL
'Open Internet Explorer
End Sub

Private Sub mnuExit_Click()
    Call Shell_NotifyIcon(NIM_DELETE, nid)
    Call cmdExit_Click
End Sub

Private Sub mnuOpen_Click()
    Me.WindowState = vbNormal
    BringWindowToTop Me.hwnd
    Me.Show
    AppActivate Me.Caption
    Call Shell_NotifyIcon(NIM_DELETE, nid)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Check1.Value = 1 Then
        Call SaveSetting(App.CompanyName, App.Title, "Name", Text1.Text)
        Call SaveSetting(App.CompanyName, App.Title, "Pass", Text2.Text)
    Else
        Call SaveSetting(App.CompanyName, App.Title, "Name", vbNullString)
        Call SaveSetting(App.CompanyName, App.Title, "Pass", vbNullString)
    End If
End Sub

Private Sub cmdSubmit_Click()
    If blnShutDown = False Then lblStatus.Caption = "Preparing to Connect to eBay..."
    WebBrowser1.Navigate "http://signin.ebay.com/aw-cgi/eBayISAPI.dll?SignIn"
End Sub

Private Sub cmdExit_Click()
    blnShutDown = True
    Call cmdSubmit_Click
End Sub



Private Sub myebay_Click()
'Open Internet Explorer
Dim IE
Set IE = CreateObject("InternetExplorer.Application")
strURL = "http://cgi1.ebay.com/aw-cgi/eBayISAPI.dll?MyEbayItemsBiddingOn&first=N&sellersort=3&biddersort=3&watchsort=3&dayssince=2&p1=0&p2=0&p3=0&p4=0&p5=0&update=0&nitem=0&rows=25&pagebid=1&pagewon=1&pagelost=1&SaveSetting=off&ssPageName=MerchMax&pass=6toUUKehQV18C.YJlEdBy.&userid=" & Text1.Text
IE.Navigate strURL, vbNormalFocus
'Open Internet Explorer
End Sub

Private Sub paypal_Click()
'Open Internet Explorer
Dim IE
Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = True
strURL = "http://www.paypal.com"
IE.Navigate strURL
'Open Internet Explorer
End Sub

Private Sub purchase_Click()
'Open Internet Explorer
Dim IE
Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = True
strURL = "https://www.paypal.com/xclick/business=feedback%40xssc.com&item_name=eBay+Login+Utility&item_number=eBay+Login+Utility&amount=2.95&no_note=1&currency_code=USD"
IE.Navigate strURL
'Open Internet Explorer
End Sub

Private Sub readme_Click()
Call Shell("notepad.exe " & App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "readme.txt", vbNormalFocus)
End Sub

Private Sub register_Click()
'Open Internet Explorer
Dim IE
Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = True
strURL = "http://www.xssc.com/ebaytools/registration.htm"
IE.Navigate strURL
'Open Internet Explorer
End Sub

Private Sub reset_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub time_Click()
'Open Internet Explorer
Dim IE
Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = True
strURL = "http://cgi3.ebay.com/aw-cgi/eBayISAPI.dll?TimeShow"
IE.Navigate strURL
'Open Internet Explorer
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    If URL <> vbNullString And URL <> "http:///" Then
        If Mid$(URL, InStr(1, URL, "?") + 1) = "SignIn" And blnShutDown = False Then 'http://signin.ebay.com/
            WebBrowser1.Document.All("userid").Value = Text1.Text
            WebBrowser1.Document.All("pass").Value = Text2.Text
            WebBrowser1.Document.All("keepMeSignInOption").Checked = True
            WebBrowser1.Document.All("SignInForm").submit
            lblStatus.Caption = "Connecting to eBay..."
        Else
            lblStatus.Caption = "eBay Login Successful"
        End If
        If blnShutDown = True Then Unload Me
    End If
End Sub
