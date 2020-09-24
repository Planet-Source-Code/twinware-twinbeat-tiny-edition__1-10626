VERSION 5.00
Object = "{24365B29-A3B5-11D1-B8B0-444553540000}#1.0#0"; "XFXFORMS.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   750
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   1770
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   1770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   960
   End
   Begin xfxFormShaper.FormShaper FormShaper1 
      Left            =   600
      Top             =   960
      _ExtentX        =   1852
      _ExtentY        =   1296
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H0080FFFF&
      Height          =   375
      Left            =   240
      Shape           =   4  'Gerundetes Rechteck
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblBeat 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Menu ee 
      Caption         =   "ee"
      Visible         =   0   'False
      Begin VB.Menu mnuoptions 
         Caption         =   "&Optionen"
         Begin VB.Menu mnuoptionsummy 
            Caption         =   "&Zeit/Datum stellen"
         End
         Begin VB.Menu mnuoptionsInet 
            Caption         =   "&Besuchen Sie www.TwinWare.de"
         End
         Begin VB.Menu mnuoptionssendmail 
            Caption         =   "&Senden Sie uns eine E-Mail"
         End
      End
      Begin VB.Menu mnuhelp 
         Caption         =   "&?"
         Begin VB.Menu mnuhelpinfo 
            Caption         =   "&Über TwinBeat TE"
         End
      End
      Begin VB.Menu mnubrake 
         Caption         =   "-"
      End
      Begin VB.Menu mnuoptionsexit 
         Caption         =   "&Beenden"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declare the function for letting the form stay on top
Private Declare Function SetWindowPos Lib "user32" _
 (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x _
 As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As _
 Long, ByVal wFlags As Long) As Long
 
'Declare the variables for let the form stay on top
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Const SWP_SHOWWINDOW = &H40
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2


'This is the Shell function for opening wesites, sending mails and open Programs or Files etc.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



Public Function GetTimeZone(Optional ByRef strTZName As String) As Long

'Get timezone
    Dim objTimeZone As TIME_ZONE_INFORMATION
    Dim lngResult As Long
    Dim i As Long
    lngResult = GetTimeZoneInformation&(objTimeZone)


    Select Case lngResult
        Case 0&, 1& 'use standard time
        GetTimeZone = -(objTimeZone.Bias + objTimeZone.StandardBias) 'into minutes


        For i = 0 To 31
            If objTimeZone.StandardName(i) = 0 Then Exit For
            strTZName = strTZName & Chr(objTimeZone.StandardName(i))
        Next

        Case 2& 'use daylight savings time
        GetTimeZone = -(objTimeZone.Bias + objTimeZone.DaylightBias) 'into minutes


        For i = 0 To 31
            If objTimeZone.DaylightName(i) = 0 Then Exit For
            strTZName = strTZName & Chr(objTimeZone.DaylightName(i))
        Next

    End Select

End Function


Public Function InternetTime()

    Dim tmpH                '\
    Dim tmpS                ' \
    Dim tmpM                '  \
    Dim itime               '   \
    Dim tmpZ                '    \
    Dim testtemp As String  '=====> Declare variables for the interntetime
    tmpH = Hour(Time)       '    /
    tmpM = Minute(Time)     '   /
    tmpS = Second(Time)     '  /
    tmpZ = GetTimeZone      '_/
    
    'calculate internettime
    itime = ((tmpH * 3600 + ((tmpM - tmpZ + 60) * 60) + tmpS) * 1000 / 86400)

    'Check out for inettime = 1000 ...
    If itime = 1000 Then
        itime = itime - 1000
    ElseIf itime < 0 Then
        itime = itime + 1000
    End If
    
    'Do I have to say something???
    InternetTime = itime
    
End Function

Private Sub Form_Click()
    
    'Call the Menu "ee" by clicking on the form (you can also rename it if you want to!)
    PopupMenu ee
    
End Sub

Private Sub Form_Load()
    
    'Keeps TwinBeatTE on top of your desktop always...
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
    SWP_NOSIZE + SWP_NOMOVE + SWP_SHOWWINDOW
    
    'Get Settings out of the application title
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 1272)
    
    'Shape the form (!You need the .ocx file!)
    FormShaper1.ShapeIt
    
    
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    'Save Settings in the application title
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting App.Title, "Settings", "MainTop", Me.Top


End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    
    'Save Settings in the application title
    SaveSetting App.Title, "Settings", "MainLeft", Me.Left
    SaveSetting App.Title, "Settings", "MainTop", Me.Top
    

    
        
        
End Sub

Private Sub lblBeat_Click()
    
    'Call the Menu "ee" by clicking on the form (you can also rename it if you want to!)
    PopupMenu ee
    
    
End Sub



Private Sub mnuhelpinfo_Click()
    
    ' The important part of this procedure is the part below:
    '               |----------------------------|
    ShellExecute 0, "Open", App.Path & "\info.htm", "", "", vbNormalFocus
    ' Thats the command!

End Sub

Private Sub mnuoptionsexit_Click()
    
    Close ' I don't know why I put Close here, I learned it
          ' in that way so do it or not...
    End
    
    
End Sub

Private Sub mnuoptionsInet_Click()
    
    ' The important part of this procedure is the part below:
    '               |-------------------------------|
    ShellExecute 0, "Open", "http://www.TwinWare.de", "", "", vbNormalFocus
    'Thats the command!

End Sub

Private Sub mnuoptionssendmail_Click()
    
    ' The important part of this procedure is the part below:
    '               |------------------------------------|
    ShellExecute 0, "Open", "mailto: support@twinware.de?Subject=TwinBeat TE", "", "", vbNormalFocus
    'Thats the command!

End Sub

Private Sub mnuoptionsummy_Click()
    
    Form2.Show 1 'Show the form2 form
    
End Sub

Private Sub Timer1_Timer()
    'Caption of lblBeat is the InternetTime + @ (The @ was set by Swath!)
    lblBeat.Caption = "@ " & (CInt(InternetTime))
    
    
End Sub

