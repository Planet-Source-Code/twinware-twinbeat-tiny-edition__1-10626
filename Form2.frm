VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Zeit/Datum stellen..."
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "&Abbrechen"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H000000FF&
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Übernehmen"
      Height          =   495
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Bitte Datum in Tag-Monat-Jahr angeben!"
      Top             =   1320
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Zentriert
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Text            =   "Bitte Zeit in HH:MM:SS angeben!"
      ToolTipText     =   "BItte Zeit in HH:MM:SS angeben!"
      Top             =   480
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4320
      Top             =   120
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aktuelles Datum:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Aktuelle Zeit:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    On Error Resume Next
    Date = Text2.Text 'Set new Date
    Time = Text1.Text 'Set new Time
    
End Sub

Private Sub Command2_Click()
    
    Unload Me 'Unload this form
    
End Sub

Private Sub Timer1_Timer()
    
    Label1(0).Caption = "Aktuelle Zeit: " & Time$  'New time
    Label1(1).Caption = "Aktuelles Datum: " & Date$ 'New Date
    
End Sub
