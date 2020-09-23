VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00AFAFAF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cellphone Carrier Finder"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "whitepages.com"
      RemotePort      =   80
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00AFAFAF&
      Caption         =   "Find"
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   11233330
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cellphone Number"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Location"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Carrier"
         Object.Width           =   4586
      EndProperty
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00AB6832&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   3600
      Width           =   4575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(e.g. ""5551234567"")"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cellphone Number:"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WholeData As String
Dim CellNumber As String
Private Sub Command1_Click()
If Len(Text1.Text) <> 10 Then Status.Caption = "Invalid Number -- Must be 10 numbers long": Exit Sub
Command1.Enabled = False
WholeData = ""
CellNumber = Text1.Text
Winsock1.Close
Winsock1.Connect
Status.Caption = "Finding Carrier..."
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub
Private Sub Winsock1_Close()
Winsock1.Close
ProcessData WholeData
End Sub
Private Sub Winsock1_Connect()
Winsock1.SendData "GET /search/Reverse_Phone?phone=" & CellNumber & vbCrLf & vbCrLf
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
Winsock1.GetData Data
WholeData = WholeData & Data
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Command1.Enabled = True
Winsock1.Close
Status.Caption = Description
End Sub
Sub ProcessData(Data As String)
Dim CellNum As String
Dim Location As String
Dim Carrier As String
On Error GoTo ErrorHandler
CellNum = Get_Between(1, Data, "The phone number " & Chr(34), Chr(34))
Location = Get_Between(1, Data, "The phone number " & Chr(34) & CellNum & Chr(34) & " is a", "based phone number")
Carrier = Get_Between(1, Data, "based phone number and the registered carrier is ", Chr(10))
Set AddData = ListView1.ListItems.Add(, , CellNum)
Location = Replace(Location, Chr(9), "")
Location = Replace(Location, Chr(10), "")
Location = Mid(Location, 2, Len(Location) - 2)
Carrier = Replace(Carrier, "&amp;", "&")
AddData.SubItems(1) = Location
AddData.SubItems(2) = Carrier
Command1.Enabled = True
Status.Caption = "Carrier Found"
Exit Sub
ErrorHandler:
Status.Caption = "Error Finding Carrier"
Command1.Enabled = True
End Sub
