VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Utilitário Ping"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkUnicos 
      Caption         =   "Only diferent hosts"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   3135
   End
   Begin MSFlexGridLib.MSFlexGrid GridHosts 
      Height          =   1815
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   -2147483633
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   ""
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAdicionar 
      Caption         =   "&Add"
      Default         =   -1  'True
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdSobre 
      Caption         =   "A&bout"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton cmdTempo 
      Caption         =   "&Time"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox cmpHost 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "www.altavista.com"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Timer tmrTempo 
      Interval        =   1000
      Left            =   3240
      Top             =   2160
   End
   Begin VB.Label Label4 
      Caption         =   "seconds."
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lblTempo 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   255
      Left            =   1560
      TabIndex        =   4
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Verifying every"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Add  host(IP or URL):"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub EchoIt(ByVal IP As String, Optional ListPos)
   
   If Len(IP) < 7 Or InStr(1, IP, ".") = 0 Then Exit Sub
   
   lblStatus = IP
   
   Dim ECHO As ICMP_ECHO_REPLY
   Dim Pos As Integer, StatusCode As String
   
   
   Call Ping(Trim(IP), ECHO)
   
   
   StatusCode = GetStatusCode(ECHO.status)
   
   
End Sub
Public Sub cmdAdicionar_Click()
Me.MousePointer = vbHourglass
tmrTempo.Enabled = False
If Not IsNumeric(cmpHost.Text) Then
   cmpHost.Text = AddrByName(cmpHost.Text)
End If
For host = 1 To GridHosts.Rows
    If GridHosts.TextMatrix(host - 1, 0) = cmpHost.Text And chkUnicos.Value <> 0 Then
       MsgBox "This host is already being monitored", vbOKOnly + vbInformation, "Host already watched"
       cmpHost.Text = ""
       Exit Sub
    End If
Next host
GridHosts.Rows = GridHosts.Rows + 1
GridHosts.TextMatrix(GridHosts.Rows - 1, 0) = cmpHost.Text
GridHosts.TextMatrix(GridHosts.Rows - 1, 1) = "Offline"
GridHosts.TextMatrix(GridHosts.Rows - 1, 2) = "n/a"
GridHosts.TextMatrix(GridHosts.Rows - 1, 3) = NameByAddr(GridHosts.TextMatrix(GridHosts.Rows - 1, 0))
If GridHosts.TextMatrix(GridHosts.Rows - 1, 3) = "" Then
   GridHosts.TextMatrix(GridHosts.Rows - 1, 3) = "Host ocult or unavailable"
End If
cmpHost.Text = ""
tmrTempo.Enabled = True
Me.MousePointer = vbNormal
End Sub
Private Sub cmdRemover_Click()
If GridHosts.Rows = 2 Then
   GridHosts.Rows = 1
ElseIf GridHosts.Rows = 1 Then
   Exit Sub
Else
   GridHosts.RemoveItem (GridHosts.Row)
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSobre_Click()
MsgBox "KPing, version 1.0" + vbNewLine + "By: Kerplunk_boy!" + vbNewLine + "E-mail: jfk@faccat.br", vbOKOnly + vbInformation, "About..."
End Sub

Private Sub cmdTempo_Click()
tempo = InputBox("How many seconds beetwin refreshing?", "Time interval")
tmrTempo.Interval = CLng(tempo) * 1000
Me.lblTempo = tempo
End Sub

Private Sub Form_Load()
IP_Initialize
GridHosts.Rows = 1
GridHosts.ColWidth(0) = 1600
GridHosts.ColWidth(1) = 2100
GridHosts.ColWidth(2) = 500
GridHosts.ColWidth(3) = 2500
GridHosts.TextMatrix(0, 0) = "IP"
GridHosts.TextMatrix(0, 1) = "Status"
GridHosts.TextMatrix(0, 2) = "Rate"
GridHosts.TextMatrix(0, 3) = "Host"
lblTempo.Caption = tmrTempo.Interval / 1000
End Sub
Private Sub Form_Unload(Cancel As Integer)
WSACleanup
End Sub
Private Sub tmrTempo_Timer()
For host = 2 To GridHosts.Rows
    If GridHosts.TextMatrix(host - 1, 0) = "IP" Then
       Exit Sub
    Else
       EchoIt (GridHosts.TextMatrix(host - 1, 0))
       GridHosts.TextMatrix(host - 1, 1) = Estado_host
       GridHosts.TextMatrix(host - 1, 2) = IIf(Estado_host = "Online", Time_rate, "n/a")
    End If
Next host
End Sub
