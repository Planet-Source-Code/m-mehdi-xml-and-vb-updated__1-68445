VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "XML and VB (Enhanced)"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1335
      Left            =   60
      TabIndex        =   4
      Top             =   4380
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2355
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Read From XML using ADODB.Recordset"
      Height          =   495
      Left            =   3150
      TabIndex        =   3
      Top             =   120
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   690
      Width           =   5865
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Read From XML using xmlDOM"
      Height          =   495
      Left            =   1590
      TabIndex        =   1
      Top             =   120
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save XML File From Recordset"
      Height          =   495
      Left            =   30
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
'// Libraries Required //
'1. Microsoft XML, v5.0
'2. Microsoft ActiveX Data Objects 2.6 Library
'////////////////////////

Dim objXML As MSXML2.DOMDocument50
Dim objRs As ADODB.Recordset
Dim objConn As ADODB.Connection

Private Sub Command1_Click()

Set objConn = New ADODB.Connection
Set objRs = New ADODB.Recordset
Set objXML = New MSXML2.DOMDocument50


objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
             "Data Source=" & App.Path & "\database.mdb;" & _
             "User Id=admin;" & _
             "Password="

objRs.CursorLocation = adUseClient
objRs.Open "Select * from Client", objConn, adOpenKeyset, adLockReadOnly, adCmdText
objRs.save objXML, adPersistXML

objXML.async = False ' first load the full document before rendering
objXML.validateOnParse = True ' verifying and parsing the schema info
objXML.save App.Path & "\textXML.xml"
MsgBox "File saved as xml at location " & vbCrLf & App.Path & "\textXML.xml"

objRs.Close
objConn.Close

Set objRs = Nothing
Set objConn = Nothing
Set objXML = Nothing

End Sub

Private Sub Command2_Click()

Text1.Text = Clear
Set objXML = New MSXML2.DOMDocument50
objXML.Load App.Path & "\cd_catalog.xml"

Dim strText As String
Dim objNodeList As IXMLDOMNodeList

strText = "Bob Dylan"
Set objNodeList = objXML.selectNodes("CATALOG/CD[ARTIST='" & strText & "']")

Dim x, y
For Each x In objNodeList
    For Each y In x.childNodes
        Text1.Text = Text1.Text & vbCrLf & y.nodeName & " - " & y.Text
    Next
Next

End Sub

Private Sub Command3_Click()
Set objRs = New ADODB.Recordset
objRs.CursorLocation = adUseClient
objRs.Open App.Path & "\textXML.xml"
Dim intC As Integer, intD As Integer
For intC = 0 To objRs.RecordCount - 1
    For intD = 0 To objRs.fields.Count - 1
        Text1.Text = Text1.Text & vbCrLf & objRs.fields(intD).Name & Space(5) & objRs.fields(intD).Value
    Next intD
    objRs.MoveNext
Next intC
Set DataGrid1.DataSource = objRs
DataGrid1.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
If objRs.State = 1 Then objRs.Close
Set objRs = Nothing
End Sub
