VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo ekspor gambar ke Ms Excel"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEkspor 
      Caption         =   "Ekspor"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Option Explicit

Private Sub addImage(ByVal objWBook As Object, ByVal imageName As String, ByVal kolom As String, ByVal iRow As Long, _
                    ByVal width As Double, ByVal height As Double, _
                    Optional minTop As Integer = 10, Optional plusLeft As Integer = 16, Optional worksheet As Long = 1)
                     
    Dim objPict As Object
    
    Set objPict = objWBook.Worksheets(worksheet).Pictures.Insert(imageName)
    With objPict
        .Top = objWBook.Worksheets(worksheet).Range(kolom & iRow).Top - minTop
        .Left = objWBook.Worksheets(worksheet).Range(kolom & iRow).Left + plusLeft
        .width = width
        .height = height
    End With
    Set objPict = Nothing
End Sub

Private Sub cmdEkspor_Click()
    Dim rs          As ADODB.Recordset
    
    Dim objExcel    As Object
    Dim objWBook    As Object
    Dim objWSheet   As Object

    Dim initRow     As Long
    Dim strSql      As String
    
    On Error GoTo errHandle
    
    Screen.MousePointer = vbHourglass
    DoEvents
    
    'Create the Excel object
    Set objExcel = CreateObject("Excel.application") 'bikin object
    
    'Create the workbook
    Set objWBook = objExcel.Workbooks.Add
    
    Set objWSheet = objWBook.Worksheets(1)
    With objWSheet
        initRow = 5
        
        strSql = "SELECT * FROM siswa"
        Set rs = conn.Execute(strSql)
        If Not rs.EOF Then
            Do While Not rs.EOF
                .cells(initRow, 5) = "NIS"
                .cells(initRow, 6) = ": " & rs("nis").Value
                
                .cells(initRow + 1, 5) = "Nama"
                .cells(initRow + 1, 6) = ": " & rs("nama").Value
                
                .cells(initRow + 2, 5) = "Alamat"
                .cells(initRow + 2, 6) = ": " & rs("alamat").Value
                
                strSql = "SELECT foto FROM siswa WHERE nis = '" & rs("nis").Value & "'"
                Call addImage(objWBook, getImageFromDB(strSql), "C", initRow, 45, 51, 12, 48)
                
                initRow = initRow + 5
                rs.MoveNext
            Loop
        End If
        Call closeRecordset(rs)
    End With
    
    objExcel.Visible = True
    
    If Not objWSheet Is Nothing Then Set objWSheet = Nothing
    If Not objWBook Is Nothing Then Set objWBook = Nothing
    If Not objExcel Is Nothing Then Set objExcel = Nothing
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

errHandle:
    If Not objWSheet Is Nothing Then Set objWSheet = Nothing
    If Not objWBook Is Nothing Then Set objWBook = Nothing
    If Not objExcel Is Nothing Then Set objExcel = Nothing
End Sub

Private Sub Form_Load()
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\sampledb.mdb"
End Sub
