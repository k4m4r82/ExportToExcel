Attribute VB_Name = "Module1"
Option Explicit

Public Const CHUNK_SIZE     As Long = 16384

Public conn                 As ADODB.Connection
Dim rsImage                 As ADODB.Recordset

Dim i                       As Long
Dim lsize                   As Long
Dim iChunks                 As Long
Dim nFragmentOffset         As Long

Dim nHandle                 As Integer
Dim varChunk()              As Byte

Private Function fileExists(ByVal strNamaFile As String) As Boolean
    If Not (Len(strNamaFile) > 0) Then fileExists = False: Exit Function

    If Dir$(strNamaFile, vbNormal) = "" Then
        fileExists = False
    Else
        fileExists = True
    End If
End Function

Public Sub closeRecordset(ByVal vRs As ADODB.Recordset)
    On Error Resume Next

    If Not (vRs Is Nothing) Then
        If vRs.State = adStateOpen Then
            vRs.Close
            Set vRs = Nothing
        End If
    End If
End Sub

Public Function getImageFromDB(ByVal query As String) As String
    Dim sFile           As String

    On Error GoTo errHandle

    Set rsImage = New ADODB.Recordset
    rsImage.Open query, conn, adOpenForwardOnly, adLockReadOnly
    If Not rsImage.EOF Then
        If Not IsNull(rsImage(0).Value) Then
            nHandle = FreeFile

            sFile = App.Path & "\output.bin"
            If fileExists(sFile) Then Kill sFile
            DoEvents

            Open sFile For Binary Access Write As nHandle

            lsize = rsImage(0).ActualSize
            iChunks = lsize \ CHUNK_SIZE
            nFragmentOffset = lsize Mod CHUNK_SIZE

            varChunk() = rsImage(0).GetChunk(nFragmentOffset)
            Put nHandle, , varChunk()
            For i = 1 To iChunks
                 ReDim varChunk(CHUNK_SIZE) As Byte

                 varChunk() = rsImage(0).GetChunk(CHUNK_SIZE)
                 Put nHandle, , varChunk()
                 DoEvents
            Next
            Close nHandle

            getImageFromDB = sFile
        End If
    End If
    Call closeRecordset(rsImage)

    Exit Function
errHandle:
    getImageFromDB = ""
End Function

