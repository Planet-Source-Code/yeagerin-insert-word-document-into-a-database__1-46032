Attribute VB_Name = "Module1"
Public Function SaveFileToDB(ByVal FileName As String, FieldName As String, Tabla As String, Source As String) As Boolean
'**************************************************************
'PURPOSE: SAVES DATA FROM BINARY FILE (e.g., .EXE, WORD DOCUMENT
'CONTROL TO RECORDSET RS IN FIELD NAME FIELDNAME

'FIELD TYPE MUST BE BINARY (OLE OBJECT IN ACCESS)

'REQUIRES: REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS 2.0 or ABOVE

'SAMPLE USAGE
Dim sConn As String
Dim oConn As New ADODB.Connection
Dim oRs As New ADODB.Recordset
'
'
sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Source & ";Persist Security Info=False"
'
oConn.Open sConn
oRs.Open "SELECT * FROM " & Tabla, oConn, adOpenKeyset, adLockOptimistic
oRs.AddNew


Dim iFileNum As Integer
Dim lFileLength As Long

Dim abBytes() As Byte
Dim iCtr As Integer

'On Error GoTo ErrorHandler
If Dir(FileName) = "" Then Exit Function
If Not TypeOf oRs Is ADODB.Recordset Then Exit Function

'read file contents to byte array
iFileNum = FreeFile
Open FileName For Binary Access Read As #iFileNum
lFileLength = LOF(iFileNum)
ReDim abBytes(lFileLength)
Get #iFileNum, , abBytes()

'put byte array contents into db field
oRs.Fields(FieldName).AppendChunk abBytes()
Close #iFileNum
oRs.Update
oRs.Close

SaveFileToDB = True

ErrorHandler:
End Function

Public Function LoadFileFromDB(ByVal FileName As String, FieldName As String, Tabla As String, Source As String) As Boolean
'************************************************
'PURPOSE: LOADS BINARY DATA IN RECORDSET RS,
'FIELD FieldName TO a File Named by the FileName parameter

'REQUIRES: REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS 2.0 or ABOVE

'SAMPLE USAGE
Dim sConn As String
Dim oConn As New ADODB.Connection
Dim oRs As New ADODB.Recordset
'
'
sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Source & ";Persist Security Info=False"

oConn.Open sConn
oRs.Open "SELECT * FROM " & Tabla, oConn, adOpenKeyset, adLockOptimistic

Dim iFileNum As Integer
Dim lFileLength As Long
Dim abBytes() As Byte
Dim iCtr As Integer
Dim contador As Integer 'contador de documentos
contador = 0
On Error GoTo ErrorHandler
While Not oRs.EOF
    If TypeOf oRs Is ADODB.Recordset Then
        iFileNum = FreeFile
        Nombre = FileName & contador & ".doc"
        Open Nombre For Binary As #iFileNum
        lFileLength = LenB(oRs(FieldName))
        abBytes = oRs(FieldName).GetChunk(lFileLength)
        Put #iFileNum, , abBytes()
        Close #iFileNum
        contador = contador + 1
    End If
    oRs.MoveNext
Wend
LoadFileFromDB = True
ErrorHandler:
End Function
