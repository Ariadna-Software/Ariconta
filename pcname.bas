Attribute VB_Name = "pcname"
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

Public Const MAX_COMPUTERNAME_LENGTH = 255


Public Function ComputerName() As String
    'Devuelve el nombre del equipo actual
    Dim sComputerName As String
    Dim ComputerNameLength As Long
    
    On Error GoTo ECN
    sComputerName = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
    ComputerNameLength = MAX_COMPUTERNAME_LENGTH
    Call GetComputerName(sComputerName, ComputerNameLength)
     ComputerName = Mid(sComputerName, 1, ComputerNameLength)
    
    Exit Function
ECN:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Fun: ComputerName"
    ComputerName = DesdeFichero
End Function


Private Function DesdeFichero() As String
Dim NF As Integer
Dim C As String

On Error GoTo EDesdeFichero
    DesdeFichero = "ERROR"
    NF = FreeFile
    Open App.path & "\NomPC.dat" For Input As #NF
    Line Input #NF, C
    Close #NF
    DesdeFichero = C
    Exit Function
EDesdeFichero:
    MsgBox Err.Number & " - " & Err.Description, vbCritical, "Fun: ComputerName"
End Function

