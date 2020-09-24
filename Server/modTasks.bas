Attribute VB_Name = "modTasks"
Public Enum AWAYSTATUS
    msgNONE = 0
    msgAWAY = 1
    msgBUSY = 2
    msgBRB = 3
End Enum

Type Client
    Auth As Boolean
    User As String
    Status As AWAYSTATUS
    iCount As Long
    rREC As String
    Pong As Boolean
    Version As String
End Type

Public Clients() As Client

Public db As New ADODB.Connection
Function EncryptPassword(Number As Byte, _
    DecryptedPassword) As String
    Dim Password As String
    Dim temp As Integer
    Counter = 1
    Do Until Counter = Len(DecryptedPassword) + 1
        temp = Asc(Mid(DecryptedPassword, Counter, 1))
        'see if even
        If Counter Mod 2 = 0 Then
            temp = temp - Number
        Else
            temp = temp + Number
        End If
        temp = temp Xor (10 - Number)
        Password = Password & Chr$(temp)
        Counter = Counter + 1
    Loop
    EncryptPassword = Password
End Function






Function GetPiece(from, delim, Index) As String
    Dim temp$
    Dim Count
    Dim Where
    '
    temp$ = from & delim
    Where = InStr(temp$, delim)
    Count = 0
    Do While (Where > 0)
        Count = Count + 1
        If (Count = Index) Then
            GetPiece = Left$(temp$, Where - 1)
            Exit Function
        End If
        temp$ = Right$(temp$, Len(temp$) - Where)
        Where = InStr(temp$, delim)
    'Doevents
    Loop
    If (Count = 0) Then
        GetPiece = from
    Else
        GetPiece = ""
    End If
End Function


Function DecryptPassword(Number As Byte, _
    EncryptedPassword) As String
    Dim Password As String
    Dim temp As Integer
    Counter = 1
    Do Until Counter = Len(EncryptedPassword) + 1
        temp = Asc(Mid(EncryptedPassword, _
            Counter, 1)) Xor (10 - Number)
        'see if even
        If Counter Mod 2 = 0 Then
            temp = temp + Number
        Else
            temp = temp - Number
        End If
        Password = Password & Chr$(temp)
        Counter = Counter + 1
    Loop
    DecryptPassword = Password
End Function


Sub Main()


'CHANGE THE SERVER THROUGH HERE!
strServer = "localhost"


If App.PrevInstance Then End

ConnectionString = "Driver={Microsoft Access Driver (*.mdb)};Dbq=mdbAccess.mdb;DefaultDir=" & App.Path & ";UID=;PWD="

'ConnectionString = "driver={SQL Server};" & _
 '       "server=" & strServer & ";" & _
  '      "uid=sa;pwd=;" & _
   '     "database=vsag;"
        
db.ConnectionString = ConnectionString
db.ConnectionTimeout = 10
db.Open
Load frmMain


End Sub


