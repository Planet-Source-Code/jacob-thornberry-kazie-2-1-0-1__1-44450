Attribute VB_Name = "RegEdit"
Public Sub CreateKey(Folder As String, Value As String)

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
b.RegWrite Folder, Value

End Sub

Public Sub CreateIntegerKey(Folder As String, Value As Integer)

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
b.RegWrite Folder, Value, "REG_DWORD"


End Sub

Public Function ReadKey(Value As String) As String

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
r = b.RegRead(Value)
ReadKey = r
End Function


Public Sub DeleteKey(Value As String)

Dim b As Object
On Error Resume Next
Set b = CreateObject("Wscript.Shell")
b.RegDelete Value
End Sub
