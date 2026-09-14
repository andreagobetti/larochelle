<%
'Option Explict

Class MultiDimensionalDictionary
'######################################################################
'Named Dictionary Recordset Object
'######################################################################
Public SetKey
Private Dict
Private AddNewRecord
'######################################################################
Private Sub Class_Initialize
Set Dict = Server.CreateObject("Scripting.Dictionary")
Set AddNewRecord = Server.CreateObject("Scripting.Dictionary")
End Sub
'######################################################################
Private Sub Class_Terminate
Set Dict = Nothing
Set AddNewRecord = Nothing
End Sub
'######################################################################
Public Sub Update
 Dict.Add SetKey, AddNewRecord
Set AddNewRecord = Server.CreateObject("Scripting.Dictionary")
End Sub
'######################################################################
Public Function GetCollection
Set GetCollection = Dict
End Function
'######################################################################
Public Sub SetField(byVal Key, byVal Value)
If AddNewRecord.Exists(Key) = False Then
 AddNewRecord.Add Key, Value
Else
 AddNewRecord(Key) = Value
End If
End Sub
'######################################################################
Public Function Records
 Records = Dict.Keys
End Function
'######################################################################
Public Function Fields(byVal PrimaryKey)
Fields = Dict(PrimaryKey).Keys
End Function
'######################################################################
Public Function Item(byVal Key, byVal Value)
On Error Resume Next
If Dict.Item(Key).Exists(Value) Then
 Item = Dict.Item(Key).Item(Value)
End If
On Error Goto 0
End Function
'######################################################################
Public Function Exists(byVal Key, byVal Value)
'On Error Resume Next
If IsNull(Value) Or Value = "" Then
 Exists = Dict.Item(Key).Exists
Else
 Exists = Dict.Item(Key).Exists(Value)
End If
'On Error Goto 0
End Function

Public Function Count
 Count = Dict.Count
End Function
'######################################################################
End class
%>