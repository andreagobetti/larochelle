<%@ LANGUAGE="VBSCRIPT" %>
<%
Dim objDict
'Create an instance of the Dictionary object
Set objDict = Server.CreateObject("Scripting.Dictionary") 

'Add some key/item pairs to a Dictionary
objDict.Add "key1","item1" 'Add value item1 with key key1
objDict.Add "key2","item2" 'Add value item2 with key key2
objDict.Add "key3","item3" 'Add value item3 with key key3

'retrieve the item value for the "key1" and print it out 
myItem = objDict.Item("key1")
Response.Write("The value corresponding to the key1 is " & myItem & "<br />") 

'remove the pair key2/item2
myItem = objDict.Remove("key2")

'change the value of the item with the key "key1"
objDict.Item("key1") = "newValue"
Response.Write("The new value corresponding to the key1 is " & objDict.Item("key1") & "<br />") 

'change the value of an existing key "key1" to "newKey", 
'without changing the value of the corresponding item
objDict.Key("key1") = "newKey"
Response.Write("The value corresponding to the newKey is " & objDict.Item("newKey") & "<br />") 

'retrieve all the keys and items from the Dictionary and print them out
allKeys = objDict.Keys   'Get all the keys into an array
allItems = objDict.Items 'Get all the items into an array 

For i = 0 To objDict.Count - 1 'Iterate through the array
  myKey = allKeys(i)   'This is the key value
  myItem = allItems(i) 'This is the item value
  Response.Write("The " & i & " value in the Dictionary is " & myItem & "<br />")
Next

'remove all the items
objDict.RemoveAll

'free up resources
set objDict=nothing
%>