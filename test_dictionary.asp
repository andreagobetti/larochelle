<%
 Dim d
 Set d=Server.CreateObject("Scripting.Dictionary")
 d.Add "re","Red"
 d.Add "bl","Blue"
 chiave="gr"
 valore="Green"
 d.Add chiave,valore
 chiave="pi"
 valore="Pink"
 d.Add chiave,valore
 Response.Write("The value of key gr is: " & d.Item("gr"))
 %>


