<%
  Set objWSHNetwork = Server.CreateObject("WScript.Network") 
  Response.Write objWSHNetwork.ComputerName
%>