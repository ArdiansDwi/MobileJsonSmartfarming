<% 
Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Driver={SQL Server};" & _ 
	      "Server=192.168.0.250;" & _ 
   	      "Database=tebu;" & _
	      "Uid=sa;" & _ 
 	      "Pwd=sukseskan;" 



%>