<% 
Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Driver={SQL Server};" & _ 
	      "Server=SRV-TIMBANGAN;" & _ 
   	      "Database=tebu;" & _
	      "Uid=sa;" & _ 
 	      "Pwd=sukseskan;" 



%>