<% 
Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Driver={SQL Server};" & _ 
	      "Server=192.168.0.2;" & _ 
   	      "Database=Kendaraan;" & _
	      "Uid=sa;" & _ 
 	      "Pwd=sukseskan2;" 



%>