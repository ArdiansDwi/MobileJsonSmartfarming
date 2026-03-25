<% 
Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Driver={SQL Server};" & _ 
	      "Server=192.168.0.101;" & _ 
   	      "Database=SmartFarming;" & _
	      "Uid=sa;" & _ 
 	      "Pwd=bululawang2014#;" 



%>