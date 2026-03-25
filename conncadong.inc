<% 
Set Conn = Server.CreateObject("ADODB.Connection")
    Conn.Open "Driver={SQL Server};" & _ 
	      "Server=192.168.0.199;" & _ 
   	      "Database=KebunBibit;" & _
	      "Uid=sa;" & _ 
 	      "Pwd=ar1;" 



%>