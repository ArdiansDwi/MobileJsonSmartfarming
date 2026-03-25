<!--#include file="connSmartFarming.inc" -->



<% 


db = "Driver={SQL Server};Server=192.168.0.101;Database=SmartFarming;Uid=sa;Pwd=bululawang2014#;" 

sql = "INSERT INTO tbltest (tes1,tes2) VALUES ('1',100)"

Set oConn=Server.CreateObject("ADODB.Connection")
oConn.Open db
oConn.Execute sql

oConn.close
Set oConn=nothing

response.write("sukses")
			

%>

