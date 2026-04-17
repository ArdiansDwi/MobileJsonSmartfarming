<!--#include file="connSmartFarming.inc" -->
<%



On Error Resume Next
'membuat query'   
	querytbl = "UPDATE users_id SET password = '" & Request.QueryString("password") & "' WHERE user_id = '" & Request.QueryString("user_id") & "'"
		
	'Ambil data'
	set rd = server.CreateObject("ADODB.RECORDSET")
	rd.Open querytbl, Conn,3,1		

	response.write ("Sukses")

	


%>