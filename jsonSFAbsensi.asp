<!--#include file="connSmartFarming.inc" -->


<% 
	
 On Error Resume Next 

	'membuat query'                   
	'                          
	querytbl = "insert into tblabsensi (kdid, tgl, nomaster, iddevice, koordinatx, koordinaty, lokasifile,lokasi) 	values ('" & Request.QueryString("kdid") & "',isnull('" & Request.QueryString("tgl") & "',getdate()),'"	& Request.QueryString("nomaster") & "','" & Request.QueryString("iddevice") & "','" & Request.QueryString("koordinatx") & "','" & Request.QueryString("koordinaty") & "','" & Request.QueryString("lokasifile") & ".JPEG','" & Request.QueryString("lokasi") & "')"
		
	'Ambil data'
	set rd = server.CreateObject("ADODB.RECORDSET")
	rd.Open querytbl, Conn,3,1
	'response.write("sukses")'		

	response.write ("query result : " & querytbl)

%>
