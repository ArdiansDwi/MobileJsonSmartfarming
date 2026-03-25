<!--#include file="connSmartFarming.inc" -->


<% 
	
 On Error Resume Next 

	'membuat query'                   
	'                          
	querytbl = "insert into tblbrix (kdid, tglentry, user_id, nopetak, register, koordinatx, koordinaty, lokasifile, lokasi, brix, kecamatan, mg) 	values ('" & Request.QueryString("kdid") & "',isnull('" & Request.QueryString("tgl") & "',getdate()),'" & Request.QueryString("user_id") & "','" & Request.QueryString("nopetak") & "','" & Request.QueryString("register") & "','" & Request.QueryString("koordinatx") & "','" & Request.QueryString("koordinaty") & "','" & Request.QueryString("lokasifile") & ".JPEG','" & Request.QueryString("lokasi") & "', '" & Request.QueryString("brix") & "', '"& Request.QueryString("kecamatan") &"', '"& Request.QueryString("mg") &"')"
		
	'Ambil data'
	set rd = server.CreateObject("ADODB.RECORDSET")
	rd.Open querytbl, Conn,3,1
	'response.write("sukses")'		

	response.write ("query result : " & querytbl)

%>
