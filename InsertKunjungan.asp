<!--#include file="connSmartFarming.inc" -->


<% 
	
 On Error Resume Next 
	
	'membuat query'   
	querytbl = "insert into tblkunjungan (kdid, tgl, nomaster, iddevice, koordinatx, koordinaty, kodepetani, namaPetani, Alamat, judul, deskripsi, lokasifile1, lokasifile2, lokasifile3, lokasifile4, lokasi) values ('" & Request.QueryString("kdid") & "','" & Request.QueryString("tgl") & "','" & Request.QueryString("nomaster") & "','" & Request.QueryString("iddevice") & "','" & Request.QueryString("koordinatx") & "','" & Request.QueryString("koordinaty") & "','" & Request.QueryString("kodepetani") & "','" & Request.QueryString("namapetani") & "','" & Request.QueryString("alamat") & "','" & Request.QueryString("judul") & "','" & Request.QueryString("deskripsi") & "','" & Request.QueryString("lokasifile1") & "','" & Request.QueryString("lokasifile2") & "','" & Request.QueryString("lokasifile3") & "','" & Request.QueryString("lokasifile4") & "','" & Request.QueryString("lokasi") & "')"
		
	'Ambil data'
	set rd = server.CreateObject("ADODB.RECORDSET")
	rd.Open querytbl, Conn,3,1
	'response.write("sukses")'		

	response.write ("query result : " & querytbl)

%>
