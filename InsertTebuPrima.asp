<!--#include file="connSmartFarming.inc" -->


<% 
	
 On Error Resume Next 

	'membuat query'                   
	'                       
'value = Request.QueryString("tgl")  

	querytbl = "insert into tblkunjungan (kdid, tgl, nomaster, iddevice, koordinatx, koordinaty, kodepetani, namaPetani, Alamat, deskripsi, lokasifile1, lokasifile2, lokasifile3, lokasifile4, lokasi, register, varietas, masaTanam, brixAtas) values ('" & Request.QueryString("kdid") & "','" & Request.QueryString("tgl") & "','" & Request.QueryString("nomaster") & "','" & Request.QueryString("iddevice") & "','" & Request.QueryString("koordinatx") & "','" & Request.QueryString("koordinaty") & "','" & Request.QueryString("kodepetani") & "','" & Request.QueryString("namapetani") & "','" & Request.QueryString("alamat") & "','" & Request.QueryString("deskripsi") & "','" & Request.QueryString("lokasifile1") & "','" & Request.QueryString("lokasifile2") & "','" & Request.QueryString("lokasifile3") & "','" & Request.QueryString("lokasifile4") & "','" & Request.QueryString("lokasi") & "','" & Request.QueryString("register") & "','" & Request.QueryString("varietas") & "','" & Request.QueryString("masatanam") & "','" & Request.QueryString("brixatas") & "')"
		
	'Ambil data'
	set rd = server.CreateObject("ADODB.RECORDSET")
	rd.Open querytbl, Conn,3,1
	'response.write("sukses")'		

	response.write ("query result : " & querytbl)

%>
