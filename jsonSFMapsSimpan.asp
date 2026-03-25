<!--#include file="connSmartFarming.inc" -->


<% 
	
 On Error Resume Next 

	'membuat query'                   
	'                  
	if (Request.QueryString("register") = "Pilih Register") then       

		response.write ("Pilih Register Dahulu")
		
	elseif (Request.QueryString("jnslahan") = "Pilih Jenis Lahan") then       

		response.write ("Pilih Jenis Lahan Dahulu")

	elseif (Request.QueryString("masatanam") = "Pilih Masa Tanam") then       

		response.write ("Pilih Masa Tanam Dahulu")

	elseif (Request.QueryString("varietas") = "Pilih Varietas") then       

		response.write ("Pilih Varietas Dahulu")

	elseif (Request.QueryString("kategori") = "Pilih Kategori") then       

		response.write ("Pilih Kategori Dahulu")

	else
	

	querytbl = "insert into tblRegisterPetak (tglentry, user_id, register, nopetak, kecamatan, desa, luas, petugas, jnslahan, masatanam, varietas, kategori, mg) values (getdate(),'" & Request.QueryString("user_id") & "','"& Request.QueryString("register") & "','"	& Request.QueryString("nopetak") & "','"	& Request.QueryString("kecamatan") & "','"	& Request.QueryString("desa") & "','"	& Request.QueryString("luas") & "','"	& Request.QueryString("petugas") & "','"	& Request.QueryString("jnslahan") & "','"	& Request.QueryString("masatanam") & "','"	& Request.QueryString("varietas") & "','"	& Request.QueryString("kategori") & "','"	& Request.QueryString("mg") & "')"

		'Ambil data'
		set rd = server.CreateObject("ADODB.RECORDSET")
		rd.Open querytbl, Conn,3,1
		'response.write("sukses")'		
		response.write ("Sukses")
	end if
	'

	

%>
