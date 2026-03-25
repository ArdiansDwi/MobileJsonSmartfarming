<!--#include file="connKendaraan.inc" -->

<% 
 On Error Resume Next 
'membuat query'
userid=request.querystring("user_id")
kd_permintaan=request.querystring("kd_permintaan")

						querytbl =  "update tblPermintaanBBM set acc_skw=1, tglacc_skw=getdate(), acc_skk=1, tglacc_skk=getdate(), acc_ca=1, tglacc_ca=getdate(),acc_oleh='" &userid& "' where kd_permintaan='" &kd_permintaan& "'"

	'Ambil data'
	set rd = server.CreateObject("ADODB.RECORDSET")
	rd.Open querytbl, Conn,3,1
	'response.write("sukses")'		

	response.write ("query result : " & querytbl)

%>