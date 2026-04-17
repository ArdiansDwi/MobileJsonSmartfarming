<!--#include file="connTebu.inc" -->


<% 
	
 On Error Resume Next 

	'membuat query'                   
	'                       
'value = Request.QueryString("tgl")  


	'membuat query'
										'membuat query'
						'querytbl = "select * from users_id where user_id='" & Request.QueryString("user_id") & "' and password='" & Request.QueryString("password") & "'"
						kdspa = REPLACE(Request.QueryString("kdspa"),"_","#")
						'response.write(kdspa)
						querytbl = "SELECT   TOP (1) kdspa, tglberlaku, kdafdeling, afdeling, kdptn, status, jml, jmlcetak, SisaJatah, Nama, kdbap, bap, nopetak, wil, kemasakan, KodeInsert, kdPos, Pos, 'Sukses' AS Pesan FROM vJatahSPTA_Mobile_print where kdspa='" & kdspa & "' order by kdptn"

						'http://115.85.64.67/json/jsonsdmlogin.asp?user_id=a&password=a'

						'response.write(querytbl)

						'Ambil data'
						set rd = server.CreateObject("ADODB.RECORDSET")
						rd.Open querytbl, conn,3,1
						i = 1

						while not rd.eof
							
							
							Register = rd.fields ("kdptn")
							Nama = rd.fields ("nama")
							Afdeling = rd.fields ("Afdeling")
							KodeInsert = rd.fields ("KodeInsert")
							bap = rd.fields ("bap")
							kemasakan = rd.fields ("kemasakan")



							i=i+1
							rd.movenext 
						wend

						'response.write("[" & left(jsonString,len(jsonString)-1) & "]")

	kdspa = replace(Request.QueryString("kdspa"),"_","#")
	KodeInsert = KodeInsert
	UserId = Request.QueryString("userid") 

	'insert ke Tabel tblspta_mobile
	 qinsert = "insert into tblspta_mobile(tglspta, Register, Nama, Afdeling, NoSPTA, KodeSticker, NoKendaraan, Pos, x, y, Lokasi, Foto1, Foto2, Keterangan, kdspa,userid,Petugas,KodeInsert) values ('" & Request.QueryString("tglspta") & "','" & Register  & "','" & Nama & "','" & Afdeling & "','" & Request.QueryString("nospta") & "','" & Request.QueryString("kodesticker") & "',UPPER('" & Request.QueryString("nokendaraan") & "'),'" & Request.QueryString("pos") & "','" & Request.QueryString("x") & "','" & Request.QueryString("y") & "','" & Request.QueryString("lokasi") & "','" & Request.QueryString("foto1") & "','" & Request.QueryString("foto2") & "','" & Request.QueryString("keterangan") & "',replace('" & Request.QueryString("kdspa") & "','_','#'),'" & Request.QueryString("userid") & "','" & Request.QueryString("Petugas") & "','" & KodeInsert & "')"

	
	set iD = server.CreateObject("ADODB.RECORDSET")
	iD.Open qinsert, Conn,3,1
	


	'update kuota SPTA tblKuotaNoInduk
	qinsert = "update tblKuotaNoInduk set status=1,TglTerakhirCetak=getdate(),jmlcetak=isnull(jmlcetak,0)+1,Cetak='Mobile' +'" & Request.QueryString("userid") & "' where kdspa='" & kdspa & "'"

	set iD = server.CreateObject("ADODB.RECORDSET")
	iD.Open qinsert, Conn,3,1

	'insert ke table tblspa
	qinsertspa = "insert into tblspa (kdspaKuota,kdspa,tglberlaku,kdafdeling,afdeling,drjam,sdjam,TglEntry,status,statusgwg,Kdptn,Kdwkt,nmuser,namaptn,kdBap,Versi,keterangan,KodeStiker,NoKend,Pos) select top 1 kdspa as kdspaKuota,'"&KodeInsert&"',tglberlaku,kdafdeling,afdeling ,convert(datetime,convert(varchar,tglberlaku,23)+' 06:00:00',120) as drjam,convert(datetime,convert(varchar,tglberlaku+1,23)+' 05:59:59',120) as sdjam,getdate() as TglEntry,1 as Status,0 as statusgwg,kdptn as kdptn,0 as Kdwkt,'"&UserId&"' as nmuser,nama as namaptn,kdBap,0 as Versi,'SPTAMobile','" & Request.QueryString("kodesticker") & "' as KodeStiker, '" & Request.QueryString("nokendaraan") & "' as NoKend,'" & Request.QueryString("pos") & "' as Pos  from vJatahSPTA_Mobile_print where kdspa='"&kdspa&"' order by TglEntry"

	set iDi = server.CreateObject("ADODB.RECORDSET")
	iDi.Open qinsertspa, Conn,3,1



   'hasil insert ke tabel spta'
   querytblspta = "select 'Sukses' as Pesan,left(Kdptn,2) as KdKUD,spa as NoSPTA,'PG. Krebet Baru Malang' as Pabrik,convert(varchar,getdate(),103) + ' '+convert(varchar,getdate(),108) as Tanggal,afdeling, '" & Request.QueryString("pos") & "' as Pos, '"&kemasakan&"' as Varitas,'"&bap&"' as BeritaAcara,'Surat Perintah Tebang Angkut' as Judul,Kdptn as NoInduk,namaptn as Pemilik,NoKend as NoKendaraan,'24JAM' as Berlaku,convert(varchar,drjam,103)+' '+convert(varchar,drjam,8) as Dari,convert(varchar,sdjam,103)+' '+convert(varchar,sdjam,8) as Sampai, '" & Request.QueryString("Petugas") & "' as petugas from tblSPA where ltrim(rtrim(kdspa))='"&KodeInsert&"' and kdspaKuota='"&kdspa&"'"
  ' response.write(querytblspta)

						'http://115.85.64.67/json/jsonsdmlogin.asp?user_id=a&password=a'

						'response.write(querytbl)

						'Ambil data'
						set rd = server.CreateObject("ADODB.RECORDSET")
						rd.Open querytblspta, conn,3,1
						i = 1

						jsonString = ""


						while not rd.eof
							
							jsonString = jsonString

							
							Register = rd.fields ("kdptn")
							Nama = rd.fields ("nama")
							Afdeling = rd.fields ("Afdeling")
							KodeInsert = rd.fields ("KodeInsert")



							recd = "{"
							For each item in rd.Fields
								fd = item.Name
								recd = recd & """" & item.Name & """" & " : " & """" & rd.fields (fd) & """," 
							Next


							jsonString = jsonString & recd
							jsonString = left(jsonString,len(jsonString)-1) & "},"


							i=i+1
							rd.movenext 
						wend

						response.write("[" & left(jsonString,len(jsonString)-1) & "]")



	


%>
