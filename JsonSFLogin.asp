<!--#include file="connSmartFarming.inc" -->


<% 
	
 On Error Resume Next 



						qinsert = "insert into tblLog (tanggal,user_id,aktifitas) values (getdate(),'" & Request.QueryString("user_id") & "' , 'Login Aplikasi')"
 						set iD = server.CreateObject("ADODB.RECORDSET")
						iD.Open qinsert, Conn,3,1




						querytbl = "SELECT a.id, a.created_at, a.updated_at, a.user_id, a.Nama_user, a.Jabatan, a.Afdeling_id, a.Rayon_id, a.UnitPG_id, a.AP_id, a.tahun_aktif, a.nomaster, a.password, a.skw_id, a.skk_id, a.ca_id, a.device_id, a.aktif, ISNULL(b.user_id, 'ERRORR')  AS rfid,(select versi from tblversi where kdid=1) as versi FROM dbo.users_id AS a LEFT OUTER JOIN dbo.tblSPTAtanpaRfId AS b ON a.user_id = b.user_id where a.user_id='" & Request.QueryString("user_id") & "' and a.password='" & Request.QueryString("password") & "' and a.aktif=1"


						'response.write(querytbl)

						'Ambil data'
						set rd = server.CreateObject("ADODB.RECORDSET")
						rd.Open querytbl, conn,3,1
						i = 1

						jsonString = ""

						if not rd.eof then
							session("iduser") = rd.fields("user_id")
							session("txtnama") = rd.fields("Nama_user")
						end if

						if rd.eof then
							qinsert = "insert into tblLog (tanggal,user_id,aktifitas) values (getdate(),'" & Request.QueryString("user_id") & "' , 'Gagal Login Aplikasi')"
							set iD = server.CreateObject("ADODB.RECORDSET")
							iD.Open qinsert, Conn,3,1
						else 
							qinsert = "insert into tblLog (tanggal,user_id,aktifitas) values (getdate(),'" & Request.QueryString("user_id") & "' , 'Berhasil Login Aplikasi')"
							set iD = server.CreateObject("ADODB.RECORDSET")
							iD.Open qinsert, Conn,3,1
						end if

						while not rd.eof
							
							jsonString = jsonString

						


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
