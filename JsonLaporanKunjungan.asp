<!--#include file="connSmartFarming.inc" -->

<% 
 On Error Resume Next 
'membuat query'
 userid=request.querystring("user_id")
 tgl=left(request.querystring("tgldb"),2)
 bln=mid(request.querystring("tgldb"),3,2)
 thn=right(request.querystring("tgldb"),4)
 tgldb=thn+"-"+bln+"-"+tgl+" 00:00:00"

 tgl2=left(request.querystring("tgldb2"),2)
 bln2=mid(request.querystring("tgldb2"),3,2)
 thn2=right(request.querystring("tgldb2"),4)
 tgldb2=thn2+"-"+bln2+"-"+tgl2+" 23:59:59"
  if (userid="ARDI") or (userid="M") or (userid="R") or (userid="KS007") or (userid="KS005") or (userid="KS006") then userid = "CA001" end if




						'querytbl =  "SELECT top 1 kdid, CONVERT(VARCHAR, tgl, 105) as tgl, nomaster, iddevice, koordinatx, koordinaty, kodePetani, namaPetani, Alamat, judul, deskripsi, lokasifile1, lokasifile2, lokasifile3, lokasifile4, nama_user, convert(varchar, tgl, 8) as waktu FROM vkunjungankebun where plpg_id+skk_id+skw_id+ca_id like'%"&userid&"%' and tgl>=convert(datetime, '"&tgldb&"', 103) and tgl<=convert(datetime, '"&tgldb2&"', 103) order by tgl"

						'SELECT top 3 a.kdid,  b.nomaster, a.iddevice, a.koordinatx, a.koordinaty, a.kodePetani,  a.namaPetani, a.Alamat, a.judul, a.deskripsi, a.lokasifile1,  a.lokasifile2,  a.lokasifile3,  a.lokasifile4, b.nama_user, a.tgl as waktu FROM tblkunjungan a INNER JOIN dbo.users_id AS b ON a.iddevice = b.user_id  where iddevice ='"&userid&"' order by tgl desc '

						if (UCase(left(userid,2))="PL") then
						querytbl =  "SELECT a.kdid,  b.nomaster, a.iddevice, a.koordinatx, a.koordinaty, a.kodePetani,  a.namaPetani, a.Alamat, a.judul, a.deskripsi, a.lokasifile1,  a.lokasifile2,  a.lokasifile3,  a.lokasifile4, b.nama_user, a.tgl as waktu FROM tblkunjungan a INNER JOIN dbo.users_id AS b ON a.iddevice = b.user_id  where iddevice ='"&userid&"' and tgl>='"&tgldb&"' and tgl<='"&tgldb2&"' order by tgl desc"
						else 
						querytbl =  "SELECT kdid,tgl,  b.nomaster, iddevice, koordinatx, koordinaty, kodePetani,  namaPetani, Alamat, judul, deskripsi, lokasifile1,  lokasifile2,  lokasifile3,  lokasifile4, nama_user, convert(varchar,tgl,103)+' '+convert(varchar,tgl,8) as waktu FROM tblkunjungan a INNER JOIN dbo.users_id AS b ON a.iddevice = b.user_id  where iddevice+skk_id+skw_id+ca_id like '%"&userid&"%' and tgl>='"&tgldb&"' and tgl<='"&tgldb2&"' order by kdid desc "
						end if


						'querytbl =  "SELECT top 1 kdid,tgl,  b.nomaster, iddevice, koordinatx, koordinaty, kodePetani,  namaPetani, Alamat, judul, deskripsi, lokasifile1,  lokasifile2,  lokasifile3,  lokasifile4, nama_user, convert(varchar,tgl,103)+' '+convert(varchar,tgl,8) as waktu FROM tblkunjungan a INNER JOIN dbo.users_id AS b ON a.iddevice = b.user_id  where iddevice+skk_id+skw_id+ca_id like '%"&userid&"%' and tgl>='"&tgldb&"' and tgl<='"&tgldb2&"' order by tgl "


						'querytbl = "select 1 as waktu"

						'response.write(querytbl)


						set rd = server.CreateObject("ADODB.RECORDSET")
						rd.Open querytbl, conn,3,1
						i = 1

						jsonString = ""

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