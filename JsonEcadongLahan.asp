<!--#include file="conncadong.inc" -->

<% 
 On Error Resume Next 
'membuat query'
 userid=request.querystring("user_id")
 tahun=request.querystring("tahun")
 tgl=left(request.querystring("tgldb"),2)
 bln=mid(request.querystring("tgldb"),3,2)
 thn=right(request.querystring("tgldb"),4)
 tgldb=thn+"-"+bln+"-"+tgl+" 00:00:00"

 tgl2=left(request.querystring("tgldb2"),2)
 bln2=mid(request.querystring("tgldb2"),3,2)
 thn2=right(request.querystring("tgldb2"),4)
 tgldb2=thn2+"-"+bln2+"-"+tgl2+" 23:59:59"
  if (userid="ARDI") or (userid="M") or (userid="R") or (userid="KS007") or (userid="KS005") or (userid="KS006") then userid = "CA001" end if

''if tahun="" then tahun="2000" end if

'tahun="2022"



						'querytbl =  "SELECT     a.NPS+cast(a.mt as varchar) as Kode,a.NPS, c.Register, isnull(max(e.PLPG_RAK),b.PLPG) as PLPG, c.Desa, c.Kec, c.Ktgr, isnull(max(e.MasaTanam),'') as MasaTanam, b.LuasUkur as Luas, isnull(a.Taksasi,0) as Taksasi, isnull(a.Nilai_Sewa,0) AS IPL, isnull(max(e.BiayaRAK),0) as RAK, SUM(isnull(d.biaya,0)) AS Realisasi,case when isnull(max(e.BiayaRAK),0)>0 and  SUM(isnull(d.biaya,0))>0 then (SUM(isnull(d.biaya,0))/isnull(max(e.BiayaRAK),0))*100 else 0 end as Persen, a.MT,b.nama,b.LokasiKebun,max(e.varietas) as Varietas,a.status,max(e.SKW_RAK) as SKW,max(e.SKK_RAK) as SKK,max(TotalJuring) as TotalJuring,max(GotMalang) as GotMalang,max(GotMujur) as GotMujur,max(GotKeliling) as GotKeliling, a.MT+1 as MG, 0 as persen FROM  tbDetailSewa a left outer JOIN tbPengajuanLahan b ON a.NPS = b.NPS LEFT OUTER JOIN tbACC c ON a.NPS = c.NPS LEFT OUTER JOIN vGarap_Rincian d ON a.NPS = d.NPS AND a.MT = d.MT LEFT OUTER JOIN vRAK_terakhir e on a.NPS=e.NPS and a.MT=e.MT where b.ket='ACC' and a.mt=2022 GROUP BY a.NPS, c.Register, b.PLPG, c.Desa, c.Kec, c.Ktgr, b.LuasUkur, a.Taksasi, a.Nilai_Sewa, a.RAK, a.MT,b.nama,b.LokasiKebun,a.status"

						querytbl="SELECT     a.NPS+cast(a.mt as varchar) as Kode,a.NPS, c.Register, isnull(max(e.PLPG_RAK),b.PLPG) as PLPG, c.Desa, c.Kec, c.Ktgr, isnull(max(e.MasaTanam),'') as MasaTanam, b.LuasUkur as Luas, isnull(a.Taksasi,0) as Taksasi, isnull(a.Nilai_Sewa,0) AS IPL,isnull((select  SUM(CASE WHEN kategori IN ('C. Bon Dalam', 'D. Bon Luar') THEN biaya ELSE 0 END) AS BiayaRAK from tblRAK_Rincian where NPS=a.nps and MT=a.MT and TglEntry=max(e.tglentry)),0) as RAK , SUM(isnull(d.biaya,0)) AS Realisasi ,case when SUM(isnull(d.biaya,0)) >0 then round(SUM(isnull(d.biaya,0))/(select  SUM(CASE WHEN kategori IN ('C. Bon Dalam', 'D. Bon Luar') THEN biaya ELSE 0 END) AS BiayaRAK from tblRAK_Rincian where NPS=a.nps and MT=a.MT and TglEntry=max(e.tglentry))*100,0) else 0 end  as Persen, a.MT ,b.nama,b.LokasiKebun,isnull(max(e.varietas),'') as Varietas,a.status,isnull(max(e.SKW_RAK),'') as SKW,isnull(max(e.SKK_RAK),'') as SKK ,isnull(max(TotalJuring),0) as TotalJuring,isnull(max(GotMalang),0) as GotMalang,isnull(max(GotMujur),0) as GotMujur ,isnull(max(GotKeliling),0) as GotKeliling , a.MT+1 as MG FROM tbDetailSewa a left outer JOIN tbPengajuanLahan b ON a.NPS = b.NPS LEFT OUTER JOIN tbACC c ON a.NPS = c.NPS  LEFT OUTER JOIN vGarap_Rincian d ON a.NPS = d.NPS AND a.MT = d.MT LEFT OUTER JOIN vRAK_valid e on a.NPS=e.NPS and a.MT=e.MT where b.ket='ACC' and a.mt='"&tahun&"' GROUP BY a.NPS, c.Register, b.PLPG, c.Desa, c.Kec, c.Ktgr, b.LuasUkur, a.Taksasi, a.Nilai_Sewa, a.RAK, a.MT ,b.nama,b.LokasiKebun,a.status order by a.nps"

						'querytbl="SELECT     a.NPS+cast(a.mt as varchar) as Kode,a.NPS, c.Register, isnull(max(e.PLPG_RAK),b.PLPG) as PLPG, c.Desa, c.Kec, c.Ktgr, isnull(max(e.MasaTanam),'') as MasaTanam, b.LuasUkur as Luas, isnull(a.Taksasi,0) as Taksasi, isnull(a.Nilai_Sewa,0) AS IPL,(select  SUM(CASE WHEN kategori IN ('C. Bon Dalam', 'D. Bon Luar') THEN biaya ELSE 0 END) AS BiayaRAK from tblRAK_Rincian where NPS=a.nps and MT=a.MT and TglEntry=max(e.tglentry)) as RAK , SUM(isnull(d.biaya,0)) AS Realisasi ,case when SUM(isnull(d.biaya,0)) >0 then round(SUM(isnull(d.biaya,0))/(select  SUM(CASE WHEN kategori IN ('C. Bon Dalam', 'D. Bon Luar') THEN biaya ELSE 0 END) AS BiayaRAK from tblRAK_Rincian where NPS=a.nps and MT=a.MT and TglEntry=max(e.tglentry))*100,0) else 0 end  as Persen, a.MT ,b.nama,b.LokasiKebun,max(e.varietas) as Varietas,a.status,max(e.SKW_RAK) as SKW,max(e.SKK_RAK) as SKK ,max(TotalJuring) as TotalJuring,max(GotMalang) as GotMalang,max(GotMujur) as GotMujur ,max(GotKeliling) as GotKeliling , a.MT+1 as MG FROM tbDetailSewa a left outer JOIN tbPengajuanLahan b ON a.NPS = b.NPS LEFT OUTER JOIN tbACC c ON a.NPS = c.NPS  LEFT OUTER JOIN vGarap_Rincian d ON a.NPS = d.NPS AND a.MT = d.MT LEFT OUTER JOIN vRAK_valid e on a.NPS=e.NPS and a.MT=e.MT where b.ket='ACC' and a.mt=2022 GROUP BY a.NPS, c.Register, b.PLPG, c.Desa, c.Kec, c.Ktgr, b.LuasUkur, a.Taksasi, a.Nilai_Sewa, a.RAK, a.MT ,b.nama,b.LokasiKebun,a.status order by a.nps"



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