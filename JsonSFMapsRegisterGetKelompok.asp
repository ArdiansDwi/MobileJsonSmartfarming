<!--#include file="connTebu.inc" -->


<% 
	
 On Error Resume Next 


						'membuat query'


						'querytbl = "SELECT TOP (1) kdspa, SisaJatah FROM vJatahSPTA_Mobile_register where kdspa=replace('" & Request.QueryString("kdspa") & "','_','#')"

						' querytbl = "select top 1 register as kdspa,isnull(b.sisa,0) as SisaJatah from vRegisterPos a inner join vjatahsptapos_kelompok b on left(a.register,5)=b.kdkel and a.pos=b.pos where upper(a.pos) in (select upper(pos) from vpostebu where wilayah=(select Wilayah from vPosTebu where upper(Pos) = upper('"&Request.QueryString("pos")&"'))) and a.register = '"&Request.QueryString("kdspa")&"')"

						' if Request.QueryString("user_id")="ARDI"  then 
						' 	querytbl = "select top 1 kdkel as kdspa,sisa as SisaJatah from vjatahsptapos_kelompok  where kdkel=left('"&Request.QueryString("kdspa")&"',5) order by sisa desc"
						' else
						' querytbl = "select top 1 kdkel as kdspa,sisa as SisaJatah from vjatahsptapos_kelompok  where upper(pos) in (select upper(pos) from vpostebu where wilayah=(select Wilayah from vPosTebu where upper(Pos) = upper('"&Request.QueryString("pos")&"'))) and kdkel=left('"&Request.QueryString("kdspa")&"',5) order by sisa desc"
						' end if


						querytbl = "select top 1 register as kdspa,isnull(b.sisa,0) as SisaJatah from vRegisterPos a inner join vjatahsptapos_kelompok_sisa b on left(a.register,5)=b.kdkel and a.pos=b.pos where a.register = '"&Request.QueryString("kdspa")&"'"

						'querytbl = "select a.kelompok as kdkel,case when max(b.tglgiling)=max(b.tglaktif) then isnull(max(b.jatah)+max(b.tambahan)-max(b.terpakai),0)  else 0 end as sisa from vkelompokpos a left outer join vjatahSPTApos_kelompok_source b on a.kelompok=b.kd_kt where a.pos is not null and a.kelompok='"&Request.QueryString("kdspa")&"' group by a.kelompok,a.ketua,a.tglgiling,a.pos"



						'Ambil data'
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
