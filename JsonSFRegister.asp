<!--#include file="connTebu.inc" -->


<% 
	
 On Error Resume Next 
'membuat query'
 					if Request.QueryString("kdkel")="Pilih Kelompok"  then 	
						querytbl = "select  'PilihRegister' as kdspa, 'Pilih Register' as afdeling, 'Pilih Register' as Nama, 'Pilih Register' as  kdptn, '0' sisajatah, 'Pilih Register'  as register"

					elseif Request.QueryString("kdkel")<>""  then

						if Request.QueryString("user_id")="ARDI"  then
						'querytbl = "select  kdspa,afdeling, Nama,left(Nama,12)+' : '+kdptn+' - '+cast(sisajatah as varchar) as  kdptn, sisajatah,kdptn as register from vJatahSPTA_Mobile_register where left(kdptn,5) = '"&Request.QueryString("kdkel")&"' order by nama,kdptn,tglberlaku desc"
						
						querytbl = "select a.register as kdspa,a.afdeling,a.Nama,left(a.Nama,12)+' : '+register as kdptn,0 as SisaJatah,a.register,a.pos from vRegisterPos a where left(register,5) = '"&Request.QueryString("kdkel")&"' order by nama, register desc"

						elseif Request.QueryString("pos")<>""  then
						'querytbl = "select  kdspa,afdeling, Nama,left(Nama,12)+' : '+kdptn+' - '+cast(sisajatah as varchar)  as  kdptn, sisajatah,kdptn as register from vJatahSPTA_Mobile_register where upper(pos) in (select upper(pos) from vpostebu where wilayah=(select Wilayah from vPosTebu where upper(Pos) = upper('"&Request.QueryString("pos")&"'))) and left(kdptn,5) = '"&Request.QueryString("kdkel")&"' order by nama, kdptn,tglberlaku desc"

						querytbl = "select a.register as kdspa,a.afdeling,a.Nama,left(a.Nama,12)+' : '+register as kdptn,0 as SisaJatah,a.register,a.pos from vRegisterPos a where  upper(a.pos) in (select upper(pos) from vpostebu where wilayah=(select Wilayah from vPosTebu where upper(Pos) = upper('"&Request.QueryString("pos")&"'))) and left(register,5) = '"&Request.QueryString("kdkel")&"' order by nama, register desc"
						else
						querytbl = "select  'PilihRegister' as kdspa, 'Pilih Register' as afdeling, 'Pilih Register' as Nama, 'Pilih Register' as  kdptn, '0' sisajatah, 'Pilih Register'  as register"
						end if 

					' else

					' 	if Request.QueryString("user_id")="ARDI"  then
					' 	querytbl = "select  kdspa,afdeling, Nama,left(Nama,12)+' : '+kdptn+' - '+cast(sisajatah as varchar) as  kdptn, sisajatah,kdptn as register from vJatahSPTA_Mobile_register order by nama,kdptn,tglberlaku desc"

					' 	elseif Request.QueryString("pos")<>""  then
					' 	querytbl = "select  kdspa,afdeling, Nama,left(Nama,12)+' : '+kdptn+' - '+cast(sisajatah as varchar)  as  kdptn, sisajatah,kdptn as register from vJatahSPTA_Mobile_register where upper(pos) in (select upper(pos) from vpostebu where wilayah=(select Wilayah from vPosTebu where upper(Pos) = upper('"&Request.QueryString("pos")&"'))) order by nama, kdptn,tglberlaku desc"
					' 	else
					' 	querytbl = "select kdspa,afdeling, Nama,kdptn+' : '+Nama+' - '+afdeling as  kdptn, sisajatah,kdptn as register from vJatahSPTA_Mobile_register where kdspa<>'210530071639071#192.168.0.73' and kdptn='EQ010005J40' order by tglberlaku desc, kdptn"
					' 	end if 
					end if

						'http://115.85.64.67/json/jsonsdmlogin.asp?user_id=a&password=a'

						'response.write(querytbl)

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
