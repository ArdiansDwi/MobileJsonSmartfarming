<!--#include file="connTebu.inc" -->


<% 
	
 On Error Resume Next 




						
						'membuat query'
						if Request.QueryString("user_id")="ARDI"  then
						querytbl = "SELECT KD_KT as kdkel, KD_KT+ ' : ' + REPLACE(KETUA,'''','`') as kelompok FROM [M_KEL] where ltrim(rtrim(ketua))<>'' order by KD_KT"
						'querytbl = "select kdkel,kelompok as kelompok from vJatahSPTApos_kelompok where pos is not null and sisa>0 group by kdkel,kelompok order by kdkel"

						else
						querytbl = "SELECT KD_KT as kdkel, KD_KT+ ' : ' + REPLACE(KETUA,'''','`') as kelompok FROM [M_KEL] where ltrim(rtrim(ketua))<>'' order by KD_KT"
						'elseif Request.QueryString("pos")<>""  then
						'querytbl = "select kdkel, kelompok from vJatahSPTApos_kelompok where upper(pos) in (select upper(pos) from vpostebu where wilayah=(select Wilayah from vPosTebu where upper(Pos) = upper('"&Request.QueryString("pos")&"'))) and sisa>0 group by kdkel,kelompok order by kdkel"
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
