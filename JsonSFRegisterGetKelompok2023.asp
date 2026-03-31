<!--#include file="connTebu.inc" -->


<% 
	
 On Error Resume Next 



						querytbl = "select top 1 register as kdspa,isnull(b.sisa,0) as SisaJatah from vRegisterPos a inner join vjatahsptapos_kelompok_sisa b on left(a.register,5)=b.kdkel and a.pos=b.pos where a.register = '"&Request.QueryString("kdspa")&"'"


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
