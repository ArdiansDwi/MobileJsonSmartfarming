<!--#include file="connTebu.inc" -->


<% 
	
 On Error Resume Next 


						'membuat query'
						querytbl = "select top 1 isnull(nokend,'') as nokend, kodestiker, tglentry from tblspa where kodestiker='" & Request.QueryString("kodestiker") & "' ORDER BY inc desc"


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
