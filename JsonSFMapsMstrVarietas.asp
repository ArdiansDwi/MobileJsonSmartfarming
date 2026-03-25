<!--#include file="connTebu.inc" -->


<% 
	
 On Error Resume Next 




						
						'membuat query'
						if Request.QueryString("user_id")="ARDI"  then
						querytbl = "SELECT  inc, varietas, jenis, kategori FROM tblVarietas order by inc"

						else
						querytbl = "SELECT  inc, varietas, jenis, kategori FROM tblVarietas order by inc"
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
