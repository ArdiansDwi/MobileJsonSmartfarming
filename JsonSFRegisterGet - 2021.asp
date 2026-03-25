<!--#include file="connTebu.inc" -->


<% 
	
 On Error Resume Next 


						'membuat query'

						'querytbl = "SELECT TOP (1) kdspa, tglberlaku, kdafdeling, afdeling, kdptn, status, jml, jmlcetak, SisaJatah, Nama, bap, nopetak, wil, kemasakan FROM vJatahSPTA_Mobile_register where kdspa=replace('" & Request.QueryString("kdspa") & "','_','#')"

						querytbl = "SELECT TOP (1) kdspa, SisaJatah FROM vJatahSPTA_Mobile_register where kdspa=replace('" & Request.QueryString("kdspa") & "','_','#')"
						'response.write(Request.QueryString("kdspa"))

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
