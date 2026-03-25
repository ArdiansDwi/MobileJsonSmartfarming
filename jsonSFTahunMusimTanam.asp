<!--#include file="connSmartFarming.inc" -->

<% 
 On Error Resume Next 


  ' if (userid="ARDI") or (userid="M") or (userid="R") then userid = "CA001" end if

						querytbl =  "SELECT   tahun from vTahunMusimTanam order by tahun desc"

						

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