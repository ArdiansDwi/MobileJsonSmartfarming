<!--#include file="connSmartFarming.inc" -->


<% 
	
 On Error Resume Next 

	'membuat query'                   
	'                          
	querytbl = "Select tglserver, tglformat1, tglformat2, tglformat3, tglformat4, tglformat5 from vtglserver"
	
	'Ambil data'
	set rd = server.CreateObject("ADODB.RECORDSET")
	rd.Open querytbl, conn,3,1
	'response.write("sukses")'	
	
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
	

	response.write ("query result : " & querytbl)
%>
