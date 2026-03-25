<!--#include file="connSmartFarming.inc" -->


<% 
	
 On Error Resume Next 


						'membuat query'
						user = request.querystring("user_id")
            			if (user="ARDI") or (user="M") or (user="R") or (user="KS007") or (user="KS005") or (user="KS006") then user = "CA001" end if

						querytbl = "SELECT  id,urut,menu,intent,url,image,[level] FROM tblmenu where level like '%'+left('"&user&"',2)+'%' and isnull(image,'') <> '' and groupmenu = 'utama' order by urut"

'querytbl = "SELECT id,urut,menu,intent,url,image,[level] FROM tblmenu where [level]= '" & Request.QueryString("level") & "' order by urut"

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
