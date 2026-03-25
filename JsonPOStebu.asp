<!--#include file="connTebu.inc" -->


<% 
	
 On Error Resume Next 


						'membuat query'
						'querytbl = "select top 1  cast(abs(koordinatx)-abs(round(cast('"&Request.QueryString("x")&"' as float),7) as float) as Toleransix,kd, Pos, x, y, Keterangan, Koordinat, koordinatX, KoordinatY, RangeM from vpostebu where (abs(koordinatx)-abs(cast('"&Request.QueryString("x")&"' as float)) < RangeM) and (abs(koordinatx)-abs(cast('"&Request.QueryString("y")&"' as float)) < RangeM ) order by abs(abs(koordinatx)-abs(cast('"&Request.QueryString("x")&"' as float))),abs(abs(koordinaty)-abs(cast('"&Request.QueryString("y")&"' as float)))"

						x=Request.QueryString("x")
						y=Request.QueryString("y")

			

						querytbl = "select dbo.postebu('"&Request.QueryString("x")&"','"&Request.QueryString("y")&"') as Pos"


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
