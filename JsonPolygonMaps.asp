<!--#include file="connSmartFarming.inc" -->


<% 
	
 On Error Resume Next 
'membuat query'
						x=Request.QueryString("x")
						y=Request.QueryString("y")

 						
						querytbl = "SELECT [inc],[id]+';'+[Fill]+';'+[StrokeColor] as id,[koordinat],[Fill],[StrokeColor], [x], [y], [nopetak]  FROM [SmartFarming].[dbo].[vPolygonMaps_Mobile] where x between "&x&"-0.000995 and "&x&"+0.000995 and y between "&y&"-0.000995 and "&y&"+0.000995"

						'querytbl = "SELECT TOP (50) [inc],[id],[koordinat],[Fill],[StrokeColor]  FROM [SmartFarming].[dbo].[vPolygonMaps] "


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
