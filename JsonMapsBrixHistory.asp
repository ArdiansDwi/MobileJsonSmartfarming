<!--#include file="connSmartFarming.inc" -->


<% 
	
 On Error Resume Next 
'membuat query'
nopetak=request.querystring("nopetak")
mg=request.querystring("mg")
kecamatan=request.querystring("kecamatan")

 						
						querytbl = "SELECT kdid,tglentry,user_id,nopetak,register,koordinatx,koordinaty,lokasifile,lokasi, isnull(brix,0) as brix, kecamatan, mg, convert(varchar, tglentry, 103) +' '+ convert(varchar, tglentry, 108) as tgltampil FROM TblBrix where nopetak='"&nopetak&"' and mg='"&mg&"' and kecamatan='"&kecamatan&"' order by tglentry desc"


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
