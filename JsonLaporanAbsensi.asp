<!--#include file="connSmartFarming.inc" -->


<% 
	
 On Error Resume Next 
'membuat query'
						 userid=request.querystring("user_id")

 						
 						if (userid = "ARDI") then
						querytbl = "SELECT kdid, tgl, nomaster, iddevice, koordinatx, koordinaty, lokasifile, convert(varchar,tgl,103) + ' ('+nama_user+')' + ' : ' + lokasi as lokasi, nama_user, convert(varchar, tgl, 8) as Jam FROM vLaporanAbsensi  where CONVERT(VARCHAR, tgl, 101)>=CONVERT(VARCHAR, getdate()-3, 101)  order by tgl desc"
						else
						querytbl = "SELECT kdid, tgl, nomaster, iddevice, koordinatx, koordinaty, lokasifile, convert(varchar,tgl,103) + ' : ' + lokasi as lokasi, nama_user,convert(varchar, tgl, 8) as Jam FROM vLaporanAbsensi where CONVERT(VARCHAR, tgl, 101)>=CONVERT(VARCHAR, getdate()-3, 101) and iddevice='"&userid&"' order by tgl desc"

						end if

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
