<!--#include file="connSmartFarming.inc" -->

<% 
 On Error Resume Next 
'membuat query'
 nopetak=request.querystring("nopetak")
 kecamatan=request.querystring("kecamatan")




						querytbl="SELECT top 1 a.nopetak, b.x, b.y FROM tblRegisterPetak a inner join vPolygonMaps_Mobile b on a.nopetak = b.nopetak and a.kecamatan = b.kecamatan where a.nopetak='"&nopetak&"' and a.kecamatan= '"&kecamatan&"'"



						'response.write(querytbl)
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