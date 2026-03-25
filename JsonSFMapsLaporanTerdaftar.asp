<!--#include file="connSmartFarming.inc" -->

<% 
 On Error Resume Next 
'membuat query'
 kolom=request.querystring("kolom")
 cari=request.querystring("cari")




						querytbl="SELECT top 1 a.mg, count(a.mg) as countterdaftar, sum(luas) as terdaftar, round(avg(brix),2) as rata, max(brix) as tertinggi, min(brix) as terendah FROM tblRegisterPetak a left JOIN vBrixTerakhir b on a.nopetak = b.nopetak and a.kecamatan=b.kecamatan and a.mg = b.mg group by a.mg order by a.mg desc"
					



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