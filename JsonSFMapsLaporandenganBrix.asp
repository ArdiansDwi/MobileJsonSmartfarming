<!--#include file="connSmartFarming.inc" -->

<% 
 On Error Resume Next 
'membuat query'
 kolom=request.querystring("kolom")
 cari=request.querystring("cari")




						'querytbl="SELECT top 1 a.mg, count(a.mg) as countbrix, sum(luas) as luas FROM TblBrix a INNER JOIN tblRegisterPetak b on a.nopetak = b.nopetak and a.kecamatan=b.kecamatan and a.mg = b.mg group by a.mg order by a.mg desc"

						querytbl="select top 1 mg, count(mg) as countbrix, sum(luas) as luas from (SELECT a.mg, max(luas) as luas FROM TblBrix a INNER JOIN tblRegisterPetak b on a.nopetak = b.nopetak and a.kecamatan=b.kecamatan and a.mg = b.mg group by a.mg, a.nopetak, a.kecamatan) jml group by mg order by mg desc"
					



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