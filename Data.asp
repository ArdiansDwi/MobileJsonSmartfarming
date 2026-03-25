<%

tgl=0
'ST GILINGAN(((=====================================================================================================================
'pH NPP
if (rec.fields ("parameter") = "pH NPP") then
	tgl=1
	sQueryK =  "SELECT     a.Tanggal, a.Jam, ISNULL(SUM(b.PHNPP) / COUNT(b.PHNPP), 0) AS Nilai FROM vTanggalJamDisplay a LEFT OUTER JOIN web_StGilingan b ON CONVERT(varchar, a.Tanggal-" & tgl & ", 103) = CONVERT(varchar,b.tgl, 103) AND a.Jam = DATEPART(hour, b.tgl) AND ISNULL(b.PHNPP ,0) > 0 and   (b.pbrk = '" & Pabrik & "') GROUP BY a.Tanggal, a.Jam ORDER BY a.Tanggal"
	tgl=0
	sQuery = "SELECT     a.Tanggal, a.Jam, ISNULL(SUM(b.PHNPP) / COUNT(b.PHNPP), 0) AS Nilai FROM vTanggalJamDisplay a LEFT OUTER JOIN web_StGilingan b ON CONVERT(varchar, a.Tanggal-" & tgl & ", 103) = CONVERT(varchar,b.tgl, 103) AND a.Jam = DATEPART(hour, b.tgl) AND ISNULL(b.PHNPP ,0) > 0 and   (b.pbrk = '" & Pabrik & "') GROUP BY a.Tanggal, a.Jam ORDER BY a.Tanggal"

'Imbibisi % Tebu
elseif (rec.fields ("parameter") = "Imbibisi % Tebu") then
	tgl=1
	sQueryK =  "SELECT     a.Tanggal, a.Jam, ISNULL(SUM(b.Imbibisi_Ku) / SUM(b.Tebu_digiling), 0)*100 AS Nilai FROM vTanggalJamDisplay a LEFT OUTER JOIN web_StGilingan b ON CONVERT(varchar, a.Tanggal-" & tgl & ", 103) = CONVERT(varchar,b.tgl, 103) AND a.Jam = DATEPART(hour, b.tgl) AND ISNULL(b.Tebu_digiling ,0) > 0 and   (b.pbrk = '" & Pabrik & "') GROUP BY a.Tanggal, a.Jam ORDER BY a.Tanggal"
	tgl=0
	sQuery = "SELECT     a.Tanggal, a.Jam, ISNULL(SUM(b.Imbibisi_Ku) / SUM(b.Tebu_digiling), 0)*100 AS Nilai FROM vTanggalJamDisplay a LEFT OUTER JOIN web_StGilingan b ON CONVERT(varchar, a.Tanggal-" & tgl & ", 103) = CONVERT(varchar,b.tgl, 103) AND a.Jam = DATEPART(hour, b.tgl) AND ISNULL(b.Tebu_digiling ,0) > 0 and   (b.pbrk = '" & Pabrik & "') GROUP BY a.Tanggal, a.Jam ORDER BY a.Tanggal"
else

end if
%>