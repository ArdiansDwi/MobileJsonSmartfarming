<!--#include file="connSmartFarming.inc" -->
<%

' ❗ MATIKAN dulu saat debug
 On Error Resume Next 

nopetak   = Replace(Request.QueryString("nopetak"), "'", "''")
kecamatan = Replace(Request.QueryString("kecamatan"), "'", "''")					

querytbl = "SELECT TOP 1 " & _
" a.nopetak, " & _
" ISNULL(a.petugas,'PG KBB') AS petugas, " & _
" CAST(a.luas AS VARCHAR(50)) AS luas, " & _
" a.desa, " & _
" a.kecamatan, " & _
" isnull(b.register,'') AS register, " & _
" case when b.register is not null then '1' else '1' end AS validasi, " & _
" datepart(year,getdate()) AS mg " & _
"FROM [dbo].[vPolygonMaps] a " & _
"LEFT OUTER JOIN vregisterpetak b " & _
" ON b.nopetak = a.nopetak " & _
" AND b.kecamatan = a.kecamatan " & _
"WHERE LTRIM(RTRIM(a.nopetak)) = '" & nopetak & "' " & _
" AND LTRIM(RTRIM(a.kecamatan)) = '" & kecamatan & "' " & _
"ORDER BY a.nopetak DESC"

' 🔍 DEBUG (pakai kalau error)
' Response.Write querytbl
' Response.End

set rd = server.CreateObject("ADODB.RECORDSET")
rd.Open querytbl, conn, 3, 1

jsonString = ""

If Not rd.EOF Then
    Do While Not rd.EOF

        recd = "{"

        For Each item In rd.Fields
            fd = item.Name
            value = rd.Fields(fd)

            ' ✅ HANDLE NULL
            If IsNull(value) Then value = ""

            ' ✅ HANDLE karakter JSON (quote)
            value = Replace(value, """", "\""") 

            recd = recd & """" & fd & """:""" & value & ""","
        Next

        ' buang koma terakhir
        recd = Left(recd, Len(recd)-1) & "}"

        jsonString = jsonString & recd & ","

        rd.MoveNext
    Loop

    ' buang koma terakhir array
    jsonString = "[" & Left(jsonString, Len(jsonString)-1) & "]"

Else
    jsonString = "[]"
End If

Response.Write jsonString

%>