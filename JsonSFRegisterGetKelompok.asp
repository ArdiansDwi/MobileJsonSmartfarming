<!--#include file="connTebu.inc" -->
<%
On Error Resume Next

Dim register, tanggal, sql, rs
Dim afdeling, jeniskendaraan, kategori
Dim tglaktif, tgllibur, bebasspta
Dim jatahutama, jatahtambahan, jatahterpakai, jatahpr, sisajatah

register = Request("kdspa")

'========================
' STEP 0: TANGGAL LOGIC
'========================
Dim dtNow
dtNow = Now()

If Hour(dtNow) < 6 Then
    dtNow = DateAdd("d", -1, dtNow)
End If

tanggal = Year(dtNow) & "-" & Right("0" & Month(dtNow),2) & "-" & Right("0" & Day(dtNow),2)

'========================
' STEP 1: DATA UTAMA
'========================
sql = "SELECT TOP 1 " & _
      "a.register, a.afdeling, a.jeniskendaraan, a.kategori, " & _
      "ISNULL(CONVERT(VARCHAR, c.tgl, 120), '-') AS tglaktif, " & _
      "ISNULL(CONVERT(VARCHAR, d.tgl, 120), '-') AS tgllibur, " & _
      "ISNULL((SELECT bebas FROM tblSPTAbebas), 0) AS bebasspta " & _
      "FROM vRegisterPos a " & _
      "LEFT JOIN tblBatasTglSKW c ON CONVERT(DATE, c.tgl, 120) = '" & tanggal & "' " & _
      "LEFT JOIN tblKalenderSPTA d ON (d.wil = a.afdeling OR d.wil = 'SEMUA') " & _
      "AND CONVERT(DATE, d.tgl, 120) = '" & tanggal & "' " & _
      "WHERE a.register = '" & register & "'"

Set rs = conn.Execute(sql)

If Not rs.EOF Then
    afdeling = rs("afdeling")
    jeniskendaraan = rs("jeniskendaraan")
    kategori = rs("kategori")
    tglaktif = rs("tglaktif")
    tgllibur = rs("tgllibur")
    bebasspta = CInt(rs("bebasspta"))
End If
rs.Close

'========================
' STEP 2: JATAH UTAMA
'========================
jatahutama = 0

sql = "SELECT TOP 1 CASE " & _
      "WHEN '" & jeniskendaraan & "'='GANDENG' THEN gd " & _
      "WHEN '" & jeniskendaraan & "'='FUSO' THEN fs " & _
      "ELSE cd END AS jatahutama " & _
      "FROM TblJatahSPTA WHERE kd_kt = LEFT('" & register & "',5)"

Set rs = conn.Execute(sql)
If Not rs.EOF Then jatahutama = CInt(rs("jatahutama"))
rs.Close

'========================
' STEP 3: JATAH TAMBAHAN
'========================
jatahtambahan = 0

sql = "SELECT ISNULL(SUM(CASE " & _
      "WHEN '" & jeniskendaraan & "'='GANDENG' THEN gd " & _
      "WHEN '" & jeniskendaraan & "'='FUSO' THEN fs " & _
      "ELSE cd END),0) AS jatahtambahan " & _
      "FROM TblJatahSPTA_tambahan " & _
      "WHERE kd_kt = LEFT('" & register & "',5) " & _
      "AND CONVERT(DATE, tanggal,120)='" & tanggal & "'"

Set rs = conn.Execute(sql)
If Not rs.EOF Then jatahtambahan = CInt(rs("jatahtambahan"))
rs.Close

'========================
' STEP 4: TERPAKAI
'========================
jatahterpakai = 0

sql = "SELECT ISNULL(SUM(jml),0) AS jatahterpakai " & _
      "FROM tblKuotaNoInduk " & _
      "WHERE LEFT(kdptn,5)=LEFT('" & register & "',5) " & _
      "AND RIGHT(kdptn,1)=RIGHT('" & register & "',1) " & _
      "AND CONVERT(DATE,tglberlaku,120)='" & tanggal & "'"

Set rs = conn.Execute(sql)
If Not rs.EOF Then jatahterpakai = CInt(rs("jatahterpakai"))
rs.Close

'========================
' STEP 5: JATAH PR
'========================
jatahpr = 0

sql = "SELECT (" & _
      "ISNULL((SELECT SUM(jml) FROM TblBeritaAcaraTebuPR WHERE kd_ptn='" & register & "'),0) + " & _
      "ISNULL((SELECT TOP 1 Bebas FROM tblSPTAbebasPR WHERE kategori='" & kategori & "'),0) - " & _
      "ISNULL((SELECT SUM(jml) FROM tblKuotaNoInduk WHERE kdptn='" & register & "'),0) " & _
      ") AS JatahPR"

Set rs = conn.Execute(sql)
If Not rs.EOF Then jatahpr = CInt(rs("JatahPR"))
rs.Close

'========================
' STEP 6: HITUNG SISA
'========================
sisajatah = 0

If bebasspta > 0 Then
    sisajatah = 100
ElseIf tglaktif <> "-" AND tgllibur = "-" Then
    sisajatah = jatahutama + jatahtambahan - jatahterpakai
ElseIf tglaktif <> "-" AND tgllibur <> "-" Then
    sisajatah = jatahtambahan - jatahterpakai
End If

If (kategori = "PR" OR kategori = "MA") AND jatahpr < sisajatah Then
    sisajatah = jatahpr
End If

If sisajatah < 0 Then sisajatah = 0

'========================
' STEP 7: OUTPUT JSON
'========================
Response.ContentType = "application/json"

Dim json
json = "[{""kdspa"":""" & register & """,""SisaJatah"":" & sisajatah & "}]"

Response.Write json
%>