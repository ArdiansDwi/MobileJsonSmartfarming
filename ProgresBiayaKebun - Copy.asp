<html class="no-js">

<head>
	<meta charset='UTF-8'>
	 <meta http-equiv="refresh" content="300" >
	
	<title>Data</title>
	
<style>

</style>

<script src="jquery.min.js"></script>
<script src="modernizr.js"></script>
<script>
	//paste this code under head tag or in a seperate js file.
	// Wait for window load
	$(window).load(function() {
		// Animate loader off screen
		$(".se-pre-con").fadeOut("slow");;
	});
</script>
<link href="data.css" rel="stylesheet" type="text/css" />

</head>

<!--#include file="conn.inc" -->



<body background="Images/body.png">
	<p align="left">
	<!-- Paste this code after body tag -->
	<div class="se-pre-con"></div>
	<!-- Ends -->
	
	<%
if  request.Form("tgl") <> "" and  request.Form("bln") <> "" and  request.Form("tahun") <> "" then
	tgl = request.Form("tgl")
	bln = request.Form("bln")
	thn = request.Form("tahun")
else
	tgl = day(date)
	bln = month(date)
	thn = year(date)
end if



	

	trk1Ar = 0
	trk1Ar = 0
	query = "SELECT  SUM(netto) AS jml, COUNT(KATEGORI) AS rit FROM  V_WEB_BRUTO WHERE  (katlahan='TRK A') and ( pbrk='KB1' ) and ((DAY(tgl) = '" & tgl & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } > 5) OR (DAY(tgl) = '" & tgl +1 & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } < 6))"

			
			set rec = server.CreateObject("ADODB.RECORDSET")
  				rec.Open query , conn, 3, 3
				
		if rec.fields("rit")<>0 then
			trk1Ar = rec.fields("rit")
			trk1Ab = rec.fields("jml")
		else 
			trk1Ar = 0
			trk1Ab = 0
		end if

			rec.close
			set rec = nothing	
		


		trk1AMr = 0
		rtk1AMb = 0
		query = "SELECT  SUM(netto) AS jml, COUNT(KATEGORI) AS rit FROM  V_WEB_BRUTO WHERE  (katlahan='TRK A-M') and ( pbrk='KB1' ) and ((DAY(tgl) = '" & tgl & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } > 5) OR (DAY(tgl) = '" & tgl +1 & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } < 6))"
			
			set rec = server.CreateObject("ADODB.RECORDSET")
  				rec.Open query , conn, 3, 3
				
		if rec.fields("rit")<>0 then
			trk1AMr = rec.fields("rit")
			trk1AMb = rec.fields("jml")
		else 
			trk1AMr = 0
			rtk1AMb = 0
		end if
		
		
		
			rec.close
			set rec = nothing	
				



	trk1Br = 0
	trk1Bb = 0
	query = "SELECT  SUM(netto) AS jml, COUNT(KATEGORI) AS rit FROM  V_WEB_BRUTO WHERE  (katlahan='TRK B') and ( pbrk='KB1' ) and ((DAY(tgl) = '" & tgl & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } > 5) OR (DAY(tgl) = '" & tgl +1 & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } < 6))"
			
			set rec = server.CreateObject("ADODB.RECORDSET")
  				rec.Open query , conn, 3, 3
				
		if rec.fields("rit")<>0 then
			trk1Br = rec.fields("rit")
			trk1Bb = rec.fields("jml")
		else 
			trk1Br = 0
			trk1Bb = 0
		end if
		
		
		
			rec.close
			set rec = nothing

	ts1r = 0
	ts1b = 0
	query = "SELECT  SUM(netto) AS jml, COUNT(KATEGORI) AS rit FROM  V_WEB_BRUTO WHERE  (katlahan='TS') and ( pbrk='KB1' ) and((DAY(tgl) = '" & tgl & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } > 5) OR (DAY(tgl) = '" & tgl +1 & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } < 6))"
			
			set rec = server.CreateObject("ADODB.RECORDSET")
  				rec.Open query , conn, 3, 3
				
		if rec.fields("rit")<>0 then
			ts1r = rec.fields("rit")
			ts1b = rec.fields("jml")
		else 
			ts1r = 0
			ts1b = 0
		end if
		
		
		
			rec.close
			set rec = nothing

				

%>


	<form class="form-2" action="" method="post" name="frmlogin" id="frmlogin" onSubmit="" width="100%">
			<table class="table-fill-judul" width="100%" border="0">
				<thead>

				<tr class="tr-judul">
				<th class="th-judul">PROGRES BIAYA KEBUN</th>
				</tr>

				<tr class="tr-judul">
				<th class="th-judul">
					Tanggal 
          <input name="tgl" type="text" id="tgl" size="3" maxlength="3" value="<%=tgl%>" />
          Bulan 
				
				​​</th>
				</tr>
				</thead
			</table>
	</form>
			
			
		
  <table  class="table-fill" width="100%">
  
    <thead>
	<tr>
		<th class="text-left" width="25px"><div align="center">No</th>
		<th class="text-left" width="29%"><div align="center">Kategori</th>
		<th class="text-left" width="31%"><div align="center">Rit</th>
		<th class="text-left" width="31%"><div align="center">Berat (Ku)</th>
	</tr>
	</thead>
	
	<thead>
	<tr>
		 <td width="100%"  colspan="4"  bgcolor="#006600"><div align="center"><font size="2"><strong><font color="#FFFFFF">KB I</font></strong></font></div></td>
	</tr>
	</thead>
	
<tr> 
      <td height="25"><div align="center" class="style13">1</div></td>
      <td height="25"><span class="style13">TS</span></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(ts1r,0) %>
        </div></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(ts1b,1) %>
        </div></td>
    </tr>
	

   

	
    <tr> 
      <td height="25"><div align="center" class="style13">2</div></td>
      <td height="25"><span class="style13">TRK A</span></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk1Ar,0) %>
        </div></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk1Ab,1) %>
        </div></td>
    </tr>
    <tr> 
      <td height="25"><div align="center" class="style13">3</div></td>
      <td height="25"><span class="style13">TRK A-M</span></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk1AMr,0) %>
        </div></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk1AMb,1) %>
        </div></td>
    </tr>
    
    </tr>
    <tr> 
      <td height="25"><div align="center" class="style13">4</div></td>
      <td height="25"><span class="style13">TRK B</span></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk1Br,0) %>
        </div></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk1Bb,1) %>
        </div></td>
    </tr>


    <tr class="tr-judul">
	<thead>
      <th class="th-judul" height="30" colspan="2"> <div align="center"><strong>JUMLAH KB I</strong></div></th>
      <th class="th-judul" height="30"> <div align="right" class="style13"> 
          <%=FORMATNUMBER(trk1Ar+trk1AMr+trk1Br+ts1r,0) %>
        </div></th>
      <th class="th-judul" height="30"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk1Ab+trk1AMb+trk1Bb+ts1b,1) %>
        </div></th>
      <% kb1r=trk1Ar+trk1AMr+trk1Br+ts1r
		   kb1b=trk1Ab+trk1AMb+trk1Bb+ts1b
		%>
	</thead>
    </tr>
    <%
	
	trk2Ar = 0
	trk2Ar = 0
	query = "SELECT  SUM(netto) AS jml, COUNT(KATEGORI) AS rit FROM  V_WEB_BRUTO WHERE  (katlahan='TRK A') and ( pbrk='KB2' ) and ((DAY(tgl) = '" & tgl & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } > 5) OR (DAY(tgl) = '" & tgl +1 & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } < 6))"

			
			set rec = server.CreateObject("ADODB.RECORDSET")
  				rec.Open query , conn, 3, 3
				
		if rec.fields("rit")<>0 then
			trk2Ar = rec.fields("rit")
			trk2Ab = rec.fields("jml")
		else 
			trk2Ar = 0
			trk2Ab = 0
		end if

			rec.close
			set rec = nothing	
		


		trk2AMr = 0
		rtk2AMb = 0
		query = "SELECT  SUM(netto) AS jml, COUNT(KATEGORI) AS rit FROM  V_WEB_BRUTO WHERE  (katlahan='TRK A-M') and ( pbrk='KB2' ) and ((DAY(tgl) = '" & tgl & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } > 5) OR (DAY(tgl) = '" & tgl +1 & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } < 6))"
			
			set rec = server.CreateObject("ADODB.RECORDSET")
  				rec.Open query , conn, 3, 3
				
		if rec.fields("rit")<>0 then
			trk2AMr = rec.fields("rit")
			trk2AMb = rec.fields("jml")
		else 
			trk2AMr = 0
			rtk2AMb = 0
		end if
		
		
		
			rec.close
			set rec = nothing	
				



	trk2Br = 0
	trk2Bb = 0
	query = "SELECT  SUM(netto) AS jml, COUNT(KATEGORI) AS rit FROM  V_WEB_BRUTO WHERE  (katlahan='TRK B') and ( pbrk='KB2' ) and ((DAY(tgl) = '" & tgl & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } > 5) OR (DAY(tgl) = '" & tgl +1 & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } < 6))"
			
			set rec = server.CreateObject("ADODB.RECORDSET")
  				rec.Open query , conn, 3, 3
				
		if rec.fields("rit")<>0 then
			trk2Br = rec.fields("rit")
			trk2Bb = rec.fields("jml")
		else 
			trk2Br = 0
			trk2Bb = 0
		end if
		
		
		
			rec.close
			set rec = nothing

	ts2r = 0
	ts2b = 0
	query = "SELECT  SUM(netto) AS jml, COUNT(KATEGORI) AS rit FROM  V_WEB_BRUTO WHERE  (katlahan='TS') and ( pbrk='KB2' ) and((DAY(tgl) = '" & tgl & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } > 5) OR (DAY(tgl) = '" & tgl +1 & "') AND (MONTH(tgl) = '" & bln & "') AND (YEAR(tgl) = '" & thn & "') AND ({ fn HOUR(jam) } < 6))"
			
			set rec = server.CreateObject("ADODB.RECORDSET")
  				rec.Open query , conn, 3, 3
				
		if rec.fields("rit")<>0 then
			ts2r = rec.fields("rit")
			ts2b = rec.fields("jml")
		else 
			ts2r = 0
			ts2b = 0
		end if
		
		
		
			rec.close
			set rec = nothing

				

%>
    
	<tr>
		 <td width="100%"  colspan="4"  bgcolor="#006600"><div align="center"><font size="2"><strong><font color="#FFFFFF">KB I</font></strong></font></div></td>
	</tr>
	
<tr> 
      <td height="25"><div align="center" class="style13">1</div></td>
      <td height="25"><span class="style13">TS</span></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(ts2r,0) %>
        </div></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(ts2b,1) %>
        </div></td>
    </tr>
 
    <tr> 
      <td height="25"><div align="center" class="style13">2</div></td>
      <td height="25"><span class="style13">TRK A</span></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk2Ar,0) %>
        </div></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk2Ab,1) %>
        </div></td>
    </tr>
    <tr> 
      <td height="25"><div align="center" class="style13">3</div></td>
      <td height="25"><span class="style13">TRK A-M</span></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk2AMr,0) %>
        </div></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk2AMb,1) %>
        </div></td>
    </tr>
    
    <tr> 
      <td height="25"><div align="center" class="style13">4</div></td>
      <td height="25"><span class="style13">TRK B</span></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk2Br,0) %>
        </div></td>
      <td height="25"><div align="right" class="style13"> 
          <%=FORMATNUMBER(trk2Bb,1) %>
        </div></td>
    </tr>

        <%
		 	kb2r=trk2Ar+trk2AMr+trk2Br+ts2r
		    kb2b=trk2Ab+trk2AMb+trk2Bb+ts2b
		%>

    <tr> 
	<thead>
      <th class="th-judul" height="30" colspan="2"><div align="center"><strong>JUMLAH KB II</strong></div></th>
      <th class="th-judul" height="30"> <div align="right" class="style13"> 
          <%=FORMATNUMBER(kb2r,0) %>
        </div></th>
      <th class="th-judul" height="30"><div align="right" class="style13"> 
          <%=FORMATNUMBER(kb2b,1) %>
        </div></th>

	</thead>
    </tr>
    <tr> 
	<thead>
      <th height="30" colspan="2"><div align="center" class="style12 style13"><strong>TOTAL 
          KB I + KB II</strong></div></th>
      <th height="30"><div align="right" class="style13"> 
          <%=FORMATNUMBER(kb1r+kb2r,0) %>
        </div></th>
      <th height="30"><div align="right" class="style13"> 
          <%=FORMATNUMBER(kb1b+kb2b,1) %>
        </div></th>
	<thead>
    </tr>
  </table>
  <%
  
  			conn.close
			set conn = nothing		

			''end if
%>
				
			
			
	</div>
	
</p>
</body>
</html>