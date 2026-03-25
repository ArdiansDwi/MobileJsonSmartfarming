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
	
	<form class="form-2" action="" method="post" name="frmlogin" id="frmlogin" onSubmit="" width="100%">
			<table class="table-fill-judul" width="100%" border="0">
				<thead>
				<tr class="tr-judul">
				<th class="th-judul" width="300px">Pemasukan Per Kebun</th>
				<th class="th-judul">
				    Hari ke : 
				<select name="harike" id="harike" class="select">
				<%
				if request.Form("harike")<>"" then
				%> <option value="<%=request.Form("harike") %>" > <%=request.Form("harike") %> </option> <%
				end if
				perintah = "SELECT HARIke FROM v_web_pemasukan_tebu GROUP BY HARIke ORDER BY HARIke  DESC"
				set rec  = server.CreateObject("ADODB.RECORDSET")
				rec.Open PERINTAH , conn,3,1
      			
				WHILE NOT REC.EOF
				HRK = rec.fields("harike")
				%>
				<option value="<%=hrk%>" > <%=hrk%> </option>
				<%
		 		REC.MOVENEXT
				WEND
				%>
				</select >
				
					&nbsp;&nbsp;Kategori :
					
				<select name="kategori" class="select" id="kategori">
				<%
				if request.Form("kategori")<>"" then
				%> <option value="<%=request.Form("kategori") %>" > <%=request.Form("kategori") %> </option> <%
				end if
				
				perintah = "SELECT kategori FROM v_WebPemasukanTebuPerKbn GROUP BY kategori ORDER BY kategori "
	  			
				set rec  = server.CreateObject("ADODB.RECORDSET")
				rec.Open PERINTAH , conn,3,1
      			
				WHILE NOT REC.EOF
				ktgr = rec.fields("kategori")
				%>
				<option value="<%=ktgr%>" > 
				<%=ktgr%>
				</option>
				<%
		 		REC.MOVENEXT
				WEND
				%>
				</select>			
				
				&nbsp;
					
					
					
					<input type="submit" name="submit" value="Tampilkan"  btnsubmit="btnSubmit" >
				​​</th>
				</tr>
				</thead
			</table>
	</form>
			
	
	
	<div class="content">
	
			<!-- Isi halaman disini -->
			
				
				
	<table class="table-fill" width="100%">
    <thead>
		<tr>
		<th class="text-left" rowspan="2" width="25px"><div align="center">No</th>
		<th class="text-left" rowspan="2" width="20%"><div align="center">Register</th>
		<th class="text-left" rowspan="2" width="25%"><div align="center" width="100%">Kelompok</th>
		<th class="text-left" rowspan="2" width="25%"><div align="center">Kebun</th>
		<th class="text-left" rowspan="2" width="8%"><div align="center" width="50px">Luas (Ha)</th>
		<th class="text-left" rowspan="2" width="8%"><div align="center">Taksasi</th>
				
      <th class="text-left" colspan="2"><div align="center">Hari Ini</th>
      <th class="text-left" colspan="2"><div align="center">S/d Hari Ini</th>
      <th class="text-left"  rowspan="2" width="8%"><div align="center">Produktifitas (Ku/Ha)</td>
      <th class="text-left"  rowspan="2" width="8%"><div align="center">(%) Terhadap Taks</td>
    </tr>
    <tr>
     <th class="text-left" width="8%"><div align="center">Jml Rit</th>
     <th class="text-left" width="8%"><div align="center">Jml Berat (Ku)</th>
     <th class="text-left" width="8%"><div align="center">Jml Rit</th>
     <th class="text-left" width="8%"><div align="center">Jml Berat (Ku)</th>
	</tr>
	</thead>
<% 

	perintah = "select sum(berat) as berat, register,count(register) as rit,kelompok,alamat,luas,taksasi from v_WebPemasukanTebuPerKbn where harike = '" & request.Form("harike") &"' and kategori = '" & request.Form("kategori") &"' group by register,kelompok,alamat,luas,taksasi"' and not kw_netto is null  group by id_induk order by left(id_induk,4)"  

	
	
	set rec = server.CreateObject("ADODB.RECORDSET")
  			rec.Open PERINTAH , conn, 3, 3
	i=1
	while not rec.eof

%>
    
    <tr>
      <td><div align="center"><%=i %></div></td>
      <td><%=rec.fields("register") %></td>
      <% 
      
	'  perintah1 = "select * from tab_register where id_induk = '" & rec.fields("ID_INDUK") & "'"
	  
'	  	set rec1 = server.CreateObject("ADODB.RECORDSET")
 ' 			rec1.Open PERINTAH1 , conn, 3, 3
      
'	  if not rec1.eof then
	  
	  
	  %>
	  <td><%=rec.fields("kelompok") %></td>
      <td><%=rec.fields("alamat") %></td>
      <td><div align="right"><%=formatnumber(rec.fields("luas"),3) %></div></td>
      <td><div align="right"><%=rec.fields("taksasi") %></div></td>
     
      <% 
	  luas=rec.fields("luas")
	  taks=rec.fields("taksasi")
	  ttaks=ttaks+(taks*luas)
	  tluas=tluas+luas
	'  end if
	 ' rec1.close
	'	set rec1=nothing
	  %>
      
      
      <td><div align="right"><%=formatnumber(rec.fields("rit"),0) %></div></td>
      <td><div align="right"><%=formatnumber(rec.fields("berat"),0) %></div></td>
      
      <% 
 
	  rit=rit+rec.fields("rit")
	  berat=berat+rec.fields("berat")
	  
	  
	  perintah1 = " select sum(berat) as berat, count(register) as rit from v_WebPemasukanTebuPerKbn where (harike between '001' and '" & request.Form("harike") & "') and register = '" & rec.fields("register") & "' group by register "
	  
	  
	  	set rec1 = server.CreateObject("ADODB.RECORDSET")
  			rec1.Open PERINTAH1 , conn, 3, 3
      
	  if not rec1.eof then
	  
	  
	  
      
      %>
      
      <td><div align="right"><%=formatnumber(rec1.fields("rit"),0) %></div></td>
      <td><div align="right"><%=formatnumber(rec1.fields("berat"),0) %></div></td>
      
    
      <% 
	  trit=trit+rec1.fields("rit")
	  tberat=tberat+rec1.fields("berat")
	  IF luas<>0 then
	  prod=rec1.fields("berat")/luas
	  else
	  prod=0
	  end if
	  
	   rec1.close
		set rec1=nothing
	  end if 
	  if prod =0 or taks = 0 then
	  warna =0
	  else
	  
	  warna = prod/taks*100
	  end if
	  %>
      
      
      
      <% if warna >= 100 then %>
         <td bgcolor="#FF66FF"><div align="right"><%=round(prod,0)%></div></td>
      <td bgcolor="#FF66FF"><div align="right"><%=round(prod/taks*100,2)%></div></td>
     
    	<%else%>
      <td width="2%" ><div align="right"><%=round(prod,0)%></div></td>
	  <% if warna=0 then %>
	   <td width="2%"><div align="right"><%=0%></div></td>   
	   <% else %>
     <td width="2%"><div align="right"><%=round(prod/taks*100,2)%></div></td>   
     
    <%
	end if 
	end if %>    
    </tr>
    
	
	
	
	
	
	<% 
	i=i+1
	rec.movenext 
		wend
		
	%>
    <thead>
		<tr>
		<th class="text-left"  colspan="4"><div align="center">Jumlah</th>
       <% if tluas <> 0 or ttaks <>0  then %>
      <th class="text-left"  ><div align="right"><%=formatnumber(tluas,3)%></div></td>
      <th class="text-left"  ><div align="right"><%=formatnumber(ttaks/tluas,0)%></div></td>
      <th class="text-left"  ><div align="right"><%=formatnumber(rit,0)%></div></td>
      <th class="text-left"  ><div align="right"><%=formatnumber(berat,0)%></div></td>
      <th class="text-left"  ><div align="right"><%=formatnumber(trit,0)%></div></td>
      <th class="text-left"  ><div align="right"><%=formatnumber(tberat,0)%></div></td>
      <th class="text-left"  ><div align="right"><%=formatnumber(tberat/tluas,0)%></div></td>
      <th class="text-left"  ><div align="right"><%=formatnumber(0*100,2)%></div></td>
      <% else %>
      <th class="text-left"  ><div align="right"><%=0%></div></td>
      <th class="text-left"  ><div align="right"><%=0%></div></td>
      <th class="text-left"   width="2%"><div align="right"><%=0%></div></td>
      <th class="text-left"   width="2%"><div align="right"><%=0%></div></td>
      <th class="text-left"   width="2%"><div align="right"><%=0%></div></td>
      <th class="text-left"   width="2%"><div align="right"><%=0%></div></td>
      <th class="text-left"   width="2%"><div align="right"><%=0%></div></td>
      <th class="text-left"   width="2%"><div align="right"><%=0%></div></td>
      <% end if %>
    </tr>
	</thead>
  </table>

 <% rec.close
		set rec=nothing
		conn.close
		set conn = nothing %>
				
			
			
	</div>
	
</p>
</body>
</html>