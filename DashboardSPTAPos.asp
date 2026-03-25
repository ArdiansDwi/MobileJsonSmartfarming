<html class="no-js">

<head>
  <meta charset='UTF-8'>
   <meta http-equiv="refresh" content="300" >
  
  <title>Dashboard SPTA Pos</title>
  
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
<link href="data.css" rel="stylesheet" type="text/css"/>

</head>

<!--#include file="conn.inc" -->



<body background="Images/body.png">
  <p align="left">
  <!-- Paste this code after body tag -->
  <div class="se-pre-con"></div>
  <!-- Ends -->
  


      <table class="table-fill-judul" width="100%" border="0">
        <tr class="tr-judul">
        <th class="th-judul" width="200px">DASHBOARD SPTA POS PER HARI</th>
        </tr>
      </table>
      
    <div class="content">
  
      <!-- Isi halaman disini -->

  <%
  function format_date(byval vd_date) 
    if IsNull(vd_date) or not IsDate(vd_date) then 
      format_date = "" 
     exit function 
   end if 
   format_date = Day(vd_date) & "/" & Month(vd_date) & "/" & Year(vd_date) 
end function 
  
  PERINTAH =  "SELECT [pos],[HrIni],[Kemarin],[DuaHrKemarin],[SdHrIni] FROM vSPTA_dashboard_mobile order by pos"
 set rec = server.CreateObject("ADODB.RECORDSET")
  rec.Open PERINTAH , conn, 1, 3
  i = 1

    jHrIni=0
    jKemarin=0

    jDuaHrKemarin=0
    jSdHrIni=0

  %>    
        
        
  <table class="table-fill" width="100%">
    <thead>
  <tr>
        <th class="text-left" width="5%" rowspan="2"><div align="center">No</th>
        <th class="text-left" width="30%" rowspan="2"><div align="center">Pos</th>
        <th class="text-left" width="5%" colspan="4"><div align="center">SPTA</th>
      </tr>
      <tr>
        <th class="text-left" width="15%"><div align="center">Hr ini</th>
        <th class="text-left" width="15%"><div align="center">Kemarin</th>
          <th class="text-left" width="15%"><div align="center">2 Hr Kmrn</th>
        <th class="text-left" width="15%"><div align="center">Sd Hr ini</th>
      </tr>
  </thead>
<% 
  while not rec.eof
  %>
      
    <%if i mod 2 = 0 then
    response.write("<tr bgcolor='#f1fbee'>")
  elseif rec.fields ("pos") = "Pos Dampit" then
    response.write("<tr bgcolor='#f1fbee'>")
  elseif rec.fields ("pos") = "Pos Lumajang" then
    response.write("<tr bgcolor='#eef4fb'>")
  elseif rec.fields ("pos") = "Pos Selorejo" then
    response.write("<tr bgcolor='#fbfbd9'>")
  elseif rec.fields ("pos") = "Pos Wonokerto" then
    response.write("<tr bgcolor='#fbe8e3'>")
  else
  end if%>
     
      <td><div align="center"><%=i %></div></td>
      <td><div><%= rec.fields ("pos") %></div></td>
      <td><div align="right"><%=formatnumber(rec.fields ("HrIni") ,0) %></div></td>
      <td><div align="right"><%=formatnumber(rec.fields ("Kemarin") ,0) %></div></td>
      <td><div align="right"><%=formatnumber(rec.fields ("DuaHrKemarin") ,0) %></div></td>
      <td><div align="right"><%=formatnumber(rec.fields ("SdHrIni") ,0) %></div></td>
      
    </tr>
    
  
  
  
  
  
  <% 
  i=i+1

    jHrIni=jHrIni+rec.fields ("HrIni") 
    jKemarin=jKemarin+rec.fields ("Kemarin") 

    jDuaHrKemarin=jDuaHrKemarin+rec.fields ("DuaHrKemarin") 
    jSdHrIni=jSdHrIni+rec.fields ("SdHrIni") 

  rec.movenext 
  wend  
  %>
  
  
    <thead>
    <tr>
    <th class="text-left"  colspan="2"><div align="center">Jumlah</th>
    <th class="text-left"  ><div align="right"><%=formatnumber(jHrIni,0)%></div></td>
    <th class="text-left"  ><div align="right"><%=formatnumber(jKemarin,0)%></div></td>
    <th class="text-left"  ><div align="right"><%=formatnumber(jDuaHrKemarin,0)%></div></td>
    <th class="text-left"  ><div align="right"><%=formatnumber(jSdHrIni,0)%></div></td>
    </tr>
  </thead>
  </table>


 <% rec.close
    set rec=nothing
 
    %>






    <table class="table-fill-judul" width="100%" border="0">
            <tr class="tr-judul">
            <th class="th-judul" width="200px">SPTA S/D JAM YANG SAMA</th>
            </tr>
          </table>
<%
  
  PERINTAH =  "SELECT [pos],[HrIni],[Kemarin],[DuaHrKemarin],[SdHrIni] FROM vSPTA_dashboard_jamsama_mobile order by pos"
 set rec = server.CreateObject("ADODB.RECORDSET")
  rec.Open PERINTAH , conn, 1, 3
  i = 1

    jHrIni=0
    jKemarin=0

    jDuaHrKemarin=0
    jSdHrIni=0

  %>    
        
        
  <table class="table-fill" width="100%">
    <thead>
  <tr>
        <th class="text-left" width="5%" rowspan="2"><div align="center">No</th>
        <th class="text-left" width="30%" rowspan="2"><div align="center">Pos</th>
        <th class="text-left" width="5%" colspan="4"><div align="center">SPTA</th>
      </tr>
      <tr>
        <th class="text-left" width="15%"><div align="center">Hr ini</th>
        <th class="text-left" width="15%"><div align="center">Kemarin</th>
          <th class="text-left" width="15%"><div align="center">2 Hr Kmrn</th>
        <th class="text-left" width="15%"><div align="center">Sd Hr ini</th>
      </tr>
  </thead>
<% 
  while not rec.eof
  %>
      
    <% if i mod 2 = 0 then
    response.write("<tr bgcolor='#f1fbee'>")
  elseif rec.fields ("pos") = "Pos Dampit" then
    response.write("<tr bgcolor='#f1fbee'>")
  elseif rec.fields ("pos") = "Pos Lumajang" then
    response.write("<tr bgcolor='#eef4fb'>")
  elseif rec.fields ("pos") = "Pos Selorejo" then
    response.write("<tr bgcolor='#fbfbd9'>")
  elseif rec.fields ("pos") = "Pos Wonokerto" then
    response.write("<tr bgcolor='#fbe8e3'>")
  else
  end if%>
     
      <td><div align="center"><%=i %></div></td>
      <td><div><%= rec.fields ("pos") %></div></td>
      <td><div align="right"><%=formatnumber(rec.fields ("HrIni") ,0) %></div></td>
      <td><div align="right"><%=formatnumber(rec.fields ("Kemarin") ,0) %></div></td>
      <td><div align="right"><%=formatnumber(rec.fields ("DuaHrKemarin") ,0) %></div></td>
      <td><div align="right"><%=formatnumber(rec.fields ("SdHrIni") ,0) %></div></td>
      
    </tr>
    
  
  
  
  
  
  <% 
  i=i+1

    jHrIni=jHrIni+rec.fields ("HrIni") 
    jKemarin=jKemarin+rec.fields ("Kemarin") 

    jDuaHrKemarin=jDuaHrKemarin+rec.fields ("DuaHrKemarin") 
    jSdHrIni=jSdHrIni+rec.fields ("SdHrIni") 

  rec.movenext 
  wend  
  %>
  
  
    <thead>
    <tr>
    <th class="text-left"  colspan="2"><div align="center">Jumlah</th>
    <th class="text-left"  ><div align="right"><%=formatnumber(jHrIni,0)%></div></td>
    <th class="text-left"  ><div align="right"><%=formatnumber(jKemarin,0)%></div></td>
    <th class="text-left"  ><div align="right"><%=formatnumber(jDuaHrKemarin,0)%></div></td>
    <th class="text-left"  ><div align="right"><%=formatnumber(jSdHrIni,0)%></div></td>
    </tr>
  </thead>
  </table>


 <% rec.close
    set rec=nothing
 
    %>









<!-- <table class="table-fill-judul" width="100%" border="0">
        <tr class="tr-judul">
        <th class="th-judul" width="200px">SPTA PER JAM</th>
        </tr>
      </table> -->
<%
 
 '  PERINTAH =  "SELECT [jam]+1 as jam,[pos],[HrIni],[Kemarin],[DuaHrKemarin],[SdHrIni] FROM vSPTA_dashboard_perjam_mobile order by pos, jam"
 ' set rec2 = server.CreateObject("ADODB.RECORDSET")
 '  rec2.Open PERINTAH , conn, 1, 3
 '  i = 1

 '    jHrIni=0
 '    jKemarin=0

 '    jDuaHrKemarin=0
 '    jSdHrIni=0

  %>    



  <!-- <table class="table-fill" width="100%">
    <thead>
  <tr>
        <th class="text-left" width="5%" rowspan="2"><div align="center">Jam</th>
        <th class="text-left" width="30%" rowspan="2"><div align="center">Pos</th>
        <th class="text-left" width="5%" colspan="4"><div align="center">SPTA</th>
      </tr>
      <tr>
        <th class="text-left" width="15%"><div align="center">Hr ini</th>
        <th class="text-left" width="15%"><div align="center">Kemarin</th>
          <th class="text-left" width="15%"><div align="center">2 Hr Kmrn</th>
        <th class="text-left" width="15%"><div align="center">Sd Hr ini</th>
      </tr>
  </thead> -->
<% 
  ' while not rec2.eof
  %>
    
    <% 
  '   if rec2.fields ("pos") = "Pos Dampit" then
  '   response.write("<tr bgcolor='#f1fbee'>")
  ' elseif rec2.fields ("pos") = "Pos Lumajang" then
  '   response.write("<tr bgcolor='#eef4fb'>")
  ' elseif rec2.fields ("pos") = "Pos Selorejo" then
  '   response.write("<tr bgcolor='#fbfbd9'>")
  ' elseif rec2.fields ("pos") = "Pos Wonokerto" then
  '   response.write("<tr bgcolor='#fbe8e3'>")
  ' else
  ' end if
  %>

     
      <!-- <td><div align="center"><%'= rec2.fields ("jam") %></div></td>
      <td><div><%'= rec2.fields ("pos") %></div></td>
      <td><div align="right"><%'=formatnumber(rec2.fields ("HrIni") ,0) %></div></td>
      <td><div align="right"><%'=formatnumber(rec2.fields ("Kemarin") ,0) %></div></td>
      <td><div align="right"><%'=formatnumber(rec2.fields ("DuaHrKemarin") ,0) %></div></td>
      <td><div align="right"><%'=formatnumber(rec2.fields ("SdHrIni") ,0) %></div></td>
      
    </tr> -->
    
  
  
  
  <% 
  ' i=i+1

  '   jHrIni=jHrIni+rec2.fields ("HrIni") 
  '   jKemarin=jKemarin+rec2.fields ("Kemarin") 

  '   jDuaHrKemarin=jDuaHrKemarin+rec2.fields ("DuaHrKemarin") 
  '   jSdHrIni=jSdHrIni+rec2.fields ("SdHrIni") 

  ' rec2.movenext 
  ' wend  
  %>
  
  
    <!-- <thead>
    <tr>
    <th class="text-left"  colspan="2"><div align="center">Jumlah</th>
    <th class="text-left"  ><div align="right"><%'=formatnumber(jHrIni,0)%></div></td>
    <th class="text-left"  ><div align="right"><%'=formatnumber(jKemarin,0)%></div></td>
    <th class="text-left"  ><div align="right"><%'=formatnumber(jDuaHrKemarin,0)%></div></td>
    <th class="text-left"  ><div align="right"><%'=formatnumber(jSdHrIni,0)%></div></td>
    </tr>
  </thead>
  </table> -->

<% 
' rec2.close
'     set rec2=nothing
    conn.close
    set conn = nothing 
    %>







        
      
      


  </div>
  
</p>
</body>
</html>