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
        <th class="th-judul" width="200px">EKSPEDISI SPTA <% =Ucase(request.querystring("pos")) & " - " &request.querystring("userid")%></th>
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
  
  if Ucase(request.querystring("pos"))="PG KREBET BARU" or request.querystring("userid")="ARDI" or request.querystring("userid")="KS007" or request.querystring("userid")="PL049" then
  PERINTAH =  "select top 250 spa as NoSPTA,Kdptn as Register,namaptn,NoKend,convert(varchar,TglEntry,8) as TglCetak,Pos from dbo.vSPTA_Pos_ekspedisi where convert(varchar,tglspta,120)=convert(varchar,tglview,120) order by TglEntry desc"
  else
  PERINTAH =  "select  spa as NoSPTA,Kdptn as Register,namaptn,NoKend,TglEntry as TglCetak,Pos from dbo.vSPTA_Pos_ekspedisi where  TglSPTA=TglView and upper(pos)='"&Ucase(request.querystring("pos"))&"' and    convert(varchar,tglspta,120)=convert(varchar,tglview,120) order by TglEntry desc"
  end if

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
        <th class="text-left" width="5%"><div align="center">No</th>
        <th class="text-left" width="5%"><div align="center">No SPTA</th>
        <th class="text-left" width="20%"><div align="center">Register</th>
        <th class="text-left" width="5%"><div align="center">No Kend</th>
        <th class="text-left" width="5%"><div align="center">Jam</th>
  </tr>
  </thead>
<% 
  while not rec.eof
  %>
    
    <tr>

     
      <td><div align="center"><%=i %></div></td>
      <td><div><%= rec.fields ("NoSPTA") %></div></td>
      <td><div><%
      if  request.querystring("userid")="ARDI" then 
      response.write(rec.fields ("Register") &"<br>"& rec.fields ("namaptn") &"<br>"& rec.fields ("Pos"))
      else
      response.write(rec.fields ("Register") &"<br>"& rec.fields ("namaptn"))
      end if

      %></div></td>
      <td><div><%= rec.fields ("NoKend") %></div></td>
      <td><div><%= rec.fields ("TglCetak") %></div></td>
      
    </tr>
    
  
  
  
  
  
  <% 
  i=i+1


  rec.movenext 
  wend  
  %>
  
  
    <thead>
    <tr>
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