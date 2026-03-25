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
        <th class="th-judul" width="200px">JATAH SPTA <% =Ucase(request.querystring("pos")) & " - " &request.querystring("userid")%></th>
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

 vuserid=request.querystring("userid")
 vurl="JatahSPTAmobile.asp?userid="&vuserid&"&sort="
 
 vsort="POS"

 if request.querystring("sort")="sisa" then
  vsort= "sisa" '"sum(jml+0 - ISNULL(jmlcetak, 0))"
 elseif request.querystring("sort")<>"" then
  vsort=request.querystring("sort")
 end if
  
  if  request.querystring("userid")="ARDI" or request.querystring("userid")="M" or left(request.querystring("userid"),2)="PT"  then
  PERINTAH =  "select kdkel, kelompok, sum(sisa) as sisa, Pos From vJatahSPTApos_Kelompok where isnull(Pos,'')<>'' and pos<>'PG Krebet Baru' group by kdkel,kelompok,Pos order by "&vsort&",kdkel,kelompok"
  else
  PERINTAH =  "select kdkel, kelompok, sum(sisa) as sisa, Pos From vJatahSPTApos_Kelompok where isnull(Pos,'')<>'' and upper(pos) in (select upper(pos) from vpostebu where wilayah=(select Wilayah from vPosTebu where upper(Pos) = upper('"&Request.QueryString("pos")&"'))) group by kdkel,kelompok,Pos order by "&vsort&",kdkel,kelompok"
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
        <th class="text-left" width="5%"><a href=<% =vurl&"" %>><div align="center"><font color="#ffffff">No</font></div></a></th>
        <th class="text-left" width="5%"><div align="center"><a href=<% =vurl&"kdkel" %> ><div align="center"><font color="#ffffff">Kd. Kel</font></a></div></th>
        <th class="text-left" width="20%"><div align="center"><a href=<% =vurl&"kelompok" %> ><div align="center"><font color="#ffffff">Kelompok</font></a></div></th>
        <th class="text-left" width="5%"><div align="center"><a href=<% =vurl&"sisa" %> ><div align="center"><font color="#ffffff">Sisa Jatah</font></a></div></th>
        <th class="text-left" width="5%"><div align="center"><a href=<% =vurl&"POS" %> ><div align="center"><font color="#ffffff">Pos</font></a></div></th>
  </tr>
  </thead>
<% 
  while not rec.eof
  %>
    
    <tr>

     
      <td><div align="center"><%=i %></div></td>
      <td><div><%= rec.fields ("kdkel") %></div></td>
      <td><div><%= rec.fields ("kelompok") %></div></td>
      <td><div align="right"><%= rec.fields ("sisa") %></div></td>

      <td><div><%= rec.fields ("Pos") %></div></td>
      
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