<html class="no-js">

  <head>
    <meta charset='UTF-8'>
    
    <title>Laporan Kunjungan Kebun</title>
    
    <script src="jquery.min.js"></script>
    <script src="modernizr.js"></script>
    <link rel="stylesheet" href="//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css">
    <link rel="stylesheet" href="/resources/demos/style.css">
    <script src="https://code.jquery.com/jquery-1.12.4.js"></script>
    <script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
    <script>
      $(window).load(function() {
        $(".se-pre-con").fadeOut("slow");;
      });
    </script>
    <script>
      $( function() {
        $( "#datepicker" ).datepicker();
      });
    </script>
    <link href="data.css" rel="stylesheet" type="text/css" />
  </head>

  <!--#include file="connSmartFarming.inc" -->

  <!-- <body background="Images/body.png"> -->
    <body>
    <p align="left">
      <div class="se-pre-con"></div>
        <div class="content">
          
          <%
            userid=request.querystring("user_id")
            'response.write(userid)
            PERINTAH =  "SELECT kdid, tgl, nomaster, iddevice, koordinatx, koordinaty, lokasifile, lokasi, nama_user FROM vLaporanAbsensi where CONVERT(VARCHAR, tgl, 101)=CONVERT(VARCHAR, getdate(), 101) and iddevice='"&userid&"' order by tgl"
            set rec = server.CreateObject("ADODB.RECORDSET")
            rec.Open PERINTAH , conn, 1, 3
            i = 1
          %>    
        
          <!-- PERINTAH =  "SELECT [uraian],[satuan],[kb1_hrini],[kb2_hrini] FROM vlhg_mobile order by convert(int,urut)" -->
      <table class="table-fill-judul" width="100%" border="0">
      <!--   <tr class="tr-judul">
        <th class="th-judul" width="100%" colspan="2">ABSENSI HARI INI</th>
        </tr> -->

             <thead>
              <tr>
                <th class="text-left" colspan="5" ></td>
              </tr>
            </thead>

            
            <% 
              while not rec.eof
            %>
    
            <tr >          
              <td rowspan="2"  width="20"><div align="center"><b><%=i %></b></div></td>
              <td colspan="0"><div><b><%= rec.fields ("tgl") %></b></div></td>
            </tr>
            
            <tr>          
             <!--  <td><div><img src="--><% '"foto/" &  rec.fields("lokasifile")  %><!--"   alt="Foto1" width="100%"></div></td> -->
              <td colspan="0"><div><%= rec.fields ("lokasi") %></div></td>
            </tr>


  
           
<!-- 
            <thead>
              <tr>
                <th class="text-left" colspan="5" ></td>
              </tr>
            </thead>
 -->




            <% 
              i=i+1
              rec.movenext 
              wend  
            %> 


          <% rec.close
            set rec=nothing
            conn.close
            set conn = nothing 
          %>

        </div>
      </div>
    </p>
  </body>
</html>