<html class="no-js">

  <head>
    <meta charset='UTF-8'>
    <meta http-equiv="refresh" content="300" >
    
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

  <body background="Images/body.png">
    <p align="left">
      <div class="se-pre-con"></div>
        <div class="content">
          
          <%
            function format_date(byval vd_date) 
              if IsNull(vd_date) or not IsDate(vd_date) then 
                format_date = "" 
                exit function 
              end if 
              format_date = Day(vd_date) & "/" & Month(vd_date) & "/" & Year(vd_date) 
            end function      
            
            If day(date) < 10 Then d = "0"&day(date) Else d = day(date) 
            If month(date) < 10 Then m = "0"&month(date) Else m = month(date)
            y = year(date)
            htgl = m&"/"&d&"/"&y
            If request.querystring("tgl") <> "" Then htgl = request.querystring("tgl") End If
            
            user = request.querystring("user_id")
            if (user="ARDI") or (user="M") or (user="R") or (user="KS007") or (user="KS005") or (user="KS006") then user = "CA001" end if
            
            PERINTAH =  "SELECT kdid, tgl, nomaster, iddevice, koordinatx, koordinaty, kodePetani, namaPetani, Alamat, judul, deskripsi, lokasifile1, lokasifile2, lokasifile3, lokasifile4, nama_user,deskripsiweb FROM vkunjungankebun where CONVERT(VARCHAR, tgl, 101)='"&htgl&"' and plpg_id+skk_id+skw_id+ca_id like'%"&user&"%' order by tgl"
            set rec = server.CreateObject("ADODB.RECORDSET")
            rec.Open PERINTAH , conn, 1, 3
            i = 1

          'response.write(user)
          'response.write(htgl)
          %>    

        
          <!-- PERINTAH =  "SELECT [uraian],[satuan],[kb1_hrini],[kb2_hrini] FROM vlhg_mobile order by convert(int,urut)" -->
          <table class="table-fill" width="100%">
            <thead>
              <tr>
                <th colspan="5" class="text-left" width="5%"><div align="center">LAPORAN KUNJUNGAN KEBUN</th>
              </tr> 
              <tr>
                <th colspan="5" class="text-left" width="5%"><div align="left"> TANGGAL 
                  <form method="GET" class="form-2">
                    <input type="text" id="datepicker" name="tgl" value="<%= htgl%>">
                    <!-- <input type="datetime" value="26/06/2020"/> -->
                    <!-- <button type="submit" >Tampilkan</button> -->
                    <input type="hidden" name="user_id" value="<% =request.querystring("user_id") %>">
                    <input type="submit" name="submit" value="Tampilkan"  btnsubmit="btnSubmit" >
                  </form>                   
                </th>
              </tr> 
              <tr>
                <th class="text-left" width="6%"><div align="center">No</th>
                <th class="text-left" width="47%"><div align="center">Petugas</th>
                <th class="text-left" width="47%"><div align="center">Petani</th>
              </tr>
            </thead>
            <% 
              while not rec.eof
            %>
    
            <tr rowspan="8">          
              <td><div align="center"><b><%=i %></b></div></td>
              <td><div><b><%= rec.fields ("nama_user") %></b></div></td>
              <td><div><b><%= rec.fields ("namaPetani") %></b></div></td>
            </tr>
            
            <tr>          
              <td><div align="center"></div></td>
              <td colspan="2"><div>
                
                <%= rec.fields ("deskripsiweb") %>
              
              </div></td>
            </tr>

            <tr>          
              <td><div align="center"></div></td>
              <td><div> <img src="<%= "foto/" &  rec.fields("lokasifile1") & "" %>" alt="Foto1" width="100%"></div></td>
              <td><div> <img src="<%= "foto/" &  rec.fields("lokasifile2") & "" %>" alt="Foto2" width="100%"></div></td>
            </tr>

            <tr>          
              <td><div align="center"></div></td>
              <td><div> <img src="<%= "foto/" &  rec.fields("lokasifile3") & "" %>" alt="Foto3" width="100%"></div></td>
              <td><div> <img src="<%= "foto/" &  rec.fields("lokasifile4") & "" %>" alt="Foto4" width="100%"></div></td>
            </tr>
           

            <thead>
              <tr>
                <th class="text-left" colspan="5" ></td>
              </tr>
            </thead>





            <% 
              i=i+1
              rec.movenext 
              wend  
            %> 
  
            <thead>
              <tr>
                <th class="text-left" colspan="5" ></td>
              </tr>
            </thead>
          </table>

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