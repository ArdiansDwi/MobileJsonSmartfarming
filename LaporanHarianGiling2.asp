<html class="no-js">
  <head>
    <meta charset='UTF-8'>
    <meta http-equiv="refresh" content="300">    
    <title>Laporan Harian Giling</title>    
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
  <!--#include file="connLHG.inc" -->
  <body background="Images/body.png">
    <p align="left">
      <div class="se-pre-con"></div>
        <div class="content">
		<%
        function format_date(byval vd_date) 
		  If IsNull(vd_date) or not IsDate(vd_date) then 
			format_date = "" 
            exit function 
          end if 
          format_date = Day(vd_date) & "/" & Month(vd_date) & "/" & Year(vd_date) 
        end function             
            
		If day(date) < 10 Then d = "0"&day(date) Else d = day(date) 
		If month(date) < 10 Then m = "0"&month(date) Else m = month(date)
		y = year(date)
		htgl = m&"/"&d&"/"&y			
		'If request.querystring("tgl") <> "" Then htgl = request.querystring("tgl") End If
		If REQUEST.Form("tgl") <> "" Then htgl = REQUEST.Form("tgl") End If

		PERINTAH = "DELETE FROM tbldatalhg_tmp_sd"
        conn.Execute PERINTAH

		PERINTAH = "INSERT INTO tbldatalhg_tmp_sd (tanggal, hari, unit, tebumasuk, tebugiling_ts, tebugiling_trsa1, tebugiling_trsa2, tebugiling_trt, rawsugar_rs, rawsugar_shs, rawsugar_tetes, luasgiling_ts, luasgiling_tr, luasgiling_trs, luasgiling_trt, mututebu_tebukotor, mututebu_tebuterbakar, mututebu_luasterbakar, hablur_ts, hablur_trsa1, hablur_trsa2, hablur_trt, prodshs_ts, prodshs_trsa1, prodshs_trsa2, prodshs_trt, rendemen_ts, rendemen_trsa1, rendemen_trsa2, rendemen_trt, rendemen_total, prodtetes_ts, prodtetes_trsa1, prodtetes_trsa2, prodtetes_trt, curahhujan, offfarm_jamgiling, offfarm_jamberhenti_jumlah, offfarm_jamberhenti_luar, offfarm_jamberhenti_dalam, offfarm_jamberhenti_lainnya, offfarm_kapgilinginc, offfarm_kapgilinginc_2, offfarm_kapgilingexc, offfarm_kapgilingexc_2, offfarm_polnpp, offfarm_brixnpp, offfarm_poltebu, offfarm_brixtebu, offfarm_polniramentah, offfarm_brixniramentah, offfarm_imbibisi, offfarm_niramentah, offfarm_totalhablur, offfarm_prodshs, offfarm_shspg, offfarm_rendemen, offfarm_prodtetes, offfarm_tetespg, offfarm_polampas, offfarm_polblotong, offfarm_produap, offfarm_hpbtotal, offfarm_hpb1, offfarm_pshk, offfarm_kadarnira, offfarm_winterrendemen, offfarm_hpg, offfarm_bhract, offfarm_hktetes, energi_boiler_residu, energi_boiler_ido, energi_boiler_ifo, energi_boiler_kayu, energi_boiler_kopi, energi_boiler_sebuk, energi_boiler_finner, energi_boiler_batok, energi_boiler_sekam, energi_boiler_tatal, energi_boiler_ampas, energi_listrik, energi_pln, energi_diesel, energi_solar, energi_prodampas, energi_ampaskirim, energi_ampassimpan, kualitasshs_icumsa, kualitasshs_iu, kualitasshs_bjb, kualitasshs_kadarair, kualitasshs_kadarso2, kualitasshs_kadarabu, lingkungan_cair_jumlah, lingkungan_cair_bod, lingkungan_cair_cod, lingkungan_padat_blotong, lingkungan_padat_abuketel, lingkungan_udara_c02, lingkungan_udara_s02, gularetail_prodgula1, gularetail_penjgula1, gularetail_stokgula1, gularetail_prodgula05, gularetail_penjgula05, gularetail_stokgula05,ket_tanaman, ket_tuk, ket_qc, ket_instalasi, ket_pabrikasi, ver_tanaman, tglver_tanaman, ver_tuk, tglver_tuk, ver_qc, tglver_qc, ver_instalasi, tglver_instalasi, ver_pabrikasi, tglver_pabrikasi, offfarm_winterrendemen_sdhrini, energi_interkoneksi, offfarm_bahankeringampas, offfarm_kadarsabuttebu, pab_kapur, pab_belerang, pab_asam, pab_flokulan, pab_masakana, pab_masakanc, pab_masakand, pab_enzima, pab_kaporit, pab_hcl, pab_naoh, pab_lama_masakana, pab_lama_masakanc, pab_lama_masakand, thmg, offfarm_brixtetes, offfarm_poltetes, mututebu_tebukotor_manual) SELECT tanggal, hari, unit, tebumasuk, tebugiling_ts, tebugiling_trsa1, tebugiling_trsa2, tebugiling_trt, rawsugar_rs, rawsugar_shs, rawsugar_tetes, luasgiling_ts, luasgiling_tr, luasgiling_trs, luasgiling_trt, mututebu_tebukotor, mututebu_tebuterbakar, mututebu_luasterbakar, hablur_ts, hablur_trsa1, hablur_trsa2, hablur_trt, prodshs_ts, prodshs_trsa1, prodshs_trsa2, prodshs_trt, rendemen_ts, rendemen_trsa1, rendemen_trsa2, rendemen_trt, rendemen_total, prodtetes_ts, prodtetes_trsa1, prodtetes_trsa2, prodtetes_trt, curahhujan, offfarm_jamgiling, offfarm_jamberhenti_jumlah, offfarm_jamberhenti_luar, offfarm_jamberhenti_dalam, offfarm_jamberhenti_lainnya, offfarm_kapgilinginc, offfarm_kapgilinginc_2, offfarm_kapgilingexc, offfarm_kapgilingexc_2, offfarm_polnpp, offfarm_brixnpp, offfarm_poltebu, offfarm_brixtebu, offfarm_polniramentah, offfarm_brixniramentah, offfarm_imbibisi, offfarm_niramentah, offfarm_totalhablur, offfarm_prodshs, offfarm_shspg, offfarm_rendemen, offfarm_prodtetes, offfarm_tetespg, offfarm_polampas, offfarm_polblotong, offfarm_produap, offfarm_hpbtotal, offfarm_hpb1, offfarm_pshk, offfarm_kadarnira, offfarm_winterrendemen, offfarm_hpg, offfarm_bhract, offfarm_hktetes, energi_boiler_residu, energi_boiler_ido, energi_boiler_ifo, energi_boiler_kayu, energi_boiler_kopi, energi_boiler_sebuk, energi_boiler_finner, energi_boiler_batok, energi_boiler_sekam, energi_boiler_tatal, energi_boiler_ampas, energi_listrik, energi_pln, energi_diesel, energi_solar, energi_prodampas, energi_ampaskirim, energi_ampassimpan, kualitasshs_icumsa, kualitasshs_iu, kualitasshs_bjb, kualitasshs_kadarair, kualitasshs_kadarso2, kualitasshs_kadarabu, lingkungan_cair_jumlah, lingkungan_cair_bod, lingkungan_cair_cod, lingkungan_padat_blotong, lingkungan_padat_abuketel, lingkungan_udara_c02, lingkungan_udara_s02, gularetail_prodgula1, gularetail_penjgula1, gularetail_stokgula1, gularetail_prodgula05, gularetail_penjgula05, gularetail_stokgula05,ket_tanaman, ket_tuk, ket_qc, ket_instalasi, ket_pabrikasi, ver_tanaman, tglver_tanaman, ver_tuk, tglver_tuk, ver_qc, tglver_qc, ver_instalasi, tglver_instalasi, ver_pabrikasi, tglver_pabrikasi, offfarm_winterrendemen_sdhrini, energi_interkoneksi, offfarm_bahankeringampas, offfarm_kadarsabuttebu, pab_kapur, pab_belerang, pab_asam, pab_flokulan, pab_masakana, pab_masakanc, pab_masakand, pab_enzima, pab_kaporit, pab_hcl, pab_naoh, pab_lama_masakana, pab_lama_masakanc, pab_lama_masakand, thmg, offfarm_brixtetes, offfarm_poltetes, mututebu_tebukotor_manual FROM tbldatalhg where CONVERT(VARCHAR, tanggal, 101)<='"&htgl&"'"
        conn.Execute PERINTAH

		'PERINTAH="SELECT max(urut) as urut, uraian, max(satuan) as satuan, sum(isnull(kb1_hrini,0)) as kb1_hrini, sum(isnull(kb2_hrini,0)) as kb2_hrini, sum(isnull(kb1_sdhrini,0)) as kb1_sdhrini, sum(isnull(kb2_sdhrini,0)) as kb2_sdhrini FROM vlhg_mobile where CONVERT(VARCHAR, tanggal, 101)='"&htgl&"' group by uraian order by convert(int,max(urut))"

		PERINTAH="SELECT urut as urut, uraian, satuan as satuan, isnull(kb1_hrini,0) as kb1_hrini, isnull(kb2_hrini,0) as kb2_hrini, isnull(kb1_sdhrini,0) as kb1_sdhrini, isnull(kb2_sdhrini,0) as kb2_sdhrini FROM vlhg_mobile where CONVERT(VARCHAR, tanggal, 101)='"&htgl&"' order by convert(int,urut)"
        set rec = server.CreateObject("ADODB.RECORDSET")
        rec.Open PERINTAH, conn, 1, 3
        i = 1
		%>
		<table class="table-fill" width="100%">
            <thead>
              <tr>
                <th colspan="7" class="text-left" width="5%"><div align="center">LAPORAN HARIAN GILING</div></th>
              </tr> 
              <tr>
                <th colspan="7" class="text-left" width="5%"><div align="left"> TANGGAL 
                  <form method="post" class="form-2">
				    <input type="text" id="datepicker" name="tgl" value="<%=htgl%>">
					<input type="hidden" name="userid" value="<%=request.querystring("user_id")%>">
                    <input type="submit" name="submit" value="Tampilkan"  btnsubmit="btnSubmit" >
				  </form></div>
                </th>
              </tr> 
              <tr>
                <th rowspan="2" class="text-left" width="5%"><div align="center">No</div></th>
                <th rowspan="2" class="text-left" width="30%"><div align="center">Uraian</div></th>
                <th rowspan="2" class="text-left" width="10%"><div align="center">Sat</div></th>
                <th colspan="2" class="text-left" width="25%"><div align="center">KB I</div></th>
                <th colspan="2"class="text-left" width="25%"><div align="center">KB II</div></th>
              </tr>
              <tr>
                <th>H.I</th>
                <th>S/D H.I</th>
                <th>H.I</th>
                <th>S/D H.I</th>
              </tr>
            </thead> 
			<% 
              while not rec.eof
            %> 			
            <tr>          
              <td><div align="center"><%=i%></div></td>
              <td><div><%=rec.fields ("uraian")%></div></td>
              <td><div align="center"><%=rec.fields ("satuan")%></div></td>
              <td><div align="right"><%=formatnumber(rec.fields ("kb1_hrini"),2)%></div></td>
              <td><div align="right"><%=formatnumber(rec.fields ("kb1_sdhrini"),2)%></div></td>
              <td><div align="right"><%=formatnumber(rec.fields ("kb2_hrini"),2)%></div></td>
              <td><div align="right"><%=formatnumber(rec.fields ("kb2_sdhrini"),2)%></div></td>            
            </tr>			
            <% 
            i=i+1
            rec.movenext 
            wend  
            %>
            <thead>
              <tr>
                <th class="text-left" colspan="7"></th>
              </tr>
            </thead>
          </table>
          <%
		  rec.close
          set rec=nothing
          conn.close
          set conn = nothing 
          %>
      </div>
    </p>
  </body>
</html>