
  <!--#include file="connSmartFarming.inc" -->



<% 
  
 On Error Resume Next 
'membuat query'
            userid=request.querystring("user_id")
            'response.write(userid)
            'PERINTAH =  "SELECT kdid, tgl, nomaster, iddevice, koordinatx, koordinaty, lokasifile, lokasi, nama_user FROM vLaporanAbsensi where CONVERT(VARCHAR, tgl, 101)=CONVERT(VARCHAR, getdate(), 101) and iddevice='"&userid&"' order by tgl"
            PERINTAH =  "SELECT getdate() as tgl"


          

            'Ambil data'
            set rd = server.CreateObject("ADODB.RECORDSET")
            rd.Open querytbl, conn,3,1
            i = 1

            jsonString = ""

            while not rd.eof
              
              jsonString = jsonString

            


              recd = "{"
              For each item in rd.Fields
                fd = item.Name
                recd = recd & """" & item.Name & """" & " : " & """" & rd.fields (fd) & """," 
              Next


              jsonString = jsonString & recd
              jsonString = left(jsonString,len(jsonString)-1) & "},"


              i=i+1
              rd.movenext 
            wend

            response.write("[" & left(jsonString,len(jsonString)-1) & "]")

  

%>


