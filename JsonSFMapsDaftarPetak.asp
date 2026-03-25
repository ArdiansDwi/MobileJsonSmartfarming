<!--#include file="connSmartFarming.inc" -->

<% 
 On Error Resume Next 
'membuat query'
 kolom=request.querystring("kolom")
 cari=request.querystring("cari")

if (kolom="Petugas") then  
	kolom="c.Nama_user"
elseif (kolom="Nopetak") then  
	kolom="a.nopetak"
elseif (kolom="Register") then  
	kolom="a.register"
elseif (kolom="Desa") then  
	kolom="a.desa"
elseif (kolom="Kecamatan") then  
	kolom="a.kecamatan"
else  
	kolom="c.Nama_user"
end if




 						if (kolom<>"" or cari<>"") then 
						'querytbl="SELECT inc, tglentry, user_id, register, nopetak, kecamatan, desa, luas, petugas, mg, brix FROM vPetakTeregisterDanBrix where "&kolom&" like '%"&cari&"%' order by tglentry"

						querytbl="SELECT a.inc, tglentry, a.user_id, a.register, a.nopetak, a.kecamatan, a.desa, a.luas, c.Nama_user as petugas, a.mg, b.brix FROM tblRegisterPetak a left join vbrixterakhir b on a.nopetak = b.nopetak inner join users_id c on a.user_id = c.user_id where "&kolom&" like '"&cari&"%' order by tglentry "
						else
						'querytbl="SELECT inc, tglentry, a.user_id, register, nopetak, kecamatan, desa, luas, b.Nama_user as petugas, mg FROM tblRegisterPetak a inner join users_id b on a.user_id = b.user_id order by nopetak"
						querytbl="SELECT inc, tglentry, user_id, register, nopetak, kecamatan, desa, luas, petugas, mg, brix FROM vPetakTeregisterDanBrix order by tglentry"
						end if



						'response.write(querytbl)
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