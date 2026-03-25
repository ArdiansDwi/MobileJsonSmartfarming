<!--#include file="SDM.inc" -->


<% 
	
 On Error Resume Next 




	if not Request.QueryString("key") = "" then

			'response.write(Request.QueryString("key") & "<hr>")
			
			qKey = "select inc,id,access from tblJsonKey where access=1 and id='" & Request.QueryString("key") & "'"

				set rec  = server.CreateObject("ADODB.RECORDSET")
				rec.Open qKey , conn,3,1

		

				if not rec.eof then
					key =  rec.fields("id")
				end if

				if not key = "" then

						'Jika key benar dan di berikan akses'

						tabel = Request.QueryString("d")
						filt = Request.QueryString("f")
						q = Request.QueryString("q")
						order = Request.QueryString("o")
						by = Request.QueryString("b")
						limit = Request.QueryString("L")
						mode = Request.QueryString("mode")


						'cek tabel'
						query ="select top 1 * from " & tabel
						set rs = server.CreateObject("ADODB.RECORDSET")
						rs.Open query, conn,3,1

						kolom = ""
						vfilt = ""
						vorder = ""

						For each item in rs.Fields
 							'response.write (item.Name&"<br>")
 							kolom = kolom & item.Name & ","

 								if filt = item.Name then
 								 	vfilt = item.Name
 							 	end if

 							 	if order = item.Name then
 								 	vorder = item.Name
 							 	end if
 							 	

						Next

						vtabel = tabel
						vkolom = left(kolom,len(kolom)-1)
						vq = q

						if limit = "" then
							vlimit=1
						elseif IsNumeric(limit) then
							vlimit = limit
						else
							vlimit=1
						end if

						if mode="debug" then
							response.write ("Key : " & rec.fields ("id") &"<hr>")
							response.write ("Table : " & vtabel &  "<hr>")
							response.write ("Colomn : " & vkolom & "<hr>")
							response.write ("Filter : " & vfilt & "<hr>")
							response.write ("Filter Value : " & vq & "<hr>")
							response.write ("Column Order : " & vorder & "<hr>")
							response.write ("ASC / DESC : " & by & "<hr>")
							response.write ("Limit : " & vlimit & "<hr>")
						end if

						'filter'
						if not vfilt="" then
							vfilt = " where " & vfilt & " like '" & vq & "' "
						else
							vorder = ""
						end if

						'order'
						if not vorder="" and by="asc" or not vorder="" and by="desc" then
							vorder = " order by " & vorder & " " & by
						else
							vorder = ""
						end if

						'membuat query'
						querytbl = "select top " & vlimit & " " & vkolom & " from " & vtabel & vfilt & vorder

						if mode="debug" then
						response.write ("query result : " & querytbl & "<hr>")
						end if

						'Ambil data'
						set rd = server.CreateObject("ADODB.RECORDSET")
						rd.Open querytbl, conn,3,1
						i = 1

						jsonString = ""

						while not rd.eof
							
							jsonString = jsonString

							recd = "{"
							For each item in rs.Fields
								fd = item.Name
								recd = recd & """" & item.Name & """" & " : " & """" & rd.fields (fd) & """," 
							Next

							jsonString = jsonString & recd
							jsonString = left(jsonString,len(jsonString)-1) & "},"


							i=i+1
							rd.movenext 
						wend

						' jsonString =  "{""menu"" : ""Bahan Baku"",""label"" : ""Bahan Baku"",""url"" : ""123"", ""imageurl"" : ""https://img.freepik.com/free-vector/modern-flat-design-isometric-landing-page_9209-1520.jpg?size=626&ext=jpg""} "

						response.write("[" & left(jsonString,len(jsonString)-1) & "]")



				end if 





	end if
	

%>
