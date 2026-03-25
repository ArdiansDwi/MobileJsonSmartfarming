<!--#include file="connSmartFarming.inc" -->

<%
	Dim FSOobj,FilePath1, FilePath2, FilePath3, FilePath4

	FilePath1=Server.MapPath("Foto/"& Request.QueryString("kdid")&"_1.JPEG") ' located in the same director
FilePath2=Server.MapPath("Foto/"& Request.QueryString("kdid")&"_2.JPEG")
FilePath3=Server.MapPath("Foto/"& Request.QueryString("kdid")&"_3.JPEG")
FilePath4=Server.MapPath("Foto/"& Request.QueryString("kdid")&"_4.JPEG")

	Set FSOobj = Server.CreateObject("Scripting.FileSystemObject")

	if (FSOobj.fileExists(FilePath1)) AND (FSOobj.fileExists(FilePath2)) AND (FSOobj.fileExists(FilePath3)) AND (FSOobj.fileExists(FilePath4))  Then

		%>
		[{"Result" : "Success"}]
		<%
	Else

		%>
		[{"Result" : "Failed"}]
		<%
	
	End if
	
	Set FSOobj = Nothing
%> 

