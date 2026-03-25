<!--#include file="connSmartFarming.inc" -->

<%
	Dim FSOobj,FilePath1

	FilePath1=Server.MapPath("Foto/"& Request.QueryString("kdid")&"_1.JPEG") ' located in the same director

	Set FSOobj = Server.CreateObject("Scripting.FileSystemObject")

	if (FSOobj.fileExists(FilePath1))   Then

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

