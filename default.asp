<%@ Language="VBScript" %>
<!DOCTYPE html>
<html>
<head>
    <title>My Landing Page</title>
    <meta charset="UTF-8">
    <style>
        body {
            font-family: Arial;
            text-align: center;
            background-color: #f4f4f4;
        }
        .container {
            margin-top: 100px;
        }
        .btn {
            padding: 10px 20px;
            background: blue;
            color: white;
            text-decoration: none;
            border-radius: 5px;
        }
    </style>
</head>
<body>

<div class="container">
    <h1>Welcome to My Website</h1>

    <%
        Dim userName
        userName = Request.QueryString("name")

        If userName <> "" Then
            Response.Write("<p>Hello, " & userName & "!</p>")
        Else
            Response.Write("<p>Hello, Guest!</p>")
        End If
    %>

</div>

</body>
</html>