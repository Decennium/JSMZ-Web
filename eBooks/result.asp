<%
Response.Charset = "gb2312"
Response.Buffer = True

select case request("stype")
case "slib"
	TargetPage = "lib.asp"
case "ebooks"
	TargetPage = "ebooks.asp"
case "codes"
	TargetPage = "codes.asp"
case "zheda"
	TargetPage = "zheda.asp"
case else
	TargetPage = "index.asp"
end select
if trim(request("q"))="" then TargetPage = "index.asp"
Server.Transfer TargetPage
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>图书查询系统</title>
<link href="../css/css.css" rel="stylesheet">
<link rel="shortcut icon" type="image/x-icon" href="favicon.ico">
</head>
<body>
</body>
</html>