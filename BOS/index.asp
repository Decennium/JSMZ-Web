<!--#include file="Config.asp" -->
<html>
<meta name="viewport" content="width=device-width,minimum-scale=1.0, maximum-scale=2.0"/>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<link href="../css/css.css" rel="stylesheet">
<head>
<title>科技馆日常工作管理系统</title>
</head>
<body dir="ltr">
<!--#include file="../include/banner.asp"-->
<div style="float:left;">
<!--#include file="../include/left_banner.asp"-->
</div>
<%
Select Case Action
'登陆后台调用
Case "logincheck"
	Admin_User=htmlencode(Request.form("Admin_User"))
	Admin_Pass=md5(Request.form("Admin_Pass"))	
	Set mRs=conn.execute("select * from [Admin] where Name='"&Admin_User&"' and Pass='"&Admin_Pass&"'")
	If not mRs.eof then
		Session("Admin")=mRs("Name")
		Session("ShowName")=mRs("ShowName")
		Session("ComputerLab")=mRs("ComputerLab")
		Session("Rights")=mRs("Rights")
		Response.Redirect Url
		'Response.End
	Else
		Response.Write "<script>document.getElementById('Tips').innerHTML = '用户名或密码错误，请重试。';</SCRIPT>"
		'Response.End
	End If

'退出后台调用
Case "logout"
	Session.Contents.RemoveAll
	Session.Abandon
	Response.Redirect Url
	Response.End
End Select
%>
<div id="Right_Content">
<div class="HeadLine">科技馆日常工作管理系统</div>
<div class="ShowTips">
	<p>您可以随时浏览科技馆特别是信息技术组的工作情况。</p>
	<h4>但是，如果您想管理或添加信息技术组的工作情况，您需要先登录。</h4>
	<p>登录帐户不能通过在线注册的方式获得。</p>
	<p>只有信息技术组的老师才能手工建立账户。</p>
	<p>如果您需要账户，请联系信息技术组，为您手工建立账户。</p>
	<p>您可以在左侧登录。</p>
</div>
</div>
<br clear="all"><br><br><br><br><br><br><br><br><br>
<!--#include file="../include/bottom.asp"-->
</body>
</html> 