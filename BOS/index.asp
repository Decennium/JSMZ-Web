<!--#include file="Config.asp" -->
<html>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<link href="../css/css.css" rel="stylesheet">
<head>
<title>科技馆日常工作管理系统</title>
</head>
<body dir="ltr">
<!--#include file="../include/banner.asp"-->
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
		Response.Redirect Url
		'Response.End
	Else
		Response.Write "<script>alert('非法操作：用户名或密码错误！');this.location.href='?Action=login';</SCRIPT>"
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
<div style="float:left;">
<!--#include file="left_banner.asp"-->
</div>
<div style="width="100%";float:left">
<div class="HeadLine">科技馆日常工作管理系统</div>
<div class="ShowTips">
	<h4>登录账户，可以管理用户资料以及机房使用资料。</h4>
	<p>这些账户不能通过在线注册的方式获得。</p>
	<p>只有信息技术组的老师才能手工建立账户。</p>
	<p>如果您需要账户，请联系信息技术组，为您手工建立账户。</p>
	<p>您可以在右侧登录。</p>
</div>
<div id="LoginForm">
	<form method="post" Action="?Action=logincheck">
<%If Len(Session("Admin")) = 0 Then%>
	<table id="form-noindent" align="right" style="visibility:visible">
<%Else%>
	<table id="form-noindent" align="right" style="visibility:hidden">
<%End If%>
		<tr>
			<td bgcolor="#e8eefa">
			<h4>帐户登录</h4>
			<div align="center">
			帐号：<input type="text" name="Admin_User" value="" id="Admin_User" size="15">
			</div>
			<div align="center">
			密码：<input type="password" name="Admin_Pass" id="Admin_Pass" size="15">
			</div>
			<input type="submit" name="null" value="登录">
			</td>
		</tr>
	</table>
	</form>
</div>
</div>
<br clear="all"><br><br><br><br><br><br><br><br><br>
<!--#include file="../include/bottom.asp"-->
</body>
</html> 