<!--#include file="Config.asp" -->
<html>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<link href="../css/css.css" rel="stylesheet">
<head>
<title>�Ƽ����ճ���������ϵͳ</title>
</head>
<body dir="ltr">
<!--#include file="../include/banner.asp"-->
<%
Select Case Action
'��½��̨����
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
		Response.Write "<script>alert('�Ƿ��������û������������');this.location.href='?Action=login';</SCRIPT>"
		'Response.End
	End If

'�˳���̨����
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
<div class="HeadLine">�Ƽ����ճ���������ϵͳ</div>
<div class="ShowTips">
	<h4>��¼�˻������Թ����û������Լ�����ʹ�����ϡ�</h4>
	<p>��Щ�˻�����ͨ������ע��ķ�ʽ��á�</p>
	<p>ֻ����Ϣ���������ʦ�����ֹ������˻���</p>
	<p>�������Ҫ�˻�������ϵ��Ϣ�����飬Ϊ���ֹ������˻���</p>
	<p>���������Ҳ��¼��</p>
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
			<h4>�ʻ���¼</h4>
			<div align="center">
			�ʺţ�<input type="text" name="Admin_User" value="" id="Admin_User" size="15">
			</div>
			<div align="center">
			���룺<input type="password" name="Admin_Pass" id="Admin_Pass" size="15">
			</div>
			<input type="submit" name="null" value="��¼">
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