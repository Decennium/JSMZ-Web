<!--#include file="Config.asp" -->
<html>
<meta name="viewport" content="width=device-width,minimum-scale=1.0, maximum-scale=2.0"/>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<link href="../css/css.css" rel="stylesheet">
<head>
<title>�Ƽ����ճ���������ϵͳ</title>
</head>
<body dir="ltr">
<!--#include file="../include/banner.asp"-->
<div style="float:left;">
<!--#include file="../include/left_banner.asp"-->
</div>
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
		Session("ComputerLab")=mRs("ComputerLab")
		Session("Rights")=mRs("Rights")
		Response.Redirect Url
		'Response.End
	Else
		Response.Write "<script>document.getElementById('Tips').innerHTML = '�û�����������������ԡ�';</SCRIPT>"
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
<div id="Right_Content">
<div class="HeadLine">�Ƽ����ճ���������ϵͳ</div>
<div class="ShowTips">
	<p>��������ʱ����Ƽ����ر�����Ϣ������Ĺ��������</p>
	<h4>���ǣ�����������������Ϣ������Ĺ������������Ҫ�ȵ�¼��</h4>
	<p>��¼�ʻ�����ͨ������ע��ķ�ʽ��á�</p>
	<p>ֻ����Ϣ���������ʦ�����ֹ������˻���</p>
	<p>�������Ҫ�˻�������ϵ��Ϣ�����飬Ϊ���ֹ������˻���</p>
	<p>������������¼��</p>
</div>
</div>
<br clear="all"><br><br><br><br><br><br><br><br><br>
<!--#include file="../include/bottom.asp"-->
</body>
</html> 