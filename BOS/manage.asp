<!--#include file="Config.asp" -->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>�û�����ϵͳ</title>
<link href="../css/css.css" rel="stylesheet">
<link rel="shortcut icon" type="image/x-icon" href="favicon.ico">
</head>
<body>
<!--#include file="../include/banner.asp"-->
<!--#include file="left_banner.asp"-->
<script language="javascript">
//**********��ӹ���Ա����**********
function Addadmin(the){
	//�жϹ���Ա���Ʋ���Ϊ��
	if(the.Name.value==""){
		document.getElementById('Tips').innerHTML = '���Ʋ���Ϊ��';
		the.Name.focus();
		return false;
	}
	//�жϹ���Ա��ʾ���Ʋ���Ϊ��
	if(the.ShowName.value==""){
		document.getElementById('Tips').innerHTML = '��ʾ���Ʋ���Ϊ��';
		the.ShowName.focus();
		return false;
	}
	//�жϹ���Ա���벻��С��6���ַ�
	if(the.Pass.value.length<6){
		document.getElementById('Tips').innerHTML = '����Ա���벻������6���ַ���';
		the.Pass.focus();
		return false;
	}
	//�жϹ���Ա���벻�ô���16���ַ�
	if(the.Pass.value.length>16){
		document.getElementById('Tips').innerHTML = '����Ա���벻�ö���16���ַ���';
		the.Pass.focus();
		return false;
	}
	//�жϹ���Ա����������������
	if(the.Pass.value!=the.Password.value){
		document.getElementById('Tips').innerHTML = '�������벻һ�£�';
		the.Password.focus();
		return false;
	}
}

//**********�޸Ĺ���Ա����**********
function AdminModpass(the){
	//�жϹ���Ա���Ʋ���Ϊ��
	if(the.Name.value==""){
		document.getElementById('Tips').innerHTML = '���Ʋ���Ϊ��';
		the.Name.focus();
		return false;
	}
	//�жϹ���Ա��ʾ���Ʋ���Ϊ��
	if(the.ShowName.value==""){
		document.getElementById('Tips').innerHTML = '��ʾ���Ʋ���Ϊ��';
		the.ShowName.focus();
		return false;
	}
	//�жϹ���Ա�����벻��Ϊ��
	if(the.Admin_Gps.value==""){
		document.getElementById('Tips').innerHTML = '����Ա�����벻��Ϊ�գ�';
		the.Admin_Gps.focus();
		return false;
	}
	//�жϹ���Ա�����벻��С��6���ַ�
	if(the.Admin_Nps.value.length<6){
		document.getElementById('Tips').innerHTML = '����Ա�����벻��С��6���ַ���';
		the.Admin_Nps.focus();
		return false;
	}
	//�жϹ���Ա���벻�ô���16���ַ�
	if(the.Admin_Nps.value.length>16){
		document.getElementById('Tips').innerHTML = '����Ա���벻�ö���16���ַ���';
		the.Admin_Nps.focus();
		return false;
	}
	//�жϹ���Ա����������������
	if(the.Admin_Nps.value!=the.Pass.value){
		document.getElementById('Tips').innerHTML = '���������벻һ�£�';
		the.Pass.focus();
		return false;
	}
}

//**********ɾ������**********
function Deladmin()
{
  if(!confirm('ȷ��ɾ������Ա��')) return false;
}

</script>
<%
'����Ա�б�ҳ��
checkadmin

Set mRs=Server.CreateObject("adodb.recordSet")
Sql="Select Id,Name,ShowName from Admin"
mRs.open Sql,conn,1,1
%>
<div id="Right_Content" width='100%' style="float:left">
<table border="0" width="100%" style="margin-top:0px">
<thead>
	<tr>
		<th colspan="6">����Ա�б�</th>
	</tr>
</thead>
<tbody>
	<tr align="center">
		<td width="5%">ID</td>
		<td width="25%">�û���</td>
		<td width="25%">��ʾ��</td>
		<td width="25%">����</td>
		<td width="20%" colspan="2">����</td>
	</tr>
	<% do while not mRs.eof %>
	<tr align="center" id='Data'>
		<td><% =mRs("Id") %></td>
		<td><% =mRs("Name") %></td>
		<td><% =mRs("ShowName") %></td>
		<td>******</td>
		<td><a href="?Action=Admin_Modpass&Id=<% =mRs("Id") %>&Name=<% =mRs("Name") %>&ShowName=<% =mRs("ShowName") %>">�޸�</a></td>
		<td>
		<%if "'"& mRs("Name") &"'"<>"'"& Session("Admin") &"'" And mRs("Id")<>1 Then
		Response.Write"<a href=""?Action=Deladmin&Id="& mRs("Id") &""" onclick=""return Del(this);"">ɾ��</a>"
		End If
		%>
		</td>
	</tr>
	<%
	mRs.movenext
	loop
	mRs.close
	Set mRs=nothing
	%>
</tbody>
</table>
<p style="margin-bottom:10px;"><a href='?Action=Admin_Add'>��������Ա</a></p>
<%
Select Case action
'��ӹ���Ա
Case "Addadmin"
	checkadmin

dim Name,Pass
Name=htmlencode(Request.form("Name"))
ShowName=htmlencode(Request.form("ShowName"))
Pass=Request.form("Pass")
	If Name="" or Pass="" Or ShowName = "" then
		Response.Write "<script>document.getElementById('Tips').innerHTML = '�ֶβ���Ϊ��';</SCRIPT>"
	ElseIf Request("Pass")<>Request("Password") then
		Response.Write "<script>document.getElementById('Tips').innerHTML = '������֤����ȷ';</SCRIPT>"
	ElseIf len(Pass)<6 or len(Pass)>16 then
		Response.Write "<script>document.getElementById('Tips').innerHTML = '���볤��̫�̻�̫��';</SCRIPT>"
	Else
	Sql="Insert Into [Admin] (Name,Pass,ShowName) values ('"& Name &"','"& md5(Pass) &"','" & ShowName &"')"
	conn.execute(Sql)
		Response.Redirect "?Action="
		Response.End
	End If

'�޸Ĺ���Ա�������
Case "AdminModpass"
	checkadmin

	Name=htmlencode(Request.form("Name"))
	ShowName=htmlencode(Request.form("ShowName"))
	Admin_Gps=md5(Request.form("Admin_Gps"))
	Admin_Nps=md5(Request.form("Admin_Nps"))
	Pass=md5(Request.form("Pass"))

	Set mRs=conn.execute("select * from [Admin] where Id="& Id &" and Pass='"& Admin_Gps &"'")
	If mRs.eof then
		Response.Write "<script>this.location.href='?Action=Admin_Modpass';</SCRIPT>"
		Response.End
	End If

	Sql="update [Admin] Set Name='"& Name &"',Pass='"& Pass &"' where Id="& Id &""
	conn.execute(Sql)
	Response.Redirect "?Action="
	Response.End

'ɾ������Ա
Case "Deladmin"
	checkadmin

	Sql="delete * from Admin where Id="& Id
	conn.execute(Sql)
	Response.Redirect "?Action="
	Response.End

'�޸�����ҳ��
Case "Admin_Modpass"
	checkadmin
%>
<p style="margin-bottom:10px;">����Ա�ʺ��޸�</p>
<div id="ChangeAccount" name="ChangeAccount">
	<form method="post" Action="?Action=AdminModpass&Id=<% =id %>" onSubmit="return AdminModpass(this);">
		<label for="Name">��¼����</label>
		<input name="Name" type="text" value="<% =Request.Querystring("Name") %>"/>
		<label for="ShowName">��ʾ����</label>
		<input name="ShowName" type="text" value="<% =Request.Querystring("ShowName") %>"/>
		<label for="Admin_Gps">�����룺</label>
		<input name="Admin_Gps" type="password" maxlength="16" value=""/>
		<label for="Admin_Nps">�����룺</label>
		<input name="Admin_Nps" type="password" maxlength="16" value=""/>
		<label for="Pass">ȷ�����룺</label>
		<input name="Pass" type="password" maxlength="16"/>
		<input type="submit" value="�޸�" class="bmit"/>
	</form>
</div>
<%
'��������Աҳ��
Case "Admin_Add"
	checkadmin
%>
<div align="left">
	<form method="post" Action="?Action=Addadmin" onSubmit="return Addadmin(this);">
		<label for="Name">�û�����</label><input name="Name" type="text" class="input" value=""/>
		<label for="ShowName">��ʾ����</label><input name="ShowName" type="text" class="input" value=""/>
		<label for="Pass">���룺</label><input name="Pass" type="Password" class="input" value=""/>
		<label for="Password">ȷ�����룺</label><input name="Password" type="Password" class="input" value=""/>
		<input name="Submit3" type="submit" value="���"/>
	</form>
</div>
<%
'call PageControl(iCount,maxpage,page)

End Select
conn.close
Set conn=nothing
%>
</div>
<br clear="all">
<!--#include file="../include/bottom.asp"-->
<script language="javascript">
/*����ȡ��ֵ�浽������*/
c_Width =document.body.offsetWidth - parseInt(document.getElementById('Left_Banner').style.width)-40;
document.getElementById('Right_Content').style.width="" +c_Width +'px';
document.getElementById('Right_Content').style.maxWidth="" + c_Width  + 'px';

window.onresize=function(){
    document.getElementById("Right_Content").style.width=(function(){
        var x=document.body.offsetWidth - parseInt(document.getElementById('Left_Banner').style.width)-40;
        return x;
    })();
}
</script>
</body>
</html>
