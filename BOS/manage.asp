<!--#include file="Config.asp" -->
<html>
<head>
<meta name="viewport" content="width=device-width,minimum-scale=1.0, maximum-scale=2.0"/>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>用户管理系统</title>
<link href="../css/css.css" rel="stylesheet">
<link rel="shortcut icon" type="image/x-icon" href="favicon.ico">
</head>
<body>
<!--#include file="../include/banner.asp"-->
<!--#include file="../include/second_banner.asp"-->
<script language="javascript">
//**********添加管理员调用**********
function Addadmin(the){
	//判断管理员名称不能为空
	if(the.Name.value==""){
		document.getElementById('Tips').innerHTML = '名称不能为空';
		the.Name.focus();
		return false;
	}
	//判断管理员显示名称不能为空
	if(the.ShowName.value==""){
		document.getElementById('Tips').innerHTML = '显示名称不能为空';
		the.ShowName.focus();
		return false;
	}
	//判断管理员密码不得小于6个字符
	if(the.Pass.value.length>0 && the.Pass.value.length<6){
		document.getElementById('Tips').innerHTML = '管理员密码不得少于6个字符！';
		the.Pass.focus();
		return false;
	}
	//判断管理员密码不得大于16个字符
	if(the.Pass.value.length>16){
		document.getElementById('Tips').innerHTML = '管理员密码不得多于16个字符！';
		the.Pass.focus();
		return false;
	}
	//判断管理员两次新密码必须相等
	if(the.Pass.value!=the.Password.value){
		document.getElementById('Tips').innerHTML = '两次密码不一致！';
		the.Password.focus();
		return false;
	}
}

//**********修改管理员调用**********
function AdminModpass(the){
	//判断管理员名称不能为空
	if(the.Name.value==""){
		document.getElementById('Tips').innerHTML = '名称不能为空';
		the.Name.focus();
		return false;
	}
	//判断管理员显示名称不能为空
	if(the.ShowName.value==""){
		document.getElementById('Tips').innerHTML = '显示名称不能为空';
		the.ShowName.focus();
		return false;
	}
	if(the.Admin_Gps.value!="" && the.Admin_Nps.value!="" && the.Pass.value!="") {
		//判断管理员旧密码不能为空
		if(the.Admin_Gps.value==""){
			document.getElementById('Tips').innerHTML = '管理员旧密码不能为空！';
			the.Admin_Gps.focus();
			return false;
		}
		//判断管理员新密码不得小于6个字符
		if(the.Admin_Nps.value.length<6){
			document.getElementById('Tips').innerHTML = '管理员新密码不得小于6个字符！';
			the.Admin_Nps.focus();
			return false;
		}
		//判断管理员密码不得大于16个字符
		if(the.Admin_Nps.value.length>16){
			document.getElementById('Tips').innerHTML = '管理员密码不得多于16个字符！';
			the.Admin_Nps.focus();
			return false;
		}
		//判断管理员两次新密码必须相等
		if(the.Admin_Nps.value!=the.Pass.value){
			document.getElementById('Tips').innerHTML = '两次新密码不一致！';
			the.Pass.focus();
			return false;
		}
	}
	return true;
}

//**********删除调用**********
function Deladmin()
{
  if(!confirm('确认删除管理员吗？')) return false;
}

</script>
<%
'管理员列表页面
checkadmin

Set mRs=Server.CreateObject("adodb.recordSet")
Sql="Select Id,Name,ShowName,ComputerLab from Admin"
mRs.open Sql,conn,1,1
%>
<div id="Content" width='100%' style="float:left">
<table border="0" width="100%" style="margin-top:0px">
<thead>
	<tr>
		<th colspan="7">管理员列表</th>
	</tr>
</thead>
<tbody>
	<tr align="center">
		<td width="5%">ID</td>
		<td width="25%">用户名</td>
		<td width="25%">显示名</td>
		<td>机房</td>
		<td>密码</td>
		<td colspan="2">操作</td>
	</tr>
	<% do while not mRs.eof %>
	<tr align="center" id='Data'>
		<td><% =mRs("Id") %></td>
		<td><% =mRs("Name") %></td>
		<td><% =mRs("ShowName") %></td>
		<td><% =mRs("ComputerLab") %></td>
		<td>******</td>
		<td><a href="?Action=Admin_Modpass&Id=<% =mRs("Id") %>&Name=<% =mRs("Name") %>&ShowName=<% =mRs("ShowName") %>&ComputerLab=<% =mRs("ComputerLab") %>">修改</a></td>
		<td>
		<%if "'"& mRs("Name") &"'"<>"'"& Session("Admin") &"'" And mRs("Id")<>1 Then
		Response.Write"<a href=""?Action=Deladmin&Id="& mRs("Id") &""" onclick=""return Del(this);"">删除</a>"
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
<p style="margin-bottom:10px;"><a href='?Action=Admin_Add'>新增管理员</a></p>
<%
Select Case action
'添加管理员
Case "Addadmin"
	checkadmin

dim Name,Pass,ComputerLab,ShowName
Name=htmlencode(Request.form("Name"))
ShowName=htmlencode(Request.form("ShowName"))
ComputerLab=htmlencode(Request.form("ComputerLab"))
Pass=Request.form("Pass")
	If Name="" Or ShowName = "" then
		Response.Write "<script>document.getElementById('Tips').innerHTML = '字段不能为空';</SCRIPT>"
	ElseIf Request("Pass")<>Request("Password") then
		Response.Write "<script>document.getElementById('Tips').innerHTML = '密码验证不正确';</SCRIPT>"
	ElseIf len(Pass)>0 And (len(Pass)<6 or len(Pass)>16) then
		Response.Write "<script>document.getElementById('Tips').innerHTML = '密码长度太短或太长';</SCRIPT>"
	Else
		Set mRs=conn.execute("select * from [Admin] where Name='"& Name & "'")
		If not mRs.eof then
			Response.Write "<script>document.getElementById('Tips').innerHTML = '用户名重复';</SCRIPT>"
			Response.End
		Else
			Sql="Insert Into [Admin] (Name,Pass,ShowName,ComputerLab) values ('"& Name &"','"& md5(Pass) &"','" & ShowName & "','" & ComputerLab & "')"
			conn.execute(Sql)
			Response.Redirect "?Action="
			Response.End
		End If
	End If

'修改管理员密码调用
Case "AdminModpass"
	checkadmin

	Id=htmlencode(Request.form("Id"))
	Name=htmlencode(Request.form("Name"))
	ShowName=htmlencode(Request.form("ShowName"))
	ComputerLab=htmlencode(Request.form("ComputerLab"))
	Admin_Gps=md5(Request.form("Admin_Gps"))
	Admin_Nps=md5(Request.form("Admin_Nps"))
	Pass=md5(Request.form("Pass"))

	Set mRs=conn.execute("select * from [Admin] where Id="& Id)
	If mRs.eof then
		Response.Write "<script>this.location.href='?Action=Admin_Modpass';</SCRIPT>"
		Response.End
	End If
	If Len(Request.form("Pass")) >= 6 Then
		Sql="update [Admin] Set ShowName='" & ShowName & "',Pass='"& Pass &"' ,ComputerLab='" & ComputerLab & "' where Id="& Id &""
	Else
		Sql="update [Admin] Set ShowName='" & ShowName & "',ComputerLab='" & ComputerLab & "' where Id="& Id &""
	End If
	'Response.write sql
	conn.execute(Sql)
	Response.Redirect "?Action="
	Response.End

'删除管理员
Case "Deladmin"
	checkadmin

	Sql="delete * from Admin where Id="& Id
	conn.execute(Sql)
	Response.Redirect "?Action="
	Response.End

'修改密码页面
Case "Admin_Modpass"
	checkadmin
%>
<p style="margin-bottom:10px;">管理员帐号修改</p>
<div id="ChangeAccount" name="ChangeAccount">
	<form method="post" Action="?Action=AdminModpass" onSubmit="return AdminModpass(this);">
		<input type="hidden" name="Id" value="<% =id %>" />
		<label for="Name">登录名：</label>
		<input name="Name" type="text" value="<% =Request.Querystring("Name") %>" readonly size="10"/>
		<label for="ShowName">显示名：</label>
		<input name="ShowName" type="text" value="<% =Request.Querystring("ShowName") %>" size="10"/>
		<label for="ComputerLab">机房：</label>
		<select name="ComputerLab" id="ComputerLab">
			<option value="1机房"<%If Request.Querystring("ComputerLab")="1机房" Then Response.write " Selected"%>>1机房</option>
			<option value="2机房"<%If Request.Querystring("ComputerLab")="2机房" Then Response.write " Selected"%>>2机房</option>
			<option value="3机房"<%If Request.Querystring("ComputerLab")="3机房" Then Response.write " Selected"%>>3机房</option>
			<option value="4机房"<%If Request.Querystring("ComputerLab")="4机房" Then Response.write " Selected"%>>4机房</option>
			<option value="教室"<%If Request.Querystring("ComputerLab")="教室" Then Response.write " Selected"%>>教室</option>
			<option value="无机房"<%If Request.Querystring("ComputerLab")="无机房" Then Response.write " Selected"%>>无机房</option>
		</select>
		<label for="Admin_Gps">旧密码：</label>
		<input name="Admin_Gps" type="password" maxlength="16" value=""/>
		<label for="Admin_Nps">新密码：</label>
		<input name="Admin_Nps" type="password" maxlength="16" value=""/>
		<label for="Pass">确认密码：</label>
		<input name="Pass" type="password" maxlength="16"/>
		<input type="submit" value="修改" class="bmit"/>
	</form>
</div>
<%
'新增管理员页面
Case "Admin_Add"
	checkadmin
%>
<div align="left">
	<form method="post" Action="?Action=Addadmin" onSubmit="return Addadmin(this);">
		<label for="Name">用户名：</label><input name="Name" type="text" class="input" value=""/>
		<label for="ShowName">显示名：</label><input name="ShowName" type="text" class="input" value=""/>
		<label for="ComputerLab">机房：</label>
		<select name="ComputerLab" id="ComputerLab">
			<option value="1机房">1机房</option>
			<option value="2机房">2机房</option>
			<option value="3机房">3机房</option>
			<option value="4机房">4机房</option>
			<option value="教室">教室</option>
			<option value="无机房" Selected="Selected">无机房</option>
		</select>
		<label for="Pass">密码：</label><input name="Pass" type="Password" class="input" value=""/>
		<label for="Password">确认密码：</label><input name="Password" type="Password" class="input" value=""/>
		<input name="Submit3" type="submit" value="添加"/>
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
/*将获取的值存到变量里*/
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
