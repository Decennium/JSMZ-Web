<div name="Left_Banner" id="Left_Banner" style="width:150px;max-width:150px;clear:left;float:left;margin-right:10px">
<div style="width:150px;max-width:150px;clear:left;float:left">
	<a href="http://192.168.2.1/">
	<img src="../images/jsmz_logo_mini.png" alt="JSMZ Logo" border="0" align="left">
	</a>
</div>
<br clear="all">
<div align="left" style="width:150px;max-width:150px;clear:left;float:left">
<%
If instr(Request.ServerVariables("SCRIPT_NAME"),"BOS")>0 Then
	If Len(Session("Admin")) > 0 Then
%>
<p><a href="/BOS/index.asp?action=logout">登出 <%=Session("ShowName")%></a></p>
<%Else%>
<div id="LoginForm">
	<form method="post" Action="/BOS/index.asp?Action=logincheck">
<%If Len(Session("Admin")) = 0 Then%>
	<table id="form-noindent" align="right" style="visibility:visible">
<%Else%>
	<table id="form-noindent" align="right" style="visibility:hidden">
<%End If%>
		<tr>
			<td>
			<div align="center">
			帐号：<input type="text" name="Admin_User" value="" id="Admin_User" size="10">
			</div>
			<div align="center">
			密码：<input type="password" name="Admin_Pass" id="Admin_Pass" size="10">
			</div>
			<input type="submit" name="null" value="登录">
			</td>
		</tr>
	</table>
	</form>
</div>
<%
End If
End If
%>
<p><a href="manage.asp">用户管理</a></p>
<p><a href="jifang.asp">机房使用管理</a></p>
<p><a href="shiyong.asp">设备使用管理</a></p>
<p><a href="gouru.asp">设备购入管理</a></p>
<p><a href="luyinji.asp">领取录音机管理</a></p>
<p><a href="MMC.asp">多媒体教室使用管理</a></p>
<p><a href="jiankong.asp">监控设备使用管理</a></p>
<p><a href="weixiu.asp">设备维修管理</a></p>
<hr>
<p><a href="index.asp">设备外借管理</a></p>
</div>
<br clear="all">
<hr>
<div id="Tips" width="100%" align="center">小提示</div>

</div>