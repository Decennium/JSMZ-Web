<div name="Left_Banner" id="Left_Banner" style="width:150px;max-width:150px;clear:left;float:left;margin-right:10px">
<div style="width:150px;max-width:150px;clear:left;float:left">
	<a href="http://192.168.2.1/">
	<img src="../images/jsmz_logo_mini.png" alt="JSMZ Logo" border="0" align="left">
	</a>
</div>
<br clear="all">
<br clear="all">
<div align="left" style="width:150px;max-width:150px;clear:left;float:left">
<%
If instr(Request.ServerVariables("SCRIPT_NAME"),"BOS")>0 Then
	If Len(Session("Admin")) > 0 Then
%>
<p><a href="/BOS/index.asp?action=logout">登出 <%=Session("ShowName")%></a></p>
<%Else%>
<p><a href="/BOS/index.asp?action=login">登录</a></p>
<%
End If
End If
%>
<p><a href="manage.asp">用户管理系统</a></p>
<p><a href="jifang.asp">机房使用管理系统</a></p>
</div>
</div>