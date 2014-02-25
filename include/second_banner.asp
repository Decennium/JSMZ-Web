<div name="Second_Banner" id="Second_Banner" class="underline" width="100%" style="text-align:left;">

<%
'If instr(Request.ServerVariables("SCRIPT_NAME"),"BOS")>0 Then
If Len(Session("Admin")) > 0 Then
%>
<a href="/BOS/index.asp?action=logout">登出 <%=Session("ShowName")%></a>
<a href="/BOS/manage.asp">用户管理</a>
<%Else%>
	<form method="post" Action="/BOS/index.asp?Action=logincheck">
<%If Len(Session("Admin")) = 0 Then%>
	<span id="form-noindent" align="right" style="visibility:visible">
<%Else%>
	<span id="form-noindent" align="right" style="visibility:hidden">
<%End If%>
		<input type="text" name="Admin_User" value="" id="Admin_User" size="10" title="帐号">
		<input type="password" name="Admin_Pass" id="Admin_Pass" size="10" title="密码">
		<input type="submit" name="null" value="登录">
	</span>
	</form>
<%
End If
'End If
%>
||
<a href="/BOS/gouru.asp">设备购入</a>
<a href="/BOS/waijie.asp">设备外借</a>
<a href="/BOS/weixiu.asp">设备维修</a>
<a href="/BOS/luyinji.asp">领取录音机</a>
<a href="/BOS/jiankong.asp">监控设备</a>
<a href="/BOS/jifang.asp">机房使用</a>
<a href="/BOS/MMC.asp">多媒体教室</a>
<a href="/BOS/shiyong.asp">设备使用</a>
<a href="/BOS/eReading.asp">电子阅读</a>
||
<a href="/BOS/tushushi.asp">图书室开放</a>
<a href="/tz/index.asp">学校通知</a>
||
<span id="Tips" width="100%" align="center">小提示</span>

</div>