<div name="Second_Banner" id="Second_Banner" class="underline" width="100%" style="text-align:left;">

<%
'If instr(Request.ServerVariables("SCRIPT_NAME"),"BOS")>0 Then
If Len(Session("Admin")) > 0 Then
%>
<a href="/BOS/index.asp?action=logout">�ǳ� <%=Session("ShowName")%></a>
<a href="/BOS/manage.asp">�û�����</a>
<%Else%>
	<form method="post" Action="/BOS/index.asp?Action=logincheck">
<%If Len(Session("Admin")) = 0 Then%>
	<span id="form-noindent" align="right" style="visibility:visible">
<%Else%>
	<span id="form-noindent" align="right" style="visibility:hidden">
<%End If%>
		<input type="text" name="Admin_User" value="" id="Admin_User" size="10" title="�ʺ�">
		<input type="password" name="Admin_Pass" id="Admin_Pass" size="10" title="����">
		<input type="submit" name="null" value="��¼">
	</span>
	</form>
<%
End If
'End If
%>
||
<a href="/BOS/gouru.asp">�豸����</a>
<a href="/BOS/waijie.asp">�豸���</a>
<a href="/BOS/weixiu.asp">�豸ά��</a>
<a href="/BOS/luyinji.asp">��ȡ¼����</a>
<a href="/BOS/jiankong.asp">����豸</a>
<a href="/BOS/jifang.asp">����ʹ��</a>
<a href="/BOS/MMC.asp">��ý�����</a>
<a href="/BOS/shiyong.asp">�豸ʹ��</a>
<a href="/BOS/eReading.asp">�����Ķ�</a>
||
<a href="/BOS/tushushi.asp">ͼ���ҿ���</a>
<a href="/tz/index.asp">ѧУ֪ͨ</a>
||
<span id="Tips" width="100%" align="center">С��ʾ</span>

</div>