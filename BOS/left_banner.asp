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
<p><a href="/BOS/index.asp?action=logout">�ǳ� <%=Session("ShowName")%></a></p>
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
			�ʺţ�<input type="text" name="Admin_User" value="" id="Admin_User" size="10">
			</div>
			<div align="center">
			���룺<input type="password" name="Admin_Pass" id="Admin_Pass" size="10">
			</div>
			<input type="submit" name="null" value="��¼">
			</td>
		</tr>
	</table>
	</form>
</div>
<%
End If
End If
%>
<p><a href="manage.asp">�û�����</a></p>
<p><a href="jifang.asp">����ʹ�ù���</a></p>
<p><a href="shiyong.asp">�豸ʹ�ù���</a></p>
<p><a href="gouru.asp">�豸�������</a></p>
<p><a href="luyinji.asp">��ȡ¼��������</a></p>
<p><a href="MMC.asp">��ý�����ʹ�ù���</a></p>
<p><a href="jiankong.asp">����豸ʹ�ù���</a></p>
<p><a href="weixiu.asp">�豸ά�޹���</a></p>
<hr>
<p><a href="index.asp">�豸������</a></p>
</div>
<br clear="all">
<hr>
<div id="Tips" width="100%" align="center">С��ʾ</div>

</div>