<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta name="viewport" content="width=device-width,minimum-scale=1.0, maximum-scale=2.0"/>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<!--<meta http-equiv="content-type" content="text/html; charset=UTF-8"> -->
<link href="../css/css.css" rel="stylesheet">
<link rel="shortcut icon" type="image/x-icon" href="favicon.ico">
<%
EndTime=Timer()

If request("q")="" Then
	response.write "<title>���е���ͼ��</title>"
Else
	response.write "<title>" & request("q") & " - �������</title>"
End If
%>
<title><%=request("q")%> - �������</title>
<script type="text/javascript">
function setFocus()
{
document.getElementById("q").focus();
}
</script>
</head>
<body bgcolor=#ffffff onload="setFocus();" topmargin=3 marginheight=3>
<noscript>
</noscript>
<!--#include file="../include/banner.asp"-->
<form name=gs method="POST" action="result.asp">
<div align="left" style="clear:none; float:left; padding-top : 5px;">
<a id=logo href="." title="��ҳ"><img src="images/jsmz_logo_mini.png" border="0px"></a>
</div>
<div align="left" style="clear:right; float:left; ">
<div>
<span style="white-space: nowrap"><input type="radio" id="ebooks" name="stype" value="ebooks" <%if request("stype")="ebooks" then response.write "checked=true"%>><label for="ebooks">����ͼ��</label></span>
<span style="white-space: nowrap"><input type="radio" id="slib" name="stype" value="slib" <%if request("stype")="slib" then response.write "checked=true"%>><label for="slib">ѧУͼ����</label></span>
<span style="white-space: nowrap"><input type="radio" id="res" name="stype" value="res" <%if request("stype")="res" then response.write "checked=true"%>><label for="res">ѧУ��Դ��</label></span>
<span style="white-space: nowrap"><input type="radio" id="scodes" name="stype" value="codes" <%if request("stype")="codes" then response.write "checked=true"%>><label for="scodes">��Ͽ�����</label></span>
</div>
<div>
<input type="hidden" name="page" value="1">
<input  maxlength="50%" id="q" name="q" size="30%" title="������������ؼ��ʣ��ÿո����" value="<%=request("q")%>" onfocus="this.select()" onmouseover="this.select()">
<input name="btnS" type="submit" value="����">
</div>
</div>
</form>
<div class="HeadLine" style="clear:both;">
&nbsp;��<b><%=ResultCount%></b>�����"<b><%=request("q")%></b>"�Ĳ�ѯ�����
<%
If ResultCount < 1 Then
	response.write ""
Else
	EndRecode=Currentpage*PageSize
	if EndRecode > ResultCount Then
		EndRecode = ResultCount
	End If

	If EndRecode - PageSize < 0 Then
		response.write "�����ǵ�<b>" & 1 & "-" & EndRecode & "</b>�"
	Else
		response.write "�����ǵ�<b>" & PageSize*(Currentpage-1)+1 & "-" & EndRecode & "</b>�"
	End If
End If
%>
</b>������ʱ <b><%=round((EndTime-StartTime)*1000,2)%></b> ���롣
<%
Select Case request("stype")
	Case "ebooks"
		response.write "��������������Ķ�������ĵ����鼮�ˣ�"
	Case "slib"
		response.write "���ڿ���ȥͼ���ҽ�����������鼮�ˣ�"
	Case Else
		response.write ""
End Select
%>
</div>