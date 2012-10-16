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
	response.write "<title>所有电子图书</title>"
Else
	response.write "<title>" & request("q") & " - 搜索结果</title>"
End If
%>
<title><%=request("q")%> - 搜索结果</title>
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
<a id=logo href="." title="首页"><img src="images/jsmz_logo_mini.png" border="0px"></a>
</div>
<div align="left" style="clear:right; float:left; ">
<div>
<span style="white-space: nowrap"><input type="radio" id="ebooks" name="stype" value="ebooks" <%if request("stype")="ebooks" then response.write "checked=true"%>><label for="ebooks">电子图书</label></span>
<span style="white-space: nowrap"><input type="radio" id="slib" name="stype" value="slib" <%if request("stype")="slib" then response.write "checked=true"%>><label for="slib">学校图书室</label></span>
<span style="white-space: nowrap"><input type="radio" id="res" name="stype" value="res" <%if request("stype")="res" then response.write "checked=true"%>><label for="res">学校资源库</label></span>
<span style="white-space: nowrap"><input type="radio" id="scodes" name="stype" value="codes" <%if request("stype")="codes" then response.write "checked=true"%>><label for="scodes">诊断卡代码</label></span>
</div>
<div>
<input type="hidden" name="page" value="1">
<input  maxlength="50%" id="q" name="q" size="30%" title="你可以输入多个关键词，用空格隔开" value="<%=request("q")%>" onfocus="this.select()" onmouseover="this.select()">
<input name="btnS" type="submit" value="搜索">
</div>
</div>
</form>
<div class="HeadLine" style="clear:both;">
&nbsp;有<b><%=ResultCount%></b>项符合"<b><%=request("q")%></b>"的查询结果，
<%
If ResultCount < 1 Then
	response.write ""
Else
	EndRecode=Currentpage*PageSize
	if EndRecode > ResultCount Then
		EndRecode = ResultCount
	End If

	If EndRecode - PageSize < 0 Then
		response.write "以下是第<b>" & 1 & "-" & EndRecode & "</b>项。"
	Else
		response.write "以下是第<b>" & PageSize*(Currentpage-1)+1 & "-" & EndRecode & "</b>项。"
	End If
End If
%>
</b>搜索用时 <b><%=round((EndTime-StartTime)*1000,2)%></b> 毫秒。
<%
Select Case request("stype")
	Case "ebooks"
		response.write "现在你可以慢慢阅读你中意的电子书籍了！"
	Case "slib"
		response.write "现在可以去图书室借阅你中意的书籍了！"
	Case Else
		response.write ""
End Select
%>
</div>