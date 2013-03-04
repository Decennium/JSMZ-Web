<%
Dim max,min
max=87351
min=1
Randomize
BookID=(Int((max-min+1)*Rnd+min))

Response.Charset = "gb2312"
Response.Buffer = True

Set MyRs = Server.CreateObject("ADODB.RecordSet")
Set MyConn=Server.CreateObject("ADODB.Connection")

My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=Resources;uid=sa;pwd="
MyConn.Open My_conn_STRING

MySQL="SELECT Author, FileName, BookName FROM eBooks WHERE ID=" & BookID
MyRs.cursorlocation=3 
MyRs.open MySQL,MyConn,3,2

ResultCount=MyRs.recordcount

%>
<html>
<head>
<meta name="viewport" content="width=device-width,minimum-scale=1.0, maximum-scale=2.0"/>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>图书查询系统</title>
<link href="../css/css.css" rel="stylesheet">
<link rel="shortcut icon" type="image/x-icon" href="favicon.ico">
<script type="text/javascript">
function setFocus()
{
document.getElementById("q").focus();
}
function ShowElement(id) {
	var e = document.getElementById(id);
	e.style.display = 'block';
}
function HideElement(id) {
var e = document.getElementById(id);
	e.style.display = 'none';
}
if(self.frameElement.tagName=="IFRAME"){
	top.document.location.href="/eBooks/"
}
</script>
</head>
<body onload="setFocus()">
<!--#include file="../include/banner.asp"-->
<br clear=all>
<br clear=all>
<br clear=all>
<img alt="图书查询系统" src="images/jsmz_logo_large.png">
<br clear=all>
<br clear=all>
<form action="result.asp" name="f" method="post">
<div nowrap>
<span style="white-space: nowrap"><input type="radio" id="ebooks" name="stype" value="ebooks" checked="true" onclick="ShowElement('elib')"><label for="ebooks">电子图书</label></span>
<span style="white-space: nowrap"><input type="radio" id="slib" name="stype" value="slib" onclick="HideElement('elib')"><label for="slib">学校图书室</label></span>
<span style="white-space: nowrap"><input type="radio" id="res" name="stype" value="res" onclick="HideElement('elib')"><label for="res">学校资源库</label></span>
<span style="white-space: nowrap"><input type="radio" id="scodes" name="stype" value="codes" onclick="HideElement('elib')"><label for="scodes">诊断卡代码</label></span>
</div>
<br clear=all>
<div nowrap>
<input type="hidden" name="page" value="1">
<input maxlength="50%" id="q" name="q" size="30%" title="你可以输入多个关键词，用空格隔开" value="" onfocus="this.select()" onmouseover="this.select()">
<input name="btnS" type="submit" value="开始搜索">
</div>
</form>
<br clear=all>
<div id="Rnd"align="center">
<%If ResultCount > 0 then
	response.Write("<p>随机推荐书籍：<a href=""/eLibs/" & MyRs(1) & chr(34) & _
	" title=" & chr(34) &"点击即可阅读，或者下载后阅读。如果不能阅读请联系信息技术组" & chr(34) & ">" & _
	MyRs(2)  & "</a></p>")
End If

MyRs.close
'Set MyRs= Nothing
MyConn.Close
'set MyConn=nothing
%>
<div style="width:600px;text-align:left;" align="center">
<hr>
<%
'接下来显示当月通知
'Set MyRs = Server.CreateObject("ADODB.RecordSet")
'Set MyConn=Server.CreateObject("ADODB.Connection")

My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
MyConn.Open My_conn_STRING

MySQL="select top 1 * from [tongzhi] order by ShiJian desc"
MyRs.cursorlocation=3 
MyRs.open MySQL,MyConn,3,2

if not MyRs.eof then
%>
<h2 align="center">最新通知</h2>
<hr>
<p>标　题：<strong><%=MyRs("BiaoTi")%></strong></p>
<p>关键词：<%=MyRs("GuanJianCi")%></p>
<p>内　容：</p><div style="line-height:150%; text-align:justify; text-indent:2em; "><%=MyRs("NeiRong")%></div>
<p>发布者：<%=MyRs("ZuoZhe")%>，通知时间：<%=MyRs("ShiJian")%></p>
<%
else
	response.write "<h2 align=""center"">没有任何通知</h2>"
end if
MyRs.close

Set MyRs= Nothing
MyConn.Close
set MyConn=nothing

Function IIf( expr, truepart, falsepart )
   IIf = falsepart
   If expr Then IIf = truepart
End Function
%>
<hr>
</div>
</div>
<!--#include file="../include/bottom.asp"-->
</body>
</html>