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
<title>ͼ���ѯϵͳ</title>
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
<img alt="ͼ���ѯϵͳ" src="images/jsmz_logo_large.png">
<br clear=all>
<br clear=all>
<form action="result.asp" name="f" method="post">
<div nowrap>
<span style="white-space: nowrap"><input type="radio" id="ebooks" name="stype" value="ebooks" checked="true" onclick="ShowElement('elib')"><label for="ebooks">����ͼ��</label></span>
<span style="white-space: nowrap"><input type="radio" id="slib" name="stype" value="slib" onclick="HideElement('elib')"><label for="slib">ѧУͼ����</label></span>
<span style="white-space: nowrap"><input type="radio" id="res" name="stype" value="res" onclick="HideElement('elib')"><label for="res">ѧУ��Դ��</label></span>
<span style="white-space: nowrap"><input type="radio" id="scodes" name="stype" value="codes" onclick="HideElement('elib')"><label for="scodes">��Ͽ�����</label></span>
</div>
<br clear=all>
<div nowrap>
<input type="hidden" name="page" value="1">
<input maxlength="50%" id="q" name="q" size="30%" title="������������ؼ��ʣ��ÿո����" value="" onfocus="this.select()" onmouseover="this.select()">
<input name="btnS" type="submit" value="��ʼ����">
</div>
</form>
<br clear=all>
<div id="Rnd"align="center">
<%If ResultCount > 0 then
	response.Write("<p>����Ƽ��鼮��<a href=""/eLibs/" & MyRs(1) & chr(34) & _
	" title=" & chr(34) &"��������Ķ����������غ��Ķ�����������Ķ�����ϵ��Ϣ������" & chr(34) & ">" & _
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
'��������ʾ����֪ͨ
'Set MyRs = Server.CreateObject("ADODB.RecordSet")
'Set MyConn=Server.CreateObject("ADODB.Connection")

My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
MyConn.Open My_conn_STRING

MySQL="select top 1 * from [tongzhi] order by ShiJian desc"
MyRs.cursorlocation=3 
MyRs.open MySQL,MyConn,3,2

if not MyRs.eof then
%>
<h2 align="center">����֪ͨ</h2>
<hr>
<p>�ꡡ�⣺<strong><%=MyRs("BiaoTi")%></strong></p>
<p>�ؼ��ʣ�<%=MyRs("GuanJianCi")%></p>
<p>�ڡ��ݣ�</p><div style="line-height:150%; text-align:justify; text-indent:2em; "><%=MyRs("NeiRong")%></div>
<p>�����ߣ�<%=MyRs("ZuoZhe")%>��֪ͨʱ�䣺<%=MyRs("ShiJian")%></p>
<%
else
	response.write "<h2 align=""center"">û���κ�֪ͨ</h2>"
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