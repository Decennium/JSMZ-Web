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

My_conn_STRING = "Provider=SQLOLEDB;server=C3Server;database=Resources;uid=sa;pwd="
MyConn.Open My_conn_STRING

MySQL="SELECT Author, FileName, BookName FROM eBooks WHERE ID=" & BookID
MyRs.cursorlocation=3 
MyRs.open MySQL,MyConn,3,2

ResultCount=MyRs.recordcount

%>
<html>
<head>
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
<input type="radio" id="ebooks" name="stype" value="ebooks" checked="true" onclick="ShowElement('elib')"><label for="ebooks">����ͼ��</label>
<input type="radio" id="slib" name="stype" value="slib" onclick="HideElement('elib')"><label for="slib">ѧУͼ����</label>
<input type="radio" id="zheda" name="stype" value="zheda" onclick="HideElement('elib')"><label for="zheda">�����Դ��</label>
<input type="radio" id="scodes" name="stype" value="codes" onclick="HideElement('elib')"><label for="scodes">��Ͽ�����</label>
</div>
<br clear=all>
<div nowrap>
<input type="hidden" name="page" value="1">
<input maxlength="250" id="q" name="q" size="55" title="������������ؼ��ʣ��ÿո����" value="" onfocus="this.select()" onmouseover="this.select()">
<input name="btnS" type="submit" value="��ʼ����">
</div>
</form>
<br clear=all>
<div id="Rnd" align="center">
<%If ResultCount > 0 then
	response.Write("<p>����Ƽ��鼮��<a href="& chr(34) & MyRs(1) & chr(34) & _
	" title=" & chr(34) &"��������Ķ����������غ��Ķ�����������Ķ�����ϵ��Ϣ������" & chr(34) & ">" & _
	MyRs(2)  & "</a></p>")
End If

MyRs.close
Set MyRs= Nothing
MyConn.Close
set MyConn=nothing
%>
</div>
<!--#include file="../include/bottom.asp"-->
</body>
</html>