<!--#include file="../include/top.asp"-->
<html>
<head>
<meta name="viewport" content="width=device-width,minimum-scale=1.0, maximum-scale=2.0"/>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<link href="../css/css.css" rel="stylesheet">
<%
set rs=server.CreateObject("adodb.recordset") '����rs��¼��
sql="select * from [tongzhi] order by ShiJian desc" '��ȡ���ݿ��SQL��䴮,��֪ͨ��ӵ�ʱ������
rs.open sql,conn,1,1 '�򿪼�¼��
%>
<title>֪ͨ�б�</title>
</head>

<body>
<!--#include file="../include/banner.asp"-->
<div style="float:left;">
<!--#include file="../include/left_banner.asp"-->
</div>
<div style="width="100%";float:left">
<div style="text-align:left;">
<%if session("Admin")="" then%>
<a href="/BOS/">��¼����֪ͨ</a>
<%else%>
<a href="add.asp">�ύ֪ͨ</a>
<%end if%> 
<hr>
<%do while not rs.eof%>
<p>�����ߣ�<%=rs("ZuoZhe")%></p>
<p>�ꡡ�⣺<strong><%=rs("BiaoTi")%></strong></p>
<p>�ؼ��ʣ�<%=rs("GuanJianCi")%></p>
<p>�ڡ��ݣ�</p><div style="line-height:150%; text-align:justify; text-indent:2em; "><%=rs("NeiRong")%></div>
<p>֪ͨʱ�䣺</font><%=rs("ShiJian")%></p>
<%if session("Admin")<>"" then%>
<p><a href="del.asp?id=<%=rs("id")%>">[ɾ��]</a></p>
<%end if%>
<hr>
<%rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</div>
</div>
<br clear="all"><br><br><br><br><br><br><br><br><br>
<!--#include file="../include/bottom.asp"-->
</body>
</html>
