<!--#include file="../include/top.asp"-->
<html>
<head>
<meta name="viewport" content="width=device-width,minimum-scale=1.0, maximum-scale=2.0"/>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<link href="../css/css.css" rel="stylesheet">
<%
set rs=server.CreateObject("adodb.recordset") '创建rs记录集
sql="select * from [tongzhi] order by ShiJian desc" '读取数据库的SQL语句串,按通知添加的时候排序
rs.open sql,conn,1,1 '打开记录集
%>
<title>通知列表</title>
</head>

<body>
<!--#include file="../include/banner.asp"-->
<div style="float:left;">
<!--#include file="../include/left_banner.asp"-->
</div>
<div style="width="100%";float:left">
<div style="text-align:left;">
<%if session("Admin")="" then%>
<a href="/BOS/">登录管理通知</a>
<%else%>
<a href="add.asp">提交通知</a>
<%end if%> 
<hr>
<%do while not rs.eof%>
<p>发布者：<%=rs("ZuoZhe")%></p>
<p>标　题：<strong><%=rs("BiaoTi")%></strong></p>
<p>关键词：<%=rs("GuanJianCi")%></p>
<p>内　容：</p><div style="line-height:150%; text-align:justify; text-indent:2em; "><%=rs("NeiRong")%></div>
<p>通知时间：</font><%=rs("ShiJian")%></p>
<%if session("Admin")<>"" then%>
<p><a href="del.asp?id=<%=rs("id")%>">[删除]</a></p>
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
