<!--#include file="../include/top.asp"-->
<html>
<head>
<meta name="viewport" content="width=device-width,minimum-scale=1.0, maximum-scale=2.0"/>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<link href="../css/css.css" rel="stylesheet">
<%
Function IIf( expr, truepart, falsepart )
   IIf = falsepart
   If expr Then IIf = truepart
End Function
S_year=request("year")
S_month=request("month")

set rs=server.CreateObject("adodb.recordset") '创建rs记录集
time_1=s_year& "-" & s_month&"-1"
If IsDate(time_1) then
	sql="select * from [tongzhi] where ShiJian between '" & S_year &"-" & S_month &"-01 00:00:00' And '" & S_year+IIf(S_month>12,1,0) &"-" & (cint(S_month)+1 mod 12) &"-01 00:00:00' order by ShiJian desc"
Else
	sql="select * from [tongzhi] order by ShiJian desc"
End If
rs.open sql,conn,1,1 '打开记录集
%>
<title>通知列表</title>
</head>

<body>
<!--#include file="../include/banner.asp"-->
<div style="float:left;">
<!--#include file="../include/left_banner.asp"-->
</div>
<div id="Right_Content">
<div style="clear:right;padding-top:5px;text-align:left;">
<%
if session("Admin")<>"" then
	response.write "<a href=""add.asp"">发布通知</a>"
end if

For m=month(now()) to 1 step -1
	response.write " | <a href=""?year=" & year(Now()) & "&month=" & m & """>" & year(Now()) & "年" & m & "月" & "</a>"
Next
response.write " | <a href=""?year=&month="">全部</a>"
%>

<hr>
</div>
<div style="text-align:left;">
<%do while not rs.eof%>
<p>标　题：<strong><%=rs("BiaoTi")%></strong></p>
<p>关键词：<%=rs("GuanJianCi")%></p>
<p>内　容：</p><div style="line-height:150%; text-align:justify; text-indent:2em; "><%=rs("NeiRong")%></div>
<p>发布者：<%=rs("ZuoZhe")%>，通知时间：<%=rs("ShiJian")%></p>
<%if session("Admin")<>"" then%>
<p>附　件：<a href="<%="upload/"&rs("FuJian")%>"><%=rs("FuJian")%></a></p>
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
