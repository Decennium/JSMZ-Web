<!--#include file="../include/top.asp"-->
<%
'取得表单数据
ZuoZhe=htmlencode(request.form("ZuoZhe"))
BiaoTi=htmlencode(request.form("BiaoTi"))
GuanJianCi=htmlencode(request.form("GuanJianCi"))
NeiRong=htmlencode(request.form("NeiRong"))

if Len(NeiRong) And Len(BiaoTi) And Len(GuanJianCi) = False then 
	response.write("通知内容不能为空<br>")
	response.write("<a href='add.htm'>返回</a>")
	response.end()
end if
set rs=server.CreateObject("adodb.recordset") '创建rs记录集
sql="select * from [TongZhi]" '读取数据库的SQL语句串
rs.open sql,conn,3,3 '打开记录集 ，当需要对数据库进行更新操作时，用3,3，如果只需要读取数据库，用1,1。
rs.addnew
rs("ZuoZhe")=ZuoZhe
rs("BiaoTi")=BiaoTi
rs("GuanJianCi")=GuanJianCi
rs("NeiRong")=NeiRong
rs("ShiJian")=now()
rs.update '更新数据库
rs.close '关闭记录集
set rs=nothing '清空记录集
response.redirect("index.asp") '提交成功后，转向到index.asp文件，读取通知内容
%>

