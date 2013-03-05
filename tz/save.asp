<!--#include file="../include/top.asp"-->
<!--#include file="../include/UpLoadClass.asp"-->
<%
dim request2 
'建立上传对象
set request2=New UpLoadClass

'设置为手动保存模式
request2.AutoSave=2
request2.FileType=""
request2.MaxSize=1024*1024*10 '10M
request2.SavePath="upload/"
'打开对象，默认为 gb2312 字符集，故没有显示设置
request2.Open()
if request2.Save("Afile",1) then
	FileName=request2.form("Afile")
Else
	FileName=""
end if

'取得表单数据
ZuoZhe=htmlencode(request2.form("ZuoZhe"))
BiaoTi=htmlencode(request2.form("BiaoTi"))
GuanJianCi=htmlencode(request2.form("GuanJianCi"))
NeiRong=htmlencode(request2.form("NeiRong"))

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
rs("Fujian")=FileName
rs.update '更新数据库
rs.close '关闭记录集
set rs=nothing '清空记录集
'释放上传对象
set request2=nothing

response.redirect("index.asp") '提交成功后，转向到index.asp文件，读取通知内容
%>

