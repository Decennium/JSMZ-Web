<!--#include file="../include/top.asp"-->
<!--#include file="../include/UpLoadClass.asp"-->
<%
dim request2 
'�����ϴ�����
set request2=New UpLoadClass

'����Ϊ�ֶ�����ģʽ
request2.AutoSave=2
request2.FileType=""
request2.MaxSize=1024*1024*10 '10M
request2.SavePath="upload/"
'�򿪶���Ĭ��Ϊ gb2312 �ַ�������û����ʾ����
request2.Open()
if request2.Save("Afile",1) then
	FileName=request2.form("Afile")
Else
	FileName=""
end if

'ȡ�ñ�����
ZuoZhe=htmlencode(request2.form("ZuoZhe"))
BiaoTi=htmlencode(request2.form("BiaoTi"))
GuanJianCi=htmlencode(request2.form("GuanJianCi"))
NeiRong=htmlencode(request2.form("NeiRong"))

if Len(NeiRong) And Len(BiaoTi) And Len(GuanJianCi) = False then 
	response.write("֪ͨ���ݲ���Ϊ��<br>")
	response.write("<a href='add.htm'>����</a>")
	response.end()
end if

set rs=server.CreateObject("adodb.recordset") '����rs��¼��
sql="select * from [TongZhi]" '��ȡ���ݿ��SQL��䴮
rs.open sql,conn,3,3 '�򿪼�¼�� ������Ҫ�����ݿ���и��²���ʱ����3,3�����ֻ��Ҫ��ȡ���ݿ⣬��1,1��
rs.addnew
rs("ZuoZhe")=ZuoZhe
rs("BiaoTi")=BiaoTi
rs("GuanJianCi")=GuanJianCi
rs("NeiRong")=NeiRong
rs("ShiJian")=now()
rs("Fujian")=FileName
rs.update '�������ݿ�
rs.close '�رռ�¼��
set rs=nothing '��ռ�¼��
'�ͷ��ϴ�����
set request2=nothing

response.redirect("index.asp") '�ύ�ɹ���ת��index.asp�ļ�����ȡ֪ͨ����
%>

