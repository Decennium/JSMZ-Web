<!--#include file="../include/top.asp"-->
<%
'ȡ�ñ�����
ZuoZhe=htmlencode(request.form("ZuoZhe"))
BiaoTi=htmlencode(request.form("BiaoTi"))
GuanJianCi=htmlencode(request.form("GuanJianCi"))
NeiRong=htmlencode(request.form("NeiRong"))

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
rs.update '�������ݿ�
rs.close '�رռ�¼��
set rs=nothing '��ռ�¼��
response.redirect("index.asp") '�ύ�ɹ���ת��index.asp�ļ�����ȡ֪ͨ����
%>

