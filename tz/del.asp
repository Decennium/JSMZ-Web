<!--#include file="../include/top.asp"-->
<%
Sub DeleteAFile(filespec)
	If Len(trim(filespec))<1 Then Exit Sub
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(filespec) Then fso.DeleteFile(filespec)
End Sub

if session("Admin")="" then response.redirect ("index.asp")
id=request.querystring("id")
if id="" then 
	response.write("����")
	response.end()
end if

MySQL="select FuJian from [TongZhi] where id='" & id & "'"
MyRs.cursorlocation=3 
MyRs.open MySQL,Conn,3,2
If Len(MyRs("FuJian")) > 1 Then
	DeleteAFile server.mappath(".") & "\upload\" & MyRs("FuJian")
	'�ڴ˼���Ƿ�ɾ���ɹ����ɹ���ɾ����¼
End If
MyRs.Close
Set MyRs= Nothing

conn.execute("delete from [TongZhi] where id='"&id & "'") 'ɾ��֪ͨ
response.Redirect("index.asp")
%>