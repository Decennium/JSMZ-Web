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
	response.write("错误")
	response.end()
end if

MySQL="select FuJian from [TongZhi] where id='" & id & "'"
MyRs.cursorlocation=3 
MyRs.open MySQL,Conn,3,2
If Len(MyRs("FuJian")) > 1 Then
	DeleteAFile server.mappath(".") & "\upload\" & MyRs("FuJian")
	'在此检查是否删除成功，成功则删除记录
End If
MyRs.Close
Set MyRs= Nothing

conn.execute("delete from [TongZhi] where id='"&id & "'") '删除通知
response.Redirect("index.asp")
%>