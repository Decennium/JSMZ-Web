<!--#include file="../include/top.asp"-->
<%
if session("Admin")="" then response.redirect ("index.asp")
id=request.querystring("id")
if id="" then 
	response.write("����")
	response.end()
end if
conn.execute("delete from [TongZhi] where id='"&id & "'") 'ɾ��֪ͨ
response.Redirect("index.asp")
%>