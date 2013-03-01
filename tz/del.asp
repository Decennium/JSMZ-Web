<!--#include file="../include/top.asp"-->
<%
if session("Admin")="" then response.redirect ("index.asp")
id=request.querystring("id")
if id="" then 
	response.write("´íÎó")
	response.end()
end if
conn.execute("delete from [TongZhi] where id='"&id & "'") 'É¾³ýÍ¨Öª
response.Redirect("index.asp")
%>