<div id="navbar" align="left">
<%
PageCount=Int(ResultCount/(PageSize+1))+1
If PageCount > 1 Then
	response.write "���ҳ�룺"
	if CurrentPage > 4 then
		StartPage=CurrentPage-4
	Else
		StartPage=1
	end if

	if PageCount <= 1 then
		EndPage=1
	Else
		if pageCount > CurrentPage + 4 then
			EndPage = CurrentPage + 4
		else
			EndPage = pageCount
		end if
	end if

	response.write "<a href=""result.asp?stype=" & request("stype") & "&q=" & request("q") & "&page=1"">��һҳ</a> "
	for i=StartPage to EndPage
		response.write "<a href=""result.asp?stype=" & request("stype") & "&q=" & request("q") & "&page=" & i & """>" & i & "</a> "
	next 'i
	response.write "<a href=""result.asp?stype=" & request("stype") & "&q=" & request("q") & "&page=" & PageCount & """>���ҳ</a> "
End If
%>
<div>
<%
Else
	If request("stype") <> "codes" Then
		Response.Write "<h1>û���ҵ��κν��������Ĺؼ��ʣ�������������</h1>"
	End If
End If
%>
<!--#include file="../include/bottom.asp"-->
