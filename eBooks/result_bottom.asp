<div id="navbar" align="left">
<%
PageCount=Int(ResultCount/(PageSize+1))+1
If PageCount > 1 Then
	response.write "结果页码："
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

	response.write "<a href=""result.asp?stype=" & request("stype") & "&q=" & request("q") & "&page=1"">第一页</a> "
	for i=StartPage to EndPage
		response.write "<a href=""result.asp?stype=" & request("stype") & "&q=" & request("q") & "&page=" & i & """>" & i & "</a> "
	next 'i
	response.write "<a href=""result.asp?stype=" & request("stype") & "&q=" & request("q") & "&page=" & PageCount & """>最后页</a> "
End If
%>
<div>
<%
Else
	If request("stype") <> "codes" Then
		Response.Write "<h1>没有找到任何结果，请更改关键词，并重新搜索。</h1>"
	End If
End If
%>
<!--#include file="../include/bottom.asp"-->
