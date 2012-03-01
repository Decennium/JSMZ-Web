<%
Response.Charset = "gb2312"
Response.Buffer = True

StartTime=Timer()

q=request("q")
PageSize=20
Currentpage=request("page")

Set MyConn=Server.CreateObject("ADODB.Connection")
MyConn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"DATA SOURCE=" & server.mappath("data/lib.mdb")

q=Replace(q,"'","")

'MySQL="SELECT COUNT(*) As ResultCount " & _
'"FROM 图书登记簿 WHERE 图书登记簿.书名 LIKE '%" & q & "%' " & _
'"OR 图书登记簿.主要内容 LIKE '%" & q & "%' " & _
'"OR 图书登记簿.著译者 LIKE '%" & q & "%' "
MySQL_Head="SELECT COUNT(*) As ResultCount FROM 图书登记簿 WHERE "
MySQL="(图书登记簿.书名 LIKE '%quarystring%' " & _
"OR 图书登记簿.主要内容 LIKE '%quarystring%' " & _
"OR 图书登记簿.著译者 LIKE '%quarystring%') "

argArray=split(q)
For Each x In argArray
	MySQL_Body=MySQL_Body & Replace(MySQL,"quarystring",x) & "AND "
Next
MySQL = MySQL_Head & MySQL_Body & "1=1 "

Set MyRs=MyConn.Execute(MySQL)
ResultCount=MyRs("ResultCount")

'MySQL="SELECT * FROM [SELECT TOP "&PageSize&" * " & _
'"FROM (SELECT TOP "&PageSize*Currentpage&" 图书登记簿.分类号, 图书登记簿.书名, 图书登记簿.主要内容, 图书登记簿.著译者, 图书登记簿.出版社 " & _
'"FROM 图书登记簿 WHERE 图书登记簿.书名 LIKE '%" & q & "%' " & _
'"OR 图书登记簿.主要内容 LIKE '%" & q & "%' " & _
'"OR 图书登记簿.著译者 LIKE '%" & q & "%' " & _
'"ORDER BY 图书登记簿.分类号 ) ORDER BY 图书登记簿.分类号 DESC ]. AS N_Result ORDER BY 图书登记簿.分类号 "
MySQL_Head="SELECT * FROM [SELECT TOP "&PageSize&" * " & _
"FROM (SELECT TOP "&PageSize*Currentpage&" 图书登记簿.分类号, 图书登记簿.书名, 图书登记簿.主要内容, 图书登记簿.著译者, 图书登记簿.出版社 " & _
"FROM 图书登记簿 WHERE "
MySQL="(图书登记簿.书名 LIKE '%quarystring%' " & _
"OR 图书登记簿.主要内容 LIKE '%quarystring%' " & _
"OR 图书登记簿.著译者 LIKE '%quarystring%') "
MySQL_Tail="ORDER BY 图书登记簿.分类号 ) ORDER BY 图书登记簿.分类号 DESC ]. AS N_Result ORDER BY 图书登记簿.分类号 "

For Each x In argArray
	MySQL_Body=MySQL_Body & Replace(MySQL,"quarystring",x) & "AND "
Next
MySQL = MySQL_Head & MySQL_Body & "1=1 " & MySQL_Tail

'response.write MySQL
Set MyRs=MyConn.Execute(MySQL)
%>
<!--#include file="result_top.asp"-->
<%If ResultCount>0 then%>
<table align="left" width="100%">
<thead>
<tr class="odd">
<% 'Put Headings On The Table of Field Names
howmanyfields=MyRs.fields.count -1 %>
<% for i=0 to howmanyfields %>
	<th><b><%=MyRs(i).name %></b></th>
<% next %>
<tr>
</thead>
<tbody>
<% ' Get all the records
j=0
do while not MyRs.eof
j=j+1
if 1 = j mod 2 then
	%><tr id='Data' class="odd"><%
else
	%><tr id='Data'><%
end if

for i = 0 to howmanyfields
	ThisRecord = MyRs(I)
	If IsNull(ThisRecord) Then
		ThisRecord = "&nbsp;"
	end if
	select case MyRs(I).name
		case "主要内容":
			CellWidth=""
		case "书名":
			CellWidth="200px"
		case "著译者","出版社":
			CellWidth="125px"
		case "分类号":
			CellWidth="100px"
		case else
			CellWidth=""
	end select
	%><td width="<%=CellWidth%>"><%=ThisRecord%></td><%
next %>
</tr>
<% MyRs.movenext
loop %>
</tbody>
</table>
<br clear=both>
<div id="navbar" align="left">
<%
response.write "结果页码："
PageCount=Int(ResultCount/PageSize)+1
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

response.write "<a href=""result.asp?stype=slib&q=" & q & "&page=1"">第一页</a> "
for i=StartPage to EndPage
	response.write "<a href=""result.asp?stype=slib&q=" & q & "&page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""result.asp?stype=slib&q=" & q & "&page=" & PageCount & """>最后页</a> "
%>
<div>
<%
Else
	Response.Write "<h1>没有找到任何结果，请更改关键词，并重新搜索。</h1>"
End If%>
<!--#include file="result_bottom.asp"-->
<%
MyRs.close
Set MyRs= Nothing
MyConn.Close
set MyConn=nothing
%>
	</body>
</html>

