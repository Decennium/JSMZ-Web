<%
Response.Charset = "gb2312"
Response.Buffer = True

if len(Request("q")) > 0 then
	StartTime=Timer()
PageSize=20
Currentpage = 0 + request("page")
If Currentpage < 1 Then Currentpage = 1
	sq=split(Request("q"))
	q=""""
	for i =lbound(sq) to ubound(sq)
		q = q & sq(i) & """ OR """
	next
	if right(q,5)=" OR """ then q =left(q,len(q)-5)
	Dim strSearch
	Set MyRs = Server.CreateObject("ADODB.Recordset")

	strConn = "Provider=MSIDXS; Data Source=web"

	strSearch = "SELECT DocTitle, vPath, FileName, Size, Characterization,Rank FROM SCOPE()" & _
		" WHERE CONTAINS (DocTitle, '" & q & "') Order By Rank DESC"
MyRs.cursorlocation=3 
	MyRs.Open strSearch,strConn,3,2
	if MyRs.RecordCount < 1 then
		MyRs.Close
		strSearch = "SELECT DocTitle, vPath, FileName, Size, Characterization,Rank FROM SCOPE()" & _
			" WHERE FREETEXT (Characterization, '" & q & "') Order By Rank DESC"
MyRs.cursorlocation=3 
		MyRs.Open strSearch,strConn,3,2
	end if
MyRs.PageSize=PageSize
ResultCount=MyRs.RecordCount
If ResultCount > MyRs.PageSize Then
	ShowPage = MyRs.PageSize
Else
	ShowPage = ResultCount
End If
If MyRs.RecordCount > 0 Then
	MyRs.absolutepage = Currentpage
End If
%>
<!--#include file="result_top.asp"-->
<table align="left" width="80%">
<%
For i = 1 to ShowPage
	If MyRs.EOF Then Exit For
	if 1 = i mod 2 then
		response.write("<tr class="& chr(34) & "odd"& chr(34) &">")
	else
		response.write("<tr>")
	end if
		if len(MyRs("FileName"))>0 And Len(MyRs("DocTitle"))>0 then
			URL = MyRs("DocTitle")
		else
			URL = MyRs("FileName") & MyRs("DocTitle")
		end if
		URL = "<td><A HREF='" & MyRs("vPath") & "'>" & URL & " </A><span style='margin-left:10px'>" _
		& "Size:" & round(clng(MyRs("Size"))/1024,2) & "KB</span><BR>" & MyRs("Characterization") & "<td>"
'		if lcase(right(MyRs("FileName"),3))="mp3" then
'			URL= URL & "<object type='audio/mpeg' data='" & MyRs("vPath") & "'><PARAM NAME='autoplay' VALUE='false'><PARAM NAME='autostart' VALUE='false'></object>"
'		end if
		Response.Write URL
	response.write("</tr>")
	MyRs.movenext
Next
%>
</table>
<br clear=all>
<%
	response.write "<hr><p>���� " & (Timer() - StartTime)*1000 & "����.</p>"

	MyRs.Close
	Set MyRs = Nothing
end if

response.write "���ҳ�룺"
PageCount=cint(ResultCount/PageSize)+1

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

response.write "<a href=""result.asp?stype=res&q=" & Request("q") & "&page=1"">��һҳ</a> "
for i=StartPage to EndPage
	response.write "<a href=""result.asp?stype=res&q=" & Request("q") & "&page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""result.asp?stype=res&q=" & Request("q") & "&page=" & PageCount & """>���ҳ</a> "
%>
<!--#include file="result_bottom.asp"-->

</BODY>
</HTML>