<%
Response.Charset = "gb2312"
Response.Buffer = True

StartTime=Timer()

q=request("q")
PageSize=20
Currentpage=request("page")
If Currentpage < 1 Then Currentpage = 1
Set MyRs = Server.CreateObject("ADODB.RecordSet")
Set MyConn=Server.CreateObject("ADODB.Connection")

'My_conn_STRING =  "PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=" & server.mappath("data/ebooks.mdb")
My_conn_STRING = "Provider=SQLOLEDB;server=C3Server;database=Resources;uid=sa;pwd="
MyConn.Open My_conn_STRING

MySQL_Head="SELECT Author, FileName, BookName FROM eBooks WHERE "'MySQL="(InStr(1,LCase(eBooks.FileName),LCase('quarystring'),0)<>0) "
MySQL="(Author Like '%quarystring%' OR BookName Like '%quarystring%') "
'MySQL_Tail="ORDER BY eBooks.FileName"
MySQL_Tail=""
MySQL_Body = ""
q=Replace(q,"'","")
If q <> "" And q <> "*" Then
	argArray=split(q)
	For Each x In argArray
		MySQL_Body=MySQL_Body & Replace(MySQL,"quarystring",x) & "AND "
	Next
	MySQL_Body = Left(MySQL_Body,Len(MySQL_Body)-4)
Else
	MySQL_Body="1=1 "
End If

MySQL = MySQL_Head & MySQL_Body & MySQL_Tail
'response.write "<p>" & MySQL & "</p>"
MyRs.cursorlocation=3 
MyRs.open MySQL,MyConn,3,2
MyRs.PageSize=PageSize

ResultCount=MyRs.recordcount
%>
<!--#include file="result_top.asp"-->
<%If ResultCount>0 then%>
<table align="left" width="100%">
<thead>
<tr class="odd">
<% 'Put Headings On The Table of Field Names
howmanyfields=MyRs.fields.count -1 %>
<%
for i=0 to howmanyfields
	Select Case MyRs(i).Name
		Case "Author":
			response.Write "<th><b>" & "作者" & "</b></th>"
		Case "BookName":
			response.Write "<th><b>" & "书名" & "</b></th>"
		Case Else
			
	End Select
'	If MyRs(i).Name <> "FileName" Then
'		response.Write "<th><b>" & MyRs(i).name & "</b></th>"
'	End If
next %>
<tr>
</thead>
<tbody>
<% ' Get all the records
If ResultCount > MyRs.PageSize Then
	ShowPage = MyRs.PageSize
Else
	ShowPage = ResultCount
End If
MyRs.absolutepage = Currentpage
'If MyRs.EOF Then MyRs.MoveFirst
'MyRs.Move MyRs.PageSize * (MyRs.AbsolutePage - 1)
'response.write MyRs.EOF
For i = 1 to ShowPage
	If MyRs.EOF Then Exit For
	if 1 = i mod 2 then
		response.write("<tr id='Data' class="& chr(34) & "odd"& chr(34) &">")
	else
		response.write("<tr id='Data'>")
	end if
	response.Write("<td width=200>" & MyRs(0)  & "</td>")
	response.Write("<td><a href="& chr(34) & MyRs(1) & chr(34) & _
	" title=" & chr(34) &"点击即可阅读，或者下载后阅读。如果不能阅读请联系信息技术组" & chr(34) & ">" & _
	MyRs(2)  & "</a></td>")
	response.write("</tr>")
	MyRs.movenext
Next
%>
</tbody>
</table>
<br clear=both>
<div id="navbar" align="left">
<!--#include file="result_bottom.asp"-->
<%
MyRs.close
Set MyRs= Nothing
MyConn.Close
set MyConn=nothing
%>
	</body>
</html>

