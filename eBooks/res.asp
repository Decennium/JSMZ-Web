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

strSearch = "SELECT DocTitle, vPath, FileName, Size, DocAppName, Characterization,Rank FROM SCOPE()" & _
	" WHERE CONTAINS (DocTitle, '" & q & "') Order By Rank DESC"
MyRs.cursorlocation=3 
MyRs.Open strSearch,strConn,3,2
if MyRs.RecordCount < 1 then
	MyRs.Close
	strSearch = "SELECT DocTitle, vPath, FileName, Size, DocAppName, Characterization,Rank FROM SCOPE()" & _
		" WHERE CONTAINS (Characterization, '" & q & "') Order By Rank DESC"
	MyRs.cursorlocation=3 
	MyRs.Open strSearch,strConn,3,2
end if
if MyRs.RecordCount < 1 then
	MyRs.Close
	strSearch = "SELECT DocTitle, vPath, FileName, Size, DocAppName, Contents, Characterization,Rank FROM SCOPE()" & _
		" WHERE CONTAINS (Contents, '" & q & "') Order By Rank DESC"
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
<table align="left" width="100%">
<thead>
<tr class="odd">
<% 'Put Headings On The Table of Field Names
howmanyfields=MyRs.fields.count -1 %>
<%
response.Write "<th><b>" & "����" & "</b></th>"
response.Write "<th><b>" & "ժҪ" & "</b></th>"
response.Write "<th><b>" & "��С" & "</b></th>"
response.Write "<th><b>" & "����" & "</b></th>"
%>
<tr>
</thead>
<tbody>
<%
For i = 1 to ShowPage
	If MyRs.EOF Then Exit For
	if 1 = i mod 2 then
		response.write("<tr id='Data' class="& chr(34) & "odd"& chr(34) &">")
	else
		response.write("<tr id='Data'>")
	end if
	
	If Len(MyRs("DocTitle")) > 40 then
		URL = MyRs("FileName")
	Else
		If Len(MyRs("DocTitle")) > 0 then
			URL = MyRs("DocTitle")
		Else
			URL = MyRs("FileName")
		End If
	End If

	URL ="<td nowrap='nowrap'><A HREF='" & MyRs("vPath") & "'>" & URL & " </A></td>"
	If Len(MyRs("Characterization"))>0 then
		URL = URL & "<td>" & MyRs("Characterization") & "</td>"
	else
		URL = URL & "<td>" & MyRs("FileName") & "</td>"
	End If
	URL = URL & "<td>" & round(clng(MyRs("Size"))/1024,2) & "KB</td>"
	Response.Write URL
	Dim ExtentName, FileType
	FileType = "δ֪�ĵ�"
	ExtentName=LCase(Mid(MyRs("FileName"),InstrRev(MyRs("FileName"),".")+1))
	If instr("doc dot docx dotx rtf wps wpt",ExtentName) Then FileType = "�����ĵ�"
	If instr("xls csv xlsx et",ExtentName) Then FileType = "���ӱ��"
	If instr("ppt pps pptx wpp",ExtentName) Then FileType = "��ʾ�ĸ�"
	If instr("htm html php asp shtm shtml",ExtentName) Then FileType = "��ҳ�ļ�"
	If instr("txt log",ExtentName) Then FileType = "���ı��ļ�"
	If instr("pdf epub chm",ExtentName) Then FileType = "������"
	If instr("zip rar 7z",ExtentName) Then FileType = "ѹ���ļ�"
	If instr("jpg jpeg png gif bmp",ExtentName) Then FileType = "ͼƬ�ļ�"
	If instr("wav mp3 wma mpc flc ogg",ExtentName) Then FileType = "��Ƶ�ļ�"
	If instr("mp4 mpg mpeg avi flv f4v rmvb asx",ExtentName) Then FileType = "��Ƶ�ļ�"
	Response.Write "<td nowrap='nowrap'>" & FileType & "</td>"
	response.write("</tr>")
	MyRs.movenext
Next
%>
</tbody>
</table>
<br clear=all>
<!--#include file="result_bottom.asp"-->
<%
MyRs.close
Set MyRs= Nothing
%>
</BODY>
</HTML>