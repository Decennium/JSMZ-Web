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

MySQL_Head="SELECT COUNT(*) As ResultCount FROM ͼ��Ǽǲ� WHERE "
MySQL="(ͼ��Ǽǲ�.���� LIKE '%quarystring%' " & _
"OR ͼ��Ǽǲ�.��Ҫ���� LIKE '%quarystring%' " & _
"OR ͼ��Ǽǲ�.������ LIKE '%quarystring%') "

argArray=split(q)
For Each x In argArray
	MySQL_Body=MySQL_Body & Replace(MySQL,"quarystring",x) & "AND "
Next
MySQL = MySQL_Head & MySQL_Body & "1=1 "

Set MyRs=MyConn.Execute(MySQL)
ResultCount=MyRs("ResultCount")

MySQL_Head="SELECT * FROM [SELECT TOP "&PageSize&" * " & _
"FROM (SELECT TOP "&PageSize*Currentpage&" ͼ��Ǽǲ�.�����, ͼ��Ǽǲ�.����, ͼ��Ǽǲ�.��Ҫ����, ͼ��Ǽǲ�.������, ͼ��Ǽǲ�.������ " & _
"FROM ͼ��Ǽǲ� WHERE "
MySQL="(ͼ��Ǽǲ�.���� LIKE '%quarystring%' " & _
"OR ͼ��Ǽǲ�.��Ҫ���� LIKE '%quarystring%' " & _
"OR ͼ��Ǽǲ�.������ LIKE '%quarystring%') "
MySQL_Tail="ORDER BY ͼ��Ǽǲ�.����� ) ORDER BY ͼ��Ǽǲ�.����� DESC ]. AS N_Result ORDER BY ͼ��Ǽǲ�.����� "

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
		case "��Ҫ����":
			CellWidth=""
		case "����":
			CellWidth="200px"
		case "������","������":
			CellWidth="125px"
		case "�����":
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
<!--#include file="result_bottom.asp"-->
<%
MyRs.close
Set MyRs= Nothing
MyConn.Close
set MyConn=nothing
%>
	</body>
</html>

