	<!-- Created: 0:00:00 -->
<html>
	<head>
		<meta name="GENERATOR" Content="ASP Express 5.0">
		<title>Untitled</title>
	</head>
	<body>
<% Set MyConn=Server.CreateObject("ADODB.Connection")
MyConn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"DATA SOURCE=" & server.mappath("data/lib.mdb")
MySQL="Select * from [Í¼ÊéµÇ¼Ç²¾]"
Set MyRs=MyConn.Execute(MySQL)
%>
<table border="1">
<tr bgcolor=""><% 'Put Headings On The Table of Field Names
howmanyfields=MyRs.fields.count -1 %>
<% for i=0 to howmanyfields %>
       <td><b><font color=""><%=MyRs(i).name %></font> </b></td>
<% next %>
</tr>
<% ' Get all the records
 do  while not MyRs.eof %>
<tr bgcolor="" id='Data'>
<% for i = 0 to howmanyfields
	ThisRecord = MyRs(I)
	If IsNull(ThisRecord) Then
	ThisRecord = "&nbsp;"
end if %>
       <td valign=top><font color=""><%=Thisrecord%></font></td><% next %>
</tr>
<% MyRs.movenext
loop %>
</table>

<%
MyRs.close
Set MyRs= Nothing
MyConn.Close
set MyConn=nothing
%>
	</body>
</html>