<!--#include file="Config.asp" -->
<%
If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
If isNull(Conn) Then
	Set Conn=Server.CreateObject("ADODB.Connection")
	My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
	Conn.Open My_conn_STRING
End If
Action=Request.Querystring("Action")
Select Case Action
	Case "jifang"
		Sql="Select * from Jifang order by RiQi DESC"
	Case "shiyong"
		Sql="Select * from shiyong order by RiQi DESC"
	Case "gouru"
		Sql="Select * from gouru order by RiQi DESC"
	Case "luyinji"
		Sql="Select * from luyinji order by RiQi DESC"
	Case "MMC"
		Sql="Select * from MMC order by RiQi DESC"
	Case "jiankong"
		Sql="Select * from jiankong order by RiQi DESC"
	Case "weixiu"
		Sql="Select * from weixiu order by RiQi DESC"
	Case "waijie"
		Sql="Select * from waijie order by RiQi DESC"
	Case "ereading"
		Sql="Select * from eread order by RiQi DESC"
	Case "tushushi"
		Sql="Select * from tushushi order by RiQi DESC"
	Case Else
		Response.Redirect "index.asp"
		Response.End
End Select

MyRs.open Sql,Conn,3,2

howmanyfields=MyRs.fields.count -1 
xlsStream = "<table><tr>"
for i=1 to howmanyfields
	xlsStream = xlsStream & "<th>" & UCase(MyRs(i).Name) & "</th>"
next
xlsStream = xlsStream &  "</tr>"

AllRecord = MyRs.RecordCount
If MyRs.RecordCount > 65535 Then AllRecord = 65535

For i_s = 1 to AllRecord
	If MyRs.EOF Then Exit For
	xlsStream = xlsStream & "<tr>"
	for i_c = 1 to howmanyfields '²»ÏÔÊ¾Id×Ö¶Î
		ThisRecord = MyRs(i_c).value
		If IsNull(ThisRecord) Then
			ThisRecord = "&nbsp;"
		end if
		xlsStream = xlsStream & "<td>" & ThisRecord & "</td>"
	next
	xlsStream = xlsStream & "</tr>"
	MyRs.movenext
Next
xlsStream = xlsStream & "</table>"
xlsFile = Action & "-" & now() & ".XLS"

Response.ContentType="application/vnd.ms-excel"
Response.AddHeader "content-disposition","attachment;filename=" & xlsFile

Response.Write xlsStream
Response.End()
%>