<!--#include file="top.asp"-->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>����ʹ�õǼǹ���ϵͳ</title>
<link href="../css/css.css" rel="stylesheet">
<link rel="shortcut icon" type="image/x-icon" href="favicon.ico">
<script language="javascript" src="DateTime.js"></script>
<script language="javascript">
/** �ж����������������ڸ�ʽΪyyyy-mm-dd����ȷ������ */ 
function My_IsDate(mystring) { 
	var reg = /^(\d{4})-(\d{1,2})-(\d{1,2})$/; 
	var str = mystring; 
	var arr = reg.exec(str); 
	if (str == " ") return true; 
	if (!reg.test(str)&&RegExp.$2 <=12&&RegExp.$3 <=31){ 
		return false; 
		} 
		return true; 
	}
function My_IsInt(val) {
	if ((parseInt(val) == val) && (parseInt(val) >= 0)) {
		return true;
	}
		return false;
}
function My_CheckFields(the){
	if (the.Riqi.value=="") {
		document.getElementById('Tips').innerHTML = '���ڲ���Ϊ��';
		the.Riqi.focus();
		return false;
	}
	if (!My_IsDate(the.Riqi.value)){
		document.getElementById('Tips').innerHTML = '�뱣֤��������ڸ�ʽΪyyyy-mm-dd����ȷ������!';
		the.Riqi.focus();
		return false;
	}
	if (the.Banji.value=="") {
		document.getElementById('Tips').innerHTML = '�༶����Ϊ��';
		the.Banji.focus();
		return false;
	}
	if (the.Jifang.value=="") {
		document.getElementById('Tips').innerHTML = '��������Ϊ��';
		the.Jifang.focus();
		return false;
	}
	if (the.Neirong.value=="") {
		document.getElementById('Tips').innerHTML = '���ݲ���Ϊ��';
		the.Neirong.focus();
		return false;
	}
	if (the.ChuQin.value=="") {
		document.getElementById('Tips').innerHTML = '����״������Ϊ��';
		the.ChuQin.focus();
		return false;
	}
	if (the.Jiaoshi.value=="") {
		document.getElementById('Tips').innerHTML = '�ڿν�ʦ����Ϊ��';
		the.Jiaoshi.focus();
		return false;
	}
	document.getElementById('Tips').innerHTML = '';
	return true;
}
function My_CheckField(the){
	if (the.value==""){
		document.getElementById('Tips').innerHTML = '�ֶβ���Ϊ��';
		the.focus();
		return false;
	}
	document.getElementById('Tips').innerHTML = '';
	return true;
}
</script>
</head>
<body>
<!--#include file="../include/banner.asp"-->
<!--#include file="left_banner.asp"-->
<%
Action=Request.Querystring("Action")
If Action = "AddRecord" Then
'��Ӽ�¼
	If Session("Admin")="" then
	'�ж��Ƿ��½
		Response.Redirect "jifang.asp"
		Response.End
	End If

Riqi=htmlencode(Request.form("Riqi"))
Jieci=htmlencode(Request.form("Jieci"))
Banji=htmlencode(Request.form("Banji"))
Jifang=htmlencode(Request.form("Jifang"))
Neirong=htmlencode(Request.form("Neirong"))
ChuQin=htmlencode(Request.form("ChuQin"))
Jiaoshi=htmlencode(Request.form("Jiaoshi"))
Beizhu=htmlencode(Request.form("Beizhu"))

	If (Len(Riqi)>0 And Len(Jieci)>0 And Len(Banji)>0 And Len(Jifang)>0 And Len(Neirong)>0 And _
		Len(ChuQin)>0 And Len(Jiaoshi)>0) then
		If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
		If isNull(Conn) Then
			Set Conn=Server.CreateObject("ADODB.Connection")
			My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
			Conn.Open My_conn_STRING
		End If
		Sql="Select * from Jifang where Riqi='" & Riqi & "' and Jieci ='" & Jieci& "' And Banji='" & Banji & "'"
		MyRs.open Sql,Conn,3,2
		If MyRs.recordcount>0 then
			Response.Write "<script>document.getElementById('Tips').innerHTML = '����������ÿ��Ѿ��������ˡ�';</SCRIPT>"
'			MyRs.close
		Else
			Sql="Insert Into [Jifang] (Riqi,Jieci,Banji,Jifang,Neirong,ChuQin,Jiaoshi,Beizhu) values ('"& Riqi &"','"& Jieci &"','"& Banji &"','"& Jifang &"','"& Neirong &"','"& ChuQin &"','" & Jiaoshi &"','" & Beizhu &"')"
			conn.execute(Sql)
			Response.Redirect "?Action=ShowJieci"
'			Response.End
		End If
	End If

	MyRs.Close
End If
%>
<div id="Right_Content" style="align:left;float:left">
<%If Session("Admin")<>"" then%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>��ӻ���ʹ�ü�¼</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Add_Area">
<form name="AddNewJieci" id="AddNewJieci" method="post" Action="?Action=AddRecord" onSubmit="return My_CheckFields(this);">
<span style="white-space: nowrap"><label for="Riqi">���ڣ�</label><input type="text" name="Riqi" id="Riqi" size="10" readonly="readonly" onclick="choose_date_czw('Riqi')"/></span>
<span style="white-space: nowrap"><label for="Jieci">�ڴΣ�</label>
<select name="Jieci" id="Jieci">
	<option value="���">���</option>
	<option value="��һ��">��һ��</option>
	<option value="�ڶ���">�ڶ���</option>
	<option value="������">������</option>
	<option value="���Ľ�">���Ľ�</option>
	<option value="���ʱ��">���ʱ��</option>
	<option value="����ʱ��">����ʱ��</option>
	<option value="�����">�����</option>
	<option value="������">������</option>
	<option value="���߽�">���߽�</option>
	<option value="�ڰ˽�">�ڰ˽�</option>
	<option value="������Ϣʱ��">������Ϣʱ��</option>
	<option value="����ϰ��һ��">����ϰ��һ��</option>
	<option value="����ϰ�ڶ���">����ϰ�ڶ���</option>
	<option value="����ϰ������">����ϰ������</option>
</select></span>
<span style="white-space: nowrap"><label for="Banji">�༶��</label><input type="text" name="Banji" value="" id="Banji" size="3" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="Jifang">������</label>
<select name="Jifang" id="Jifang">
<%
if getip="192.168.4.250" then
	response.write("	<option value=""1����"" selected=""selected"">1����</option>" & vbcrlf & "	<option value=""2����"">2����</option>" & vbcrlf & "	<option value=""3����"">3����</option>" & vbcrlf & "	<option value=""4����"">4����</option>")
else
'getip="192.168.4.252"
	response.write("	<option value=""1����"">1����</option>" & vbcrlf & "	<option value=""2����"">2����</option>" & vbcrlf & "	<option value=""3����"">3����</option>" & vbcrlf & "	<option value=""4����"" selected=""selected"">4����</option>")
end if
%>
	<option value="����">����</option>
</select></span>
<span style="white-space: nowrap"><label for="Neirong">���ݣ�</label><input type="text" name="Neirong" value="" id="Neirong" size="40" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="ChuQin">����״����</label><input type="text" name="ChuQin" value="����" id="ChuQin" size="3" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="Jiaoshi">�ڿν�ʦ��</label><input type="text" name="Jiaoshi" value=<%=Session("ShowName")%> id="Jiaoshi" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="Beizhu">��ע��</label><input type="text" name="Beizhu" value="" id="Beizhu" size="20"/></span>
<input type="submit" value="���" onClick="return My_CheckFields(this);"/>
</form>
</div>
<hr>
<script language="javascript">
var currentTime = new Date()
var month = currentTime.getMonth() + 1
var day = currentTime.getDate()
var year = currentTime.getFullYear()
var i = currentTime.getHours() -7

document.getElementById('Riqi').value = year + "-" + month + "-" + day;
document.getElementById("Jieci").options[i].selected = true;
</script>
<%End If%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>����ʹ�����һ����</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Search_Area">
<form id="SearchJieci" name="SearchJieci" method="post" Action="?Action=Search" onSubmit="return My_CheckSearchDates(this);">
<span style="white-space: nowrap"><label for="S_Riqi">���ڣ���</label><input type="text" name="S_Riqi" id="S_Riqi" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/><label for="S_Riqi_2">��</label><input type="text" name="S_Riqi_2" id="S_Riqi_2" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/></span>
<span style="white-space: nowrap"><label for="S_Jieci">�ڴΣ�</label>
<select name="S_Jieci" id="S_Jieci">
	<option value="%" Selected="Selected">����</option>
	<option value="���">���</option>
	<option value="��һ��">��һ��</option>
	<option value="�ڶ���">�ڶ���</option>
	<option value="������">������</option>
	<option value="���Ľ�">���Ľ�</option>
	<option value="���ʱ��">���ʱ��</option>
	<option value="����ʱ��">����ʱ��</option>
	<option value="�����">�����</option>
	<option value="������">������</option>
	<option value="���߽�">���߽�</option>
	<option value="�ڰ˽�">�ڰ˽�</option>
	<option value="������Ϣʱ��">������Ϣʱ��</option>
	<option value="����ϰ��һ��">����ϰ��һ��</option>
	<option value="����ϰ�ڶ���">����ϰ�ڶ���</option>
	<option value="����ϰ������">����ϰ������</option>
</select></span>
<span style="white-space: nowrap"><label for="S_Banji">�༶��</label><input name="S_Banji" id="S_Banji" type="text" value="" size="5"/></span>
<span style="white-space: nowrap"><label for="S_Jifang">������</label>
<select name="S_Jifang" id="S_Jifang">
	<option value="%" Selected="Selected">ȫ��</option>
	<option value="1����">1����</option>
	<option value="2����">4����</option>
	<option value="3����">4����</option>
	<option value="4����">4����</option>
	<option value="����">����</option>
</select></span>
<span style="white-space: nowrap"><label for="S_Neirong">���ݣ�</label><input name="S_Neirong" id="S_Neirong" type="text" value="" size="30"/></span>
<span style="white-space: nowrap"><label for="S_Chuqin">���ڣ�</label>
<select name="S_Chuqin" id="S_Chuqin">
	<option value="-1" Selected="Selected">����ν</option>
	<option value="0">����</option>
	<option value="1">δ����</option>
</select></span>
<span style="white-space: nowrap"><label for="S_Jiaoshi">�ڿν�ʦ��</label><input name="S_Jiaoshi" id="S_Jiaoshi" type="text" value="" size="5"/></span>
<input type="submit" value="����" name="S_Submit" id="S_Submit"/>
</form>
</div>
<%
'������¼
S_Riqi=htmlencode(Request.form("S_Riqi"))
S_Riqi_2=htmlencode(Request.form("S_Riqi_2"))
S_Jieci=htmlencode(Request.form("S_Jieci"))
S_Banji=htmlencode(Request.form("S_Banji"))
S_Jifang=htmlencode(Request.form("S_Jifang"))
S_Neirong=htmlencode(Request.form("S_Neirong"))
S_Chuqin=Request.form("S_Chuqin")
S_Jiaoshi=htmlencode(Request.form("S_Jiaoshi"))

SQL="select * from Jifang where 1=1"
If Len(S_Riqi)<>0 AND Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi &"' and '" & S_Riqi_2 &"'"
If Len(S_Riqi)<>0 XOR Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi & S_Riqi_2 &"' and '" & S_Riqi & S_Riqi_2 &"'"
If Len(S_Jieci)<>0 Then SQL = SQL & " and Jieci Like '" & S_Jieci &"'"
If Len(S_Banji)<>0 Then SQL = SQL & " and Banji Like '" & S_Banji &"'"
If Len(S_Jifang)<>0 Then SQL = SQL & " and Jifang Like '" & S_Jifang &"'"
If Len(S_Neirong)<>0 Then SQL = SQL & " and Neirong='" & S_Neirong &"'"
'If Len(S_Chuqin)<>0 Then SQL = SQL & " and ChuQin='" & S_Chuqin &"'"
If S_Chuqin="0" Then
'����
	SQL = SQL & " and ChuQin='����'"
ElseIf S_ChuQin="1" Then
'δ����
	SQL = SQL & " and ChuQin <>'����'"
Else
'����ν
	SQL = SQL
End If
If Len(S_Jiaoshi)<>0 Then SQL = SQL & " and Jiaoshi='" & S_Jiaoshi &"'"
SQL = SQL & " order by Riqi desc,Jieci desc,Jifang desc"

PageSize=20
MyRs.open Sql,Conn,3,2
MyRs.PageSize=PageSize

ResultCount=MyRs.recordcount

If ResultCount>0 then
%>
<table align="left" width="100%">
<thead>
<tr class="odd">
<% 'Put Headings On The Table of Field Names
howmanyfields=MyRs.fields.count -1 

for i=0 to howmanyfields
	Select Case UCase(MyRs(i).Name)
		Case "RIQI":
			response.Write "<th><b>" & "����" & "</b></th>"
		Case "JIECI":
			response.Write "<th><b>" & "�ڴ�" & "</b></th>"
		Case "BANJI":
			response.Write "<th><b>" & "�༶" & "</b></th>"
		Case "JIFANG":
			response.Write "<th width='30px'><b>" & "����" & "</b></th>"
		Case "NEIRONG":
			response.Write "<th class='NeiRong'><b>" & "����" & "</b></th>"
		Case "CHUQIN":
			response.Write "<th width='60px'><b>" & "����״��" & "</b></th>"
		Case "JIAOSHI":
			response.Write "<th width='60px'><b>" & "�ڿν�ʦ" & "</b></th>"
		Case "BEIZHU":
			response.Write "<th class='BeiZhu'><b>" & "��ע" & "</b></th>"
		Case Else
			
	End Select
next %>
</tr>
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
For i_s = 1 to ShowPage
	If MyRs.EOF Then Exit For
	if 1 = i_s mod 2 then
		response.write("<tr id='Data' class='odd'>")
	else
		response.write("<tr id='Data'>")
	end if
	for i_c = 1 to howmanyfields '����ʾId�ֶ�
		ThisRecord = MyRs(i_c)
		If IsNull(ThisRecord) Then
			ThisRecord = "&nbsp;"
		end if
		If Ucase(MyRs(i_c).Name)="BEIZHU" Then
			Response.write("<td class='BeiZhu'>" & ThisRecord & "</td>")
		Else
			Response.write("<td>" & ThisRecord & "</td>")
		End If
	next
	response.write("</tr>")
	MyRs.movenext
Next
'j=0
'do while not MyRs.eof
'j=j+1
'if 1 = j mod 2 then
'	Response.write("<tr id='Data' class='odd'>")
'else
'	Response.write("<tr id='Data'>")
'end if

'for i = 1 to howmanyfields '����ʾId�ֶ�
'	ThisRecord = MyRs(i)
'	If IsNull(ThisRecord) Then
'		ThisRecord = "&nbsp;"
'	end if
'	If Ucase(MyRs(i).Name)="BEIZHU" Then
'		Response.write("<td class='BeiZhu'>" & ThisRecord & "</td>")
'	Else
'		Response.write("<td>" & ThisRecord & "</td>")
'	End If
'next
'Response.write("</tr>")
'MyRs.movenext
'loop
%>
</tbody>
</table>
<div id="navbar" align="left">
<br clear="left">
<%
response.write "���ҳ�룺"
PageCount=-Int(-ResultCount/(PageSize+1))
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

response.write "<a href=""jifang.asp?page=1"">��һҳ</a> "
for i=StartPage to EndPage
	response.write "<a href=""jifang.asp?page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""jifang.asp?page=" & PageCount & """>���ҳ</a> "
Else
	Response.Write "<h1>û���ҵ��κν��������Ĺؼ��ʣ�������������</h1>"
End If
'End Select
'MyRs.close
'Set MyRs= Nothing
Conn.Close
set Conn=nothing
%>
<p align="left"><a href="download.asp?Action=jifang">���ػ���ʹ������</a></p>
</div>
</div>
<br clear=all>
<br clear=all>
<!--#include file="../include/bottom.asp"-->
<script language="javascript">
/*����ȡ��ֵ�浽������*/
c_Width =document.body.offsetWidth - parseInt(document.getElementById('Left_Banner').style.width)-40;
document.getElementById('Right_Content').style.width="" +c_Width +'px';
document.getElementById('Right_Content').style.maxWidth="" + c_Width  + 'px';

window.onresize=function(){
    document.getElementById("Right_Content").style.width=(function(){
        var x=document.body.offsetWidth - parseInt(document.getElementById('Left_Banner').style.width)-40;
        return x;
    })();
}

</script>
</body>
</html>