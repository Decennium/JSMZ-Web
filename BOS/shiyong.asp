<%
'���˴���
Function htmlencode(fString)
If not isnull(fString) then
	fString = trim(fString)
	fString = Replace(fString, ">", "&gt;")
	fString = Replace(fString, "<", "&lt;")
	fString = Replace(fString, CHR(32), "&nbsp;")
	fString = Replace(fString, CHR(9), "&nbsp;")
	fString = Replace(fString, CHR(34), "&quot;")
	fString = Replace(fString, CHR(39), "&#39;")
	fString = Replace(fString, CHR(13) & CHR(10), "</p><p>")
	fString = Replace(fString, CHR(10) & CHR(10), "</p><p>")
	fString = Replace(fString, CHR(10), "<br>")
	htmlencode = fString
End If
End Function

Function getIP() 
Dim strIPAddr 

If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
	strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
	strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
	strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1) 
Else 
	strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
End If 

getIP = Trim(Mid(strIPAddr, 1, 30)) 
End Function 

Response.Charset = "gb2312"
Response.Buffer = True

'���ݿ�����
dim conn,connstr
'on error resume next

Currentpage=request("page")
If Currentpage < 1 Then Currentpage = 1

Set MyRs = Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")

My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
Conn.Open My_conn_STRING
%>
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>�豸ʹ�ù���ϵͳ</title>
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
	if (the.ShiChang.value=="") {
		document.getElementById('Tips').innerHTML = 'ʱ������Ϊ��';
		the.ShiChang.focus();
		return false;
	}
	if (the.YongTu.value=="") {
		document.getElementById('Tips').innerHTML = '��;����Ϊ��';
		the.YongTu.focus();
		return false;
	}
	if (the.ShiYongRen.value=="") {
		document.getElementById('Tips').innerHTML = 'ʹ���˲���Ϊ��';
		the.ShiYongRen.focus();
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
Select Case Action
'��Ӽ�¼
Case "AddShiYong"
	'�ж��Ƿ��½
	If Session("Admin")="" then
		Response.Redirect "shiyong.asp"
		Response.End
	End If

SheBei=htmlencode(Request.form("SheBei"))
ShiYongRen=htmlencode(Request.form("ShiYongRen"))
RiQi=htmlencode(Request.form("RiQi"))
ShiChang=htmlencode(Request.form("ShiChang"))
YongTu=htmlencode(Request.form("YongTu"))
Beizhu=htmlencode(Request.form("Beizhu"))

	If (Len(SheBei)>0 And Len(ShiYongRen)>0 And Len(RiQi)>0 And Len(ShiChang)>0 And Len(YongTu)>0) Then
		If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
		If isNull(Conn) Then
			Set Conn=Server.CreateObject("ADODB.Connection")
			My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
			Conn.Open My_conn_STRING
		End If
		Sql="INSERT INTO [ShiYong]([SheBei],[ShiYongRen],[RiQi],[ShiChang],[YongTu],[beiZhu]) VALUES ('"& SheBei &"','"& ShiYongRen &"','"& RiQi &"','"& ShiChang &"','"& YongTu &"','" & Beizhu &"')"
		conn.execute(Sql)
		Response.Redirect "?Action=ShowSheBei"
'			Response.End
	End If
'	response.write "<p>" & SQL & "</p>"
Case Else
%>
<div id="Right_Content" style="align:left;float:left">
<%If Session("Admin")<>"" then%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>����豸ʹ�ü�¼</strong></div>
<div id="Tips2" style="float:left;color:red"></div>
<br clear="all"/>
<div align="left" clear="all" id="Add_Area">
<form name="AddNewShiYong" id="AddNewShiYong" method="post" Action="?Action=AddShiYong" onSubmit="return My_CheckFields(this);">
<span style="white-space: nowrap"><label for="Shebei">�豸��</label>
<select name="Shebei" id="Shebei">
	<option value="��¼��">��¼��</option>
	<option value="�����">�����</option>
	<option value="С�����">С�����</option>
	<option value="�������">�������</option>
	<option value="���칦��">���칦��</option>
	<option value="�������ȹ���">�������ȹ���</option>
	<option value="�����豸">�����豸</option>
</select></span>
<span style="white-space: nowrap"><label for="Riqi">���ڣ�</label><input type="text" name="Riqi" id="Riqi" size="10" readonly="readonly" onclick="choose_date_czw('Riqi')"/></span>
<span style="white-space: nowrap"><label for="ShiChang">ʱ����</label><input type="text" name="ShiChang" value="" id="ShiChang" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="YongTu">��;��</label><input type="text" name="YongTu" value="" id="YongTu" size="40" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="ShiYongRen">ʹ���ˣ�</label><input type="text" name="ShiYongRen" value=<%=Session("ShowName")%> id="ShiYongRen" size="10" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="Beizhu">��ע��</label><input type="text" name="Beizhu" value="" id="Beizhu" size="30"/></span>
<input type="submit" value="���" onClick="return My_CheckFields(this);"/>
</form>
</div>
<hr style="height:1px;border:none;border-top:1px solid #e5eff8;">
<script language="javascript">
var currentTime = new Date()
var month = currentTime.getMonth() + 1
var day = currentTime.getDate()
var year = currentTime.getFullYear()
//var i = currentTime.getHours() -7

document.getElementById('Riqi').value = year + "-" + month + "-" + day;
//document.getElementById("Jieci").options[i].selected = true;
</script>
<%End If%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>�豸ʹ�����һ����</strong></div>
<div id="Tips2" style="float:left;color:red"></div>
<br clear="all"/>
<div align="left" clear="all" id="Search_Area">
<form id="SearchSheBei" name="SearchSheBei" method="post" Action="?Action=Search" onSubmit="return My_CheckSearchDates(this);">
<span style="white-space: nowrap"><label for="S_Riqi">���ڣ���</label><input type="text" name="S_Riqi" id="S_Riqi" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/><label for="S_Riqi_2">��</label><input type="text" name="S_Riqi_2" id="S_Riqi_2" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/></span>
<span style="white-space: nowrap"><label for="S_Shebei">�豸��</label>
<select name="S_Shebei" id="S_Shebei">
	<option value="%">�����豸</option>
	<option value="��¼��">��¼��</option>
	<option value="�����">�����</option>
	<option value="С�����">С�����</option>
	<option value="�������">�������</option>
	<option value="���칦��">���칦��</option>
	<option value="�������ȹ���">�������ȹ���</option>
	<option value="�����豸">�����豸</option>
</select></span>
<span style="white-space: nowrap"><label for="S_YongTu">��;��</label><input type="text" name="S_YongTu" value="" id="S_YongTu" size="40" title="�����벿�ֹؼ���" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_ShiYongRen">ʹ���ˣ�</label><input type="text" name="S_ShiYongRen" value="" id="S_ShiYongRen" size="10" onblur="return My_CheckField(this);"/></span>
<input type="submit" value="����" name="S_Submit" id="S_Submit"/>
</form>
</div>
<%
'������¼

S_Riqi=htmlencode(Request.form("S_Riqi"))
S_Riqi_2=htmlencode(Request.form("S_Riqi_2"))
S_SheBei=htmlencode(Request.form("S_SheBei"))
S_ShiYongRen=htmlencode(Request.form("S_ShiYongRen"))
S_YongTu=htmlencode(Request.form("S_YongTu"))

SQL="select * from ShiYong where 1=1"
If Len(S_Riqi)<>0 AND Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi &"' and '" & S_Riqi_2 &"'"
If Len(S_Riqi)<>0 XOR Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi & S_Riqi_2 &"' and '" & S_Riqi & S_Riqi_2 &"'"
If Len(S_SheBei)<>0 Then SQL = SQL & " and SheBei Like '" & S_Jieci &"'"
If Len(S_ShiYongRen)<>0 Then SQL = SQL & " and ShiYongRen Like '" & S_Banji &"'"
If Len(S_YongTu)<>0 Then SQL = SQL & " and YongTu Like '%" & S_Jiaoshi &"%'"
SQL = SQL & " order by SheBei desc, Riqi desc, ShiYongRen desc"

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
		Case "SHEBEI":
			response.Write "<th><b>" & "�豸" & "</b></th>"
		Case "RIQI":
			response.Write "<th><b>" & "����" & "</b></th>"
		Case "SHICHANG":
			response.Write "<th><b>" & "ʱ��" & "</b></th>"
		Case "SHIYONGREN":
			response.Write "<th><b>" & "ʹ����" & "</b></th>"
		Case "YONGTU":
			response.Write "<th width='300px'><b>" & "��;" & "</b></th>"
		Case "BEIZHU":
			response.Write "<th width='300px'><b>" & "��ע" & "</b></th>"
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
			Response.write("<td style='max-width:300px;word-wrap: break-word;'>" & ThisRecord & "</td>")
		Else
			Response.write("<td>" & ThisRecord & "</td>")
		End If
	next
	response.write("</tr>")
	MyRs.movenext
Next
%>
</tbody>
</table>
<div id="navbar" align="left">
<br clear="left">
<%
response.write "���ҳ�룺"
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

response.write "<a href=""jifang.asp?page=1"">��һҳ</a> "
for i=StartPage to EndPage
	response.write "<a href=""jifang.asp?page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""jifang.asp?page=" & PageCount & """>���ҳ</a> "
Else
	Response.Write "<h1>û���ҵ��κν��������Ĺؼ��ʣ�������������</h1>"
End If
End Select
'MyRs.close
'Set MyRs= Nothing
Conn.Close
set Conn=nothing
%>
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