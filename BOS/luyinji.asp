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
<title>��ȡ¼��������ϵͳ</title>
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
	if (the.JiaoShi.value=="") {
		document.getElementById('Tips').innerHTML = '��ȡ¼�����Ľ�ʦ����Ϊ��';
		the.JiaoShi.focus();
		return false;
	}
	if (the.XueKe.value=="") {
		document.getElementById('Tips').innerHTML = '��ʦ����ѧ�Ʋ���Ϊ��';
		the.XueKe.focus();
		return false;
	}
	if (the.FaFangRen.value=="") {
		document.getElementById('Tips').innerHTML = '¼���������˲���Ϊ��';
		the.FaFangRen.focus();
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
'���Ӽ�¼
	If Session("Admin")="" then
	'�ж��Ƿ��½
		Response.Redirect "luyinji.asp"
		Response.End
	End If

RiQi=htmlencode(Request.form("RiQi"))
JiaoShi=htmlencode(Request.form("JiaoShi"))
XueKe=htmlencode(Request.form("XueKe"))
FaFangRen=htmlencode(Request.form("FaFangRen"))
Beizhu=htmlencode(Request.form("Beizhu"))

	If (Len(RiQi)>0 And Len(JiaoShi)>0 And Len(XueKe)>0 And Len(FaFangRen)>0) Then
		If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
		If isNull(Conn) Then
			Set Conn=Server.CreateObject("ADODB.Connection")
			My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
			Conn.Open My_conn_STRING
		End If
		Sql="Select top 1 * from LuYinJi where JiaoShi='" & JiaoShi & "' Order By RiQi DESC"
		MyRs.open Sql,Conn,3,2
		If MyRs.recordcount>0 then
			If abs(DateDiff("m",MyRs(1).Value,Now()))<35 Then
				Response.Write "<script language='javascript'>document.getElementById(""Tips"").innerHTML = ""��λ��ʦ��һ̨¼������û�ù����ꡣ"";</script>"
				'Response.Write "<script language='javascript'>alert('��λ��ʦ��һ̨¼������û�ù����ꡣ');</script>"
			Else
				Sql="INSERT INTO [LuYinJi]([RiQi],[JiaoShi],[XueKe],[FaFangRen],[BeiZhu]) " & _
					"VALUES ('"& RiQi &"','"& JiaoShi &"','"& XueKe &"','"& FaFangRen &"','" & Beizhu &"')"
				conn.execute(Sql)
				'Response.Redirect "?Action=ShowSheBei"
			End If
		Else
			Sql="INSERT INTO [LuYinJi]([RiQi],[JiaoShi],[XueKe],[FaFangRen],[BeiZhu]) " & _
				"VALUES ('"& RiQi &"','"& JiaoShi &"','"& XueKe &"','"& FaFangRen &"','" & Beizhu &"')"
			conn.execute(Sql)
			'Response.Redirect "?Action=ShowSheBei"
		End If
	End If

	MyRs.Close
End If
%>
<div id="Right_Content" style="align:left;float:left">
<%If Session("Admin")<>"" then%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>������ȡ¼������¼</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Add_Area">
<form name="AddNewLingQu" id="AddNewLingQu" method="post" Action="?Action=AddRecord" onSubmit="return My_CheckFields(this);">
<span style="white-space: nowrap"><label for="Riqi">���ڣ�</label><input type="text" name="Riqi" id="Riqi" size="10" readonly="readonly" onclick="choose_date_czw('Riqi')"/></span>
<span style="white-space: nowrap"><label for="JiaoShi">��ʦ��</label><input type="text" name="JiaoShi" value="" id="JiaoShi" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="XueKe">ѧ�ƣ�</label>
<select name="XueKe" id="XueKe">
	<option value="Ӣ��">Ӣ��</option>
	<option value="����">����</option>
	<option value="�赸">�赸</option>
	<option value="����">����</option>
	<option value="����ѧ��">����ѧ��</option>
</select></span>
<span style="white-space: nowrap"><label for="FaFangRen">�����ˣ�</label><input type="text" name="FaFangRen" value='<%=Session("ShowName")%>' id="FaFangRen" size="10" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="Beizhu">��ע��</label><input type="text" name="Beizhu" value="" id="Beizhu" size="30"/></span>
<input type="submit" value="����"/>
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
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>��ȡ¼�������һ����</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Search_Area">
<form id="SearchSheBei" name="SearchSheBei" method="post" Action="?Action=Search" onSubmit="return My_CheckSearchDates(this);">
<span style="white-space: nowrap"><label for="S_Riqi">���ڣ���</label><input type="text" name="S_Riqi" id="S_Riqi" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/><label for="S_Riqi_2">��</label><input type="text" name="S_Riqi_2" id="S_Riqi_2" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/></span>
<span style="white-space: nowrap"><label for="S_JiaoShi">��ʦ��</label><input type="text" name="S_JiaoShi" value="" id="S_JiaoShi" size="40" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_XueKe">ѧ�ƣ�</label>
<select name="S_XueKe" id="S_XueKe">
	<option value="%">����ѧ��</option>
	<option value="Ӣ��">Ӣ��</option>
	<option value="����">����</option>
	<option value="�赸">�赸</option>
	<option value="����">����</option>
	<option value="����ѧ��">����ѧ��</option>
</select></span>
<span style="white-space: nowrap"><label for="S_FaFangRen">�����ˣ�</label><input type="text" name="S_FaFangRen" value="" id="S_FaFangRen" size="10" onblur="return My_CheckField(this);"/></span>
<input type="submit" value="����" name="S_Submit" id="S_Submit"/>
</form>
</div>
<%
'������¼

S_Riqi=htmlencode(Request.form("S_Riqi"))
S_Riqi_2=htmlencode(Request.form("S_Riqi_2"))
S_JiaoShi=htmlencode(Request.form("S_JiaoShi"))
S_XueKe=htmlencode(Request.form("S_XueKe"))
S_FaFangRen=htmlencode(Request.form("S_FaFangRen"))

SQL="select * from LuYinJi where 1=1"
If Len(S_Riqi)<>0 AND Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi &"' and '" & S_Riqi_2 &"'"
If Len(S_Riqi)<>0 XOR Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi & S_Riqi_2 &"' and '" & S_Riqi & S_Riqi_2 &"'"
If Len(S_JiaoShi)<>0 Then SQL = SQL & " and JiaoShi Like '" & S_JiaoShi &"'"
If Len(S_XueKe)<>0 Then SQL = SQL & " and XueKe Like '" & S_XueKe &"'"
If Len(S_FaFangRen)<>0 Then SQL = SQL & " and FaFangRen Like '" & S_FaFangRen &"'"
SQL = SQL & " order by Riqi desc, JiaoShi desc"

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
		Case "JIAOSHI":
			response.Write "<th><b>" & "��ʦ" & "</b></th>"
		Case "XUEKE":
			response.Write "<th><b>" & "ѧ��" & "</b></th>"
		Case "FAFANGREN":
			response.Write "<th><b>" & "������" & "</b></th>"
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
'End Select
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