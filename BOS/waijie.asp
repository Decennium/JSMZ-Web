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
<title>�豸������ϵͳ</title>
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
	if (the.JieQi.value=="") {
		document.getElementById('Tips').innerHTML = '���ڲ���Ϊ��';
		the.JieQi.focus();
		return false;
	}
	if (!My_IsInt(the.JieQi.value)){
		document.getElementById('Tips').innerHTML = '���ڱ���Ϊ��Ȼ��';
		the.JieQi.focus();
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
		Response.Redirect "waijie.asp"
		Response.End
	End If

Riqi=htmlencode(Request.form("Riqi"))
ShenQingRen=htmlencode(Request.form("ShenQingRen"))
SheBei=htmlencode(Request.form("SheBei"))
JieQi=htmlencode(Request.form("JieQi"))
MiaoShu=htmlencode(Request.form("MiaoShu"))
FaFangRen=htmlencode(Request.form("FaFangRen"))
Beizhu=htmlencode(Request.form("Beizhu"))

	If (Len(Riqi)>0 And Len(ShenQingRen)>0 And Len(SheBei)>0 And Len(JieQi)>0 And Len(MiaoShu)>0 And Len(FaFangRen)>0) then
		If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
		If isNull(Conn) Then
			Set Conn=Server.CreateObject("ADODB.Connection")
			My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
			Conn.Open My_conn_STRING
		End If
		Sql="Select * from WaiJie where SheBei='" & SheBei & "' and GuiHuan = 0"
		MyRs.open Sql,Conn,3,2
		If MyRs.recordcount>0 then
			Response.Write "<script>document.getElementById('Tips').innerHTML = '�豸" & MyRs(3).value & "�Ѿ���" & MyRs(2).Value & "�������δ�黹��';</SCRIPT>"
'			MyRs.close
		Else
			Sql="INSERT INTO [WaiJie] ([RiQi],[ShenQingRen],[SheBei],[JieQi],[MiaoShu],[FaFangRen],[GuiHuan],[BeiZhu]) VALUES ('"& RiQi &"','"& ShenQingRen &"','"& SheBei &"','"& JieQi &"','"& MiaoShu &"','" & FaFangRen &"',0,'" & Beizhu &"')"
			conn.execute(Sql)
			Response.Redirect "?Action=ShowJieci"
'			Response.End
		End If
	End If

	MyRs.Close
End If
If Action = "AddCheck" Then
'���Ӽ��
	If Session("Admin")="" then
	'�ж��Ƿ��½
		Response.Redirect "waijie.asp"
		Response.End
	End If
ID=htmlencode(Request.form("id"))
GuiHuanRiQi=htmlencode(Request.form("GuiHuanRiQi"))
ZhuangKuang=htmlencode(Request.form("ZhuangKuang"))
QianShouRen=htmlencode(Request.form("QianShouRen"))
	If (Len(ID)>0 And Len(GuiHuanRiQi)>0 And Len(QianShouRen)>0) then
		If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
		If isNull(Conn) Then
			Set Conn=Server.CreateObject("ADODB.Connection")
			My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
			Conn.Open My_conn_STRING
		End If
			Sql="UPDATE [WaiJie] SET [GuiHuan] = 1,[GuiHuanRiQi] ='"& GuiHuanRiQi &"',[QianShouRen] ='"& QianShouRen &"',[ZhuangKuang] ='"& ZhuangKuang &"' WHERE [id] ='"& ID &"'"
			conn.execute(Sql)
			Response.Redirect "?Action=ShowJieci"
	End If
End If
%>
<div id="Right_Content" style="align:left;float:left">
<%If Session("Admin")<>"" then%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>�����豸����¼</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Add_Area">
<form name="AddNewJieci" id="AddNewJieci" method="post" Action="?Action=AddRecord" onSubmit="return My_CheckFields(this);">
<span style="white-space: nowrap"><label for="Riqi">���ڣ�</label><input type="text" name="Riqi" id="Riqi" size="10" readonly="readonly" onclick="choose_date_czw('Riqi')"/></span>
<span style="white-space: nowrap"><label for="ShenQingRen">�����ˣ�</label><input type="text" name="ShenQingRen" value="" id="ShenQingRen" size="5" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="SheBei">�豸��</label><input type="text" name="SheBei" value="" id="SheBei" size="10" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="JieQi">���ڣ�</label><input type="text" name="JieQi" value="" id="JieQi" size="5" onblur="return My_CheckField(this);"/><label for="JieQi">�죬</label></span>
<span style="white-space: nowrap"><label for="MiaoShu">�豸������</label><input type="text" name="MiaoShu" value="һ������" id="MiaoShu" size="30" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="FaFangRen">�����ˣ�</label><input type="text" name="FaFangRen" value='<%=Session("ShowName")%>' id="FaFangRen" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="Beizhu">��ע��</label><input type="text" name="Beizhu" value="" id="Beizhu" size="20"/></span>
<input type="submit" value="����" onClick="return My_CheckFields(this);"/>
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
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>�豸������һ����</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Search_Area">
<form id="SearchJieci" name="SearchJieci" method="post" Action="?Action=Search" onSubmit="return My_CheckSearchDates(this);">
<span style="white-space: nowrap"><label for="S_Riqi">���ڣ���</label><input type="text" name="S_Riqi" id="S_Riqi" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/><label for="S_Riqi_2">��</label><input type="text" name="S_Riqi_2" id="S_Riqi_2" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/></span>
<span style="white-space: nowrap"><label for="S_ShenQingRen">�����ˣ�</label><input name="S_ShenQingRen" id="S_ShenQingRen" type="text" value="" size="5"/></span>
<span style="white-space: nowrap"><label for="S_SheBei">�豸��</label><input name="S_SheBei" id="S_SheBei" type="text" value="" size="10"/></span>
<span style="white-space: nowrap"><label for="S_GuiHuan">�Ƿ�黹��</label>
<select name="S_GuiHuan" id="S_GuiHuan">
	<option value="%">����</option>
	<option value="0" Selected="Selected">δ�黹</option>
	<option value="1">�ѹ黹</option>
</select></span>
<span style="white-space: nowrap"><label for="S_FaFangRen">�����ˣ�</label><input name="S_FaFangRen" id="S_FaFangRen" type="text" value="" size="5"/></span>
<input type="submit" value="����" name="S_Submit" id="S_Submit"/>
</form>
</div>
<%
'������¼
S_Riqi=htmlencode(Request.form("S_Riqi"))
S_Riqi_2=htmlencode(Request.form("S_Riqi_2"))
S_ShenQingRen=htmlencode(Request.form("S_ShenQingRen"))
S_SheBei=htmlencode(Request.form("S_SheBei"))
S_GuiHuan=htmlencode(Request.form("S_GuiHuan"))
S_FaFangRen=htmlencode(Request.form("S_FaFangRen"))

SQL="select * from WaiJie where 1=1"
If Len(S_Riqi)<>0 AND Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi &"' and '" & S_Riqi_2 &"'"
If Len(S_Riqi)<>0 XOR Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi & S_Riqi_2 &"' and '" & S_Riqi & S_Riqi_2 &"'"
If Len(S_ShenQingRen)<>0 Then SQL = SQL & " and ShenQingRen Like '%" & S_ShenQingRen &"%'"
If Len(S_SheBei)<>0 Then SQL = SQL & " and SheBei Like '%" & S_SheBei &"%'"
If Len(S_GuiHuan)<>0 Then SQL = SQL & " and GuiHuan = " & S_GuiHuan &""
If Len(S_FaFangRen)<>0 Then SQL = SQL & " and FaFangRen = '" & S_FaFangRen &"'"
SQL = SQL & " order by Riqi desc, SheBei desc"
'response.write sql
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
		Case "SHENQINGREN":
			response.Write "<th><b>" & "������" & "</b></th>"
		Case "SHEBEI":
			response.Write "<th><b>" & "�豸" & "</b></th>"
		Case "JIEQI":
			response.Write "<th width='50px'><b>" & "����/��" & "</b></th>"
		Case "MIAOSHU":
			response.Write "<th class='NeiRong'><b>" & "����" & "</b></th>"
		Case "FAFANGREN":
			response.Write "<th width='60px'><b>" & "������" & "</b></th>"
		Case "GUIHUAN":
			response.Write "<th width='60px'><b>" & "�黹���" & "</b></th>"
		Case "GUIHUANRIQI":
			response.Write "<th width='60px'><b>" & "�黹����" & "</b></th>"
		Case "ZHUANGKUANG":
			response.Write "<th class='NeiRong'><b>" & "�黹ʱ״��" & "</b></th>"
		Case "QIANSHOUREN":
			response.Write "<th width='60px'><b>" & "ǩ����" & "</b></th>"
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

For i_s = 1 to ShowPage
	If MyRs.EOF Then Exit For
	if 1 = i_s mod 2 then
		response.write("<tr id='Data' class='odd'>")
	else
		response.write("<tr id='Data'>")
	end if
	for i_c = 1 to howmanyfields '����ʾId�ֶ�
		ThisRecord = MyRs(i_c).Value
		If IsNull(ThisRecord) Then
			ThisRecord = ""
		End if
		If Ucase(MyRs(i_c).Name)="GUIHUANRIQI" Then
			If ThisRecord = "" Then
				Response.write("<form name='AddCheck' id='AddCheck' method='post' Action='?Action=AddCheck'><td><input type='hidden' name='id' value='" & MyRs(0).Value & "'/><input type=""text"" name=""GuiHuanRiQi"" id=""GuiHuanRiQi"" value=""" & year(now()) &"-" & month(now()) & "-" & day(now()) & """ size=""10"" readonly=""readonly"" onclick=""choose_date_czw('GuiHuanRiQi')""/><input type='submit' value='���'/></td>")
			Else
				Response.write("<td>" & ThisRecord & "</td>")
			End If
		ElseIf Ucase(MyRs(i_c).Name)="RIQI" Then
			Response.write("<td width='80px'>" & ThisRecord & "</td>")
		ElseIf Ucase(MyRs(i_c).Name)="SHENQINGREN" Then
			Response.write("<td width='80px'>" & ThisRecord & "</td>")
		ElseIf Ucase(MyRs(i_c).Name)="JIEQI" Then
			Response.write("<td style='text-align:right;padding-right:10px'>" & ThisRecord & "</td>")
		ElseIf Ucase(MyRs(i_c).Name)="JIEQI" Then
			Response.write("<td style='text-align:right;padding-right:10px'>" & ThisRecord & "</td>")
		ElseIf Ucase(MyRs(i_c).Name)="GUIHUAN" Then
			If ThisRecord  Then
				Response.write("<td>�ѹ黹</td>")
			Else
				Response.write("<td>δ�黹</td>")
			End If
		ElseIf Ucase(MyRs(i_c).Name)="ZHUANGKUANG" Then
			If ThisRecord = "" Then
				Response.write("<td><input type='text' name='ZhuangKuang' value='һ������' size='30'/></td>")
			Else
				Response.write("<td>" & ThisRecord & "</td>")
			End If
		ElseIf Ucase(MyRs(i_c).Name)="QIANSHOUREN" Then
			If ThisRecord = "" Then
				Response.write("<td><input type='text' name='QianSHouRen' value='" & Session("ShowName") & "'  size='5'/></form></td>")
			Else
				Response.write("<td>" & ThisRecord & "</td>")
			End If
		ElseIf Ucase(MyRs(i_c).Name)="BEIZHU" Then
			Response.write("<td class='BeiZhu'>" & ThisRecord & "</td>")
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

response.write "<a href=""waijie.asp?page=1"">��һҳ</a> "
for i=StartPage to EndPage
	response.write "<a href=""waijie.asp?page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""waijie.asp?page=" & PageCount & """>���ҳ</a> "
Else
	Response.Write "<h1>û���ҵ��κν��������Ĺؼ��ʣ�������������</h1>"
End If
'End Select
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