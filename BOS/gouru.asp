<!--#include file="top.asp"-->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>�豸�������ϵͳ</title>
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
function My_IsNature(val) {
	return /^[1-9][0-9]*$/.test(val);
}
function My_IsFloat(val) {
	return /^\d+(\.\d+)?$/.test(val);
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
	if (the.SheBei.value=="") {
		document.getElementById('Tips').innerHTML = '�豸���Ʋ���Ϊ��';
		the.SheBei.focus();
		return false;
	}
	if (the.PinPai.value=="") {
		document.getElementById('Tips').innerHTML = '�豸Ʒ�Ʋ���Ϊ��';
		the.PinPai.focus();
		return false;
	}
	if (the.ShuLiang.value=="") {
		document.getElementById('Tips').innerHTML = '��������Ϊ��';
		the.ShuLiang.focus();
		return false;
	}
	if (!My_IsNature(the.ShuLiang.value)){
		document.getElementById('Tips').innerHTML = '��������Ϊ��Ȼ��';
		the.ShuLiang.focus();
		return false;
	}
	if (the.DanJia.value=="") {
		document.getElementById('Tips').innerHTML = '���۲���Ϊ��';
		the.DanJia.focus();
		return false;
	}
	if (!My_IsFloat(the.DanJia.value)){
		document.getElementById('Tips').innerHTML = '���۱���Ϊ����';
		the.DanJia.focus();
		return false;
	}
	if (the.YongTu.value=="") {
		document.getElementById('Tips').innerHTML = '��;����Ϊ��';
		the.YongTu.focus();
		return false;
	}
	if (the.JingShouRen.value=="") {
		document.getElementById('Tips').innerHTML = '�����˲���Ϊ��';
		the.JingShouRen.focus();
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
		Response.Redirect "gouru.asp"
		Response.End
	End If

RiQi=htmlencode(Request.form("RiQi"))
SheBei=htmlencode(Request.form("SheBei"))
PinPai=htmlencode(Request.form("PinPai"))
XingHao=htmlencode(Request.form("XingHao"))
XuLieHao=htmlencode(Request.form("XuLieHao"))
DanWei=htmlencode(Request.form("DanWei"))
ShuLiang=htmlencode(Request.form("ShuLiang"))
DanJia=htmlencode(Request.form("DanJia"))
JingShouRen=htmlencode(Request.form("JingShouRen"))
YongTu=htmlencode(Request.form("YongTu"))
OS=htmlencode(Request.form("OS"))
OSXuLieHao=htmlencode(Request.form("OSXuLieHao"))
Beizhu=htmlencode(Request.form("Beizhu"))

	If (Len(RiQi)>0 And Len(SheBei)>0 And Len(PinPai)>0 And Len(DanWei)>0 And Len(JingShouRen)>0 And _
		cCur(DanJia)>=0 And Cint(ShuLiang)>0 And Len(YongTu)>0) then
		If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
		If isNull(Conn) Then
			Set Conn=Server.CreateObject("ADODB.Connection")
			My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
			Conn.Open My_conn_STRING
		End If
'		If (Len(XuLieHao)>2 Or Len(OSXuLieHao)>2) Then
		If (Len(XuLieHao)>2) Then
			Sql="Select * from GouRu where XuLieHao='" & XuLieHao & "'"
			MyRs.open Sql,Conn,3,2
			If MyRs.recordcount>0 then
				Response.Write "<script>document.getElementById('Tips').innerHTML = '����豸�Ĺ�������Ѿ��Ǽǡ�';</SCRIPT>"
	'			MyRs.close
			Else
				Sql="INSERT INTO [GouRu]([RiQi],[SheBei],[PinPai],[XingHao],[XuLieHao],[DanWei],[ShuLiang],[DanJia],[JingShouRen],[YongTu],[OS],[OSXuLieHao],[BeiZhu]) VALUES('" & RiQi &"','" & SheBei &"','"& PinPai &"','"& XingHao &"','"& XuLieHao &"','"& DanWei &"','"& cint(ShuLiang)&"','"& cCur(DanJia) &"','"& JingShouRen &"','" & YongTu &"','"& OS &"','"& OSXuLieHao &"','"& Beizhu &"')"
				conn.execute(Sql)
				Response.Redirect "?Action=ShowGouRu"
		'		Response.End
			End If
		Else
			Sql="INSERT INTO [GouRu]([RiQi],[SheBei],[PinPai],[XingHao],[XuLieHao],[DanWei],[ShuLiang],[DanJia],[JingShouRen],[YongTu],[OS],[OSXuLieHao],[BeiZhu]) VALUES('" & RiQi &"','" & SheBei &"','"& PinPai &"','"& XingHao &"','"& XuLieHao &"','"& DanWei &"','"& cint(ShuLiang)&"','"& cCur(DanJia) &"','"& JingShouRen &"','" & YongTu &"','"& OS &"','"& OSXuLieHao &"','"& Beizhu &"')"
			conn.execute(Sql)
			Response.Redirect "?Action=ShowGouRu"
		End If
	End If

	MyRs.Close
End If
%>
<div id="Right_Content" style="align:left;float:left">
<%If Session("Admin")<>"" then%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>����豸�����¼</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Add_Area">
<form name="AddNewGouRu" id="AddNewGouRu" method="post" Action="?Action=AddRecord" onSubmit="return My_CheckFields(this);">
<span style="white-space: nowrap"><label for="Riqi">���ڣ�</label><input type="text" name="Riqi" id="Riqi" size="10" readonly="readonly" onclick="choose_date_czw('Riqi')"/></span>
<span style="white-space: nowrap"><label for="SheBei">�豸��</label><input type="text" name="SheBei" value="" id="SheBei" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="PinPai">Ʒ�ƣ�</label><input type="text" name="PinPai" value="��" id="PinPai" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="XingHao">�ͺţ�</label><input type="text" name="XingHao" value="��" id="XingHao" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="XuLieHao">���кţ�</label><input type="text" name="XuLieHao" value="��" id="XuLieHao" size="20" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="DanWei">��λ��</label>
<select name="DanWei" id="DanWei">
	<option value="��">��</option>
	<option value="��">��</option>
	<option value="̨">̨</option>
	<option value="ֻ">ֻ</option>
	<option value="֧">֧</option>
	<option value="��">��</option>
	<option value="ƿ">ƿ</option>
	<option value="��">��</option>
	<option value="Ƭ">Ƭ</option>
	<option value="��">��</option>
	<option value="��">��</option>
	<option value="��">��</option>
</select></span>
<span style="white-space: nowrap"><label for="ShuLiang">������</label><input type="text" name="ShuLiang" value="1" id="ShuLiang" size="5" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="DanJia">���ۣ�</label><input type="text" name="DanJia" value="" id="DanJia" size="5" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="JingShouRen">�����ˣ�</label><input type="text" name="JingShouRen" value=<%=Session("ShowName")%> id="JingShouRen" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="YongTu">��;��</label><input type="text" name="YongTu" value="" id="YongTu" size="20" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="OS">����ϵͳ��</label><input type="text" name="OS" value="��" id="OS" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="OSXuLieHao">ϵͳ���кţ�</label><input type="text" name="OSXuLieHao" value="��" id="OSXuLieHao" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="Beizhu">��ע��</label><input type="text" name="Beizhu" value="" id="Beizhu" size="30"/></span>
<input type="submit" value="���" onClick="return My_CheckFields(this);"/>
</form>
</div>
<hr>
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
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>�豸�������һ����</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Search_Area">
<form id="SearchSheBei" name="SearchSheBei" method="post" Action="?Action=Search" onSubmit="return My_CheckSearchDates(this);">
<span style="white-space: nowrap"><label for="S_Riqi">���ڣ���</label><input type="text" name="S_Riqi" id="S_Riqi" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/><label for="S_Riqi_2">��</label><input type="text" name="S_Riqi_2" id="S_Riqi_2" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/></span>
<span style="white-space: nowrap"><label for="S_SheBei">�豸��</label><input type="text" name="S_SheBei" value="" id="S_SheBei" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="S_PinPai">Ʒ�ƣ�</label><input type="text" name="S_PinPai" value="" id="S_PinPai" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="S_XingHao">�ͺţ�</label><input type="text" name="S_XingHao" value="" id="S_XingHao" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="S_XuLieHao">���кţ�</label><input type="text" name="S_XuLieHao" value="" id="S_XuLieHao" size="20" onblur="return My_CheckField(this);"></span>
<!--  -->
<span style="white-space: nowrap"><label for="S_JingShouRen">�����ˣ�</label><input type="text" name="S_JingShouRen" value="" id="S_JingShouRen" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_YongTu">��;��</label><input type="text" name="S_YongTu" value="" id="S_YongTu" size="20" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="S_OS">����ϵͳ��</label><input type="text" name="S_OS" value="" id="S_OS" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="S_OSXuLieHao">ϵͳ���кţ�</label><input type="text" name="S_OSXuLieHao" value="" id="S_OSXuLieHao" size="10" onblur="return My_CheckField(this);"></span>
<input type="submit" value="����" name="S_Submit" id="S_Submit"/>
</form>
</div>
<%
'������¼

S_Riqi=htmlencode(Request.form("S_Riqi"))
S_Riqi_2=htmlencode(Request.form("S_Riqi_2"))
S_SheBei=htmlencode(Request.form("S_SheBei"))
S_ShiYongRen=htmlencode(Request.form("S_JingShouRen"))
S_YongTu=htmlencode(Request.form("S_YongTu"))
S_PinPai=htmlencode(Request.form("S_PinPai"))
S_XingHao=htmlencode(Request.form("S_XingHao"))
S_XuLieHao=htmlencode(Request.form("S_XuLieHao"))
S_OS=htmlencode(Request.form("S_OS"))
S_OSXuLieHao=htmlencode(Request.form("S_OSXuLieHao"))

SQL="select * from GouRu where 1=1"
If Len(S_Riqi)<>0 AND Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi &"' and '" & S_Riqi_2 &"'"
If Len(S_Riqi)<>0 XOR Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi & S_Riqi_2 &"' and '" & S_Riqi & S_Riqi_2 &"'"
If Len(S_SheBei)<>0 Then SQL = SQL & " and SheBei Like '%" & S_SheBei &"%'"
If Len(S_JingShouRen)<>0 Then SQL = SQL & " and JingShouRen = '" & S_JingShouRen &"'"
If Len(S_YongTu)<>0 Then SQL = SQL & " and YongTu Like '%" & S_YongTu &"%'"
If Len(S_XuLieHao)<>0 Then SQL = SQL & " and XuLieHao = '" & S_XuLieHao &"'"
If Len(S_OS)<>0 Then SQL = SQL & " and OS Like '%" & S_OS &"%'"
If Len(S_PinPai)<>0 Then SQL = SQL & " and PinPai Like '%" & S_PinPai &"%'"
If Len(S_XingHao)<>0 Then SQL = SQL & " and XingHao Like '%" & S_XingHao &"%'"
If Len(S_OSXuLieHao)<>0 Then SQL = SQL & " and OSXuLieHao = '" & S_OSXuLieHao &"'"
SQL = SQL & " order by RiQi desc, JingShouRen desc, SheBei desc"
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
			response.Write "<th width='70px'><b>" & "����" & "</b></th>"
		Case "SHEBEI":
			response.Write "<th><b>" & "�豸" & "</b></th>"
		Case "PINPAI":
			response.Write "<th><b>" & "Ʒ��" & "</b></th>"
		Case "XINGHAO":
			response.Write "<th><b>" & "�ͺ�" & "</b></th>"
		Case "XULIEHAO":
			response.Write "<th><b>" & "���к�" & "</b></th>"
		Case "DANWEI":
			response.Write "<th width='30px'><b>" & "��λ" & "</b></th>"
		Case "SHULIANG":
			response.Write "<th width='30px'><b>" & "����" & "</b></th>"
		Case "DANJIA":
			response.Write "<th><b>" & "����" & "</b></th>"
			response.Write "<th><b>" & "�ܼ�" & "</b></th>"
		Case "JINGSHOUREN":
			response.Write "<th width='50px'><b>" & "������" & "</b></th>"
		Case "YONGTU":
			response.Write "<th><b>" & "��;" & "</b></th>"
		Case "OS":
			response.Write "<th width='60px'><b>" & "����ϵͳ" & "</b></th>"
		Case "OSXULIEHAO":
			response.Write "<th width='80px'><b>" & "ϵͳ���к�" & "</b></th>"
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
'	If MyRs.EOF Then Exit For
	if 1 = i_s mod 2 then
		response.write("<tr id='Data' class='odd'>")
	else
		response.write("<tr id='Data'>")
	end if
	for i_c = 1 to howmanyfields '����ʾId�ֶ�
		ThisRecord = MyRs(i_c).Value
		If IsNull(ThisRecord) OR Len(Trim(ThisRecord))<1 Then
			ThisRecord = "&nbsp;"
		end if
		If Ucase(MyRs(i_c).Name)="BEIZHU" Then
			Response.write("<td class='BeiZhu'>" & ThisRecord & "</td>")
		ElseIf Ucase(MyRs(i_c).Name)="SHULIANG" Then
			Response.write("<td>" & ThisRecord & "</td>")
			iShuLiang=ThisRecord
		ElseIf Ucase(MyRs(i_c).Name)="DANJIA" Then
			Response.write("<td>" & FormatCurrency(ThisRecord) & "</td>")
			iDanJia=ThisRecord
			Response.write("<td>" & FormatCurrency(1*iShuLiang*iDanJia) & "</td>")
		ElseIf Ucase(MyRs(i_c).Name)="XULIEHAO" And Len(Trim(MyRs(i_c).Value)) > 2 Then
			If Session("Admin")="" then
				Response.write("<td> ******** </td>")
			Else
				Response.write("<td>" & ThisRecord & "</td>")
			End If
		ElseIf Ucase(MyRs(i_c).Name)="OSXULIEHAO" And Len(Trim(MyRs(i_c).Value)) > 2 Then
			If Session("Admin")="" then
				Response.write("<td> ******** </td>")
			Else
				Response.write("<td>" & ThisRecord & "</td>")
			End If
		Else
			Response.write("<td>" & ThisRecord & "</td>")
		End If
	next
	response.write("</tr>")
	MyRs.movenext
	If MyRs.EOF Then Exit For
Next
%>
</tbody>
</table>
<div id="navbar" align="left">
<br clear="left">
<%
response.write "���ҳ�룺"
PageCount=-Int(-ResultCount/(PageSize+1))
'Function Ceil(x) Ceil = -Int(-x) End Function
'response.write ResultCount &"-" & PageSize & "-" & PageCount

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
'StartPage=1
'EndPage = PageCount

response.write "<a href=""gouru.asp?page=1"">��һҳ</a> "
for i=StartPage to EndPage
	response.write "<a href=""gouru.asp?page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""gouru.asp?page=" & PageCount & """>���ҳ</a> "
Else
	Response.Write "<h1>û���ҵ��κν��������Ĺؼ��ʣ�������������</h1>"
End If
'End Select
'MyRs.close
'Set MyRs= Nothing
Conn.Close
set Conn=nothing
%>
<p align="left"><a href="download.asp?Action=gouru">�����豸��������</a></p>
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