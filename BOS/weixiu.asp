<!--#include file="top.asp"-->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>�豸ά�޹���ϵͳ</title>
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
If Action = "AddRecord" Then
'��Ӽ�¼
	If Session("Admin")="" then
	'�ж��Ƿ��½
		Response.Redirect "weixiu.asp"
		Response.End
	End If

RiQi=htmlencode(Request.form("RiQi"))
SheBei=htmlencode(Request.form("SheBei"))
GuZhang=htmlencode(Request.form("GuZhang"))
ShenBaoRen=htmlencode(Request.form("ShenBaoRen"))
FenXi=htmlencode(Request.form("FenXi"))
PaiChu=htmlencode(Request.form("PaiChu"))
ShiGongZhe=htmlencode(Request.form("ShiGongZhe"))
Beizhu=htmlencode(Request.form("Beizhu"))

	If (Len(SheBei)>0 And Len(GuZhang)>0 And Len(RiQi)>0 And Len(ShenBaoRen)>0 And Len(ShiGongZhe)>0 And Len(PaiChu)>0) Then
		If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
		If isNull(Conn) Then
			Set Conn=Server.CreateObject("ADODB.Connection")
			My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
			Conn.Open My_conn_STRING
		End If
		Sql="INSERT INTO [WeiXiu]([RiQi],[SheBei],[GuZhang],[ShenBaoRen],[FenXi],[PaiChu],[ShiGongZhe],[BeiZhu]) VALUES ('"& RiQi &"','"& SheBei &"','"& GuZhang &"','"& ShenBaoRen &"','"& FenXi &"','"& PaiChu &"','"& ShiGongZhe &"','"& Beizhu &"')"
		conn.execute(Sql)
		Response.Redirect "?Action=ShowSheBei"
'			Response.End
	End If

	MyRs.Close
End If
%>
<div id="Right_Content" style="align:left;float:left">
<%If Session("Admin")<>"" then%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>����豸ά�޼�¼</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Add_Area">
<form name="AddNewShiYong" id="AddNewShiYong" method="post" Action="?Action=AddRecord" onSubmit="return My_CheckFields(this);">
<span style="white-space: nowrap"><label for="Riqi">���ڣ�</label><input type="text" name="Riqi" id="Riqi" size="10" readonly="readonly" onclick="choose_date_czw('Riqi')"/></span>
<span style="white-space: nowrap"><label for="SheBei">�豸��</label><input type="text" name="SheBei" value="" id="SheBei" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="GuZhang">����������</label><input type="text" name="GuZhang" value="" id="GuZhang" size="30" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="ShenBaoRen">�걨�ˣ�</label><input type="text" name="ShenBaoRen" value="" id="ShenBaoRen" size="10" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="FenXi">���Ϸ�����</label><input type="text" name="FenXi" value="" id="FenXi" size="20" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="PaiChu">�ų������</label><input type="text" name="PaiChu" value="" id="PaiChu" size="20" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="ShiGongZhe">ʩ���ˣ�</label><input type="text" name="ShiGongZhe" value=<%=Session("ShowName")%> id="ShiYongRen" size="10" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="Beizhu">��ע��</label><input type="text" name="Beizhu" value="" id="Beizhu" size="20"/></span>
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
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>�豸ά�����һ����</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Search_Area">
<form id="SearchSheBei" name="SearchSheBei" method="post" Action="?Action=Search" onSubmit="return My_CheckSearchDates(this);">
<span style="white-space: nowrap"><label for="S_Riqi">���ڣ���</label><input type="text" name="S_Riqi" id="S_Riqi" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/><label for="S_Riqi_2">��</label><input type="text" name="S_Riqi_2" id="S_Riqi_2" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/></span>
<span style="white-space: nowrap"><label for="S_Shebei">�豸��</label><input type="text" name="S_Shebei" value="" id="S_Shebei" size="20" title="�����벿�ֹؼ���" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_GuZhang">����������</label><input type="text" name="S_GuZhang" value="" id="S_GuZhang" size="30" title="�����벿�ֹؼ���" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_ShenBaoRen">�걨�ˣ�</label><input type="text" name="S_ShenBaoRen" value="" id="S_ShenBaoRen" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_FenXi">���Ϸ�����</label><input type="text" name="S_FenXi" value="" id="S_FenXi" size="30" title="�����벿�ֹؼ���" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_PaiChu">�ų������</label><input type="text" name="S_PaiChu" value="" id="S_PaiChu" size="20" title="�����벿�ֹؼ���" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_ShiGongZhe">ʩ���ˣ�</label><input type="text" name="S_ShiGongZhe" value="" id="S_ShiGongZhe" size="5" onblur="return My_CheckField(this);"/></span>
<input type="submit" value="����" name="S_Submit" id="S_Submit"/>
</form>
</div>
<%
'������¼

S_Riqi=htmlencode(Request.form("S_Riqi"))
S_Riqi_2=htmlencode(Request.form("S_Riqi_2"))
S_SheBei=htmlencode(Request.form("S_SheBei"))
S_GuZhang=htmlencode(Request.form("S_GuZhang"))
S_PaiChu=htmlencode(Request.form("S_PaiChu"))
S_FenXi=htmlencode(Request.form("S_FenXi"))
S_ShiGongZhe=htmlencode(Request.form("S_ShiGongZhe"))
S_ShenBaoRen=htmlencode(Request.form("S_ShenBaoRen"))

SQL="select * from WeiXiu where 1=1"
If Len(S_Riqi)<>0 AND Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi &"' and '" & S_Riqi_2 &"'"
If Len(S_Riqi)<>0 XOR Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi & S_Riqi_2 &"' and '" & S_Riqi & S_Riqi_2 &"'"
If Len(S_SheBei)<>0 Then SQL = SQL & " and SheBei Like '%" & S_SheBei &"%'"
If Len(S_GuZhang)<>0 Then SQL = SQL & " and GuZhang Like '%" & S_GuZhang &"%'"
If Len(S_FenXi)<>0 Then SQL = SQL & " and FenXi Like '%" & S_FenXi &"%'"
If Len(S_PaiChu)<>0 Then SQL = SQL & " and PaiChu Like '%" & S_PaiChu &"%'"
If Len(S_ShenBaoRen)<>0 Then SQL = SQL & " and ShenBaoRen Like '%" & S_ShenBaoRen &"%'"
If Len(S_ShiGongZhe)<>0 Then SQL = SQL & " and ShiGongZhe = '" & S_ShiGongZhe &"'"
SQL = SQL & " order by Riqi desc, ShiGongZhe desc, SheBei desc"

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
		Case "SHEBEI":
			response.Write "<th><b>" & "�豸" & "</b></th>"
		Case "GUZHANG":
			response.Write "<th><b>" & "���ϱ���" & "</b></th>"
		Case "SHENBAOREN":
			response.Write "<th width='60px'><b>" & "�걨��" & "</b></th>"
		Case "FENXI":
			response.Write "<th class='NeiRong'><b>" & "���Ϸ���" & "</b></th>"
		Case "PAICHU":
			response.Write "<th class='NeiRong'><b>" & "�ų����" & "</b></th>"
		Case "SHIGONGZHE":
			response.Write "<th width='60px'><b>" & "ʩ����" & "</b></th>"
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

response.write "<a href=""weixiu.asp?page=1"">��һҳ</a> "
for i=StartPage to EndPage
	response.write "<a href=""weixiu.asp?page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""weixiu.asp?page=" & PageCount & """>���ҳ</a> "
Else
	Response.Write "<h1>û���ҵ��κν��������Ĺؼ��ʣ�������������</h1>"
End If
'End Select
'MyRs.close
'Set MyRs= Nothing
Conn.Close
set Conn=nothing
%>
<p align="left"><a href="download.asp?Action=weixiu">�����豸ά������</a></p>
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