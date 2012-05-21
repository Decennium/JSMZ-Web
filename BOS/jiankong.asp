<%
'过滤代码
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

'数据库连接
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
<title>监控设备使用管理系统</title>
<link href="../css/css.css" rel="stylesheet">
<link rel="shortcut icon" type="image/x-icon" href="favicon.ico">
<script language="javascript" src="DateTime.js"></script>
<script language="javascript">
/** 判断输入框中输入的日期格式为yyyy-mm-dd和正确的日期 */ 
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
function My_IsClassNum(val) {
	return /^([HZ][1-5][0-1][0-6](-[HZ][1-5][0-1][0-6]){0,1}\,{0,1})*$/g.test(val);
}
function My_CheckFields(the){
	if (the.Riqi.value=="") {
		document.getElementById('Tips').innerHTML = '日期不能为空';
		the.Riqi.focus();
		return false;
	}
	if (!My_IsDate(the.Riqi.value)){
		document.getElementById('Tips').innerHTML = '请保证输入的日期格式为yyyy-mm-dd或正确的日期!';
		the.Riqi.focus();
		return false;
	}
	if (!My_IsClassNum(the.ClassNum.value)){
		document.getElementById('Tips').innerHTML = '请保证输入的教室编号格式如H101-H412形式!';
		the.ClassNum.focus();
		return false;
	}
	if (the.YongTu.value=="") {
		document.getElementById('Tips').innerHTML = '用途不能为空';
		the.YongTu.focus();
		return false;
	}
	if (the.ShiYongRen.value=="") {
		document.getElementById('Tips').innerHTML = '使用人不能为空';
		the.ShiYongRen.focus();
		return false;
	}
	document.getElementById('Tips').innerHTML = '';
	return true;
}
function My_CheckField(the){
	if (the.value==""){
		document.getElementById('Tips').innerHTML = '字段不能为空';
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
'添加记录
	If Session("Admin")="" then
	'判断是否登陆
		Response.Redirect "jiankong.asp"
		Response.End
	End If

RiQi=htmlencode(Request.form("RiQi"))
ShiJianDuan=htmlencode(Request.form("ShiJianDuan"))
ClassNum=htmlencode(Request.form("ClassNum"))
YongTu=htmlencode(Request.form("YongTu"))
ShenQingRen=htmlencode(Request.form("ShenQingRen"))
CaoZuoYuan=htmlencode(Request.form("CaoZuoYuan"))
Beizhu=htmlencode(Request.form("Beizhu"))

	If (Len(RiQi)>0 And Len(ShiJianDuan)>0 And Len(ClassNum)>0 And Len(YongTu)>0 And Len(ShenQingRen)>0 And Len(CaoZuoYuan)>0) Then
		If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
		If isNull(Conn) Then
			Set Conn=Server.CreateObject("ADODB.Connection")
			My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
			Conn.Open My_conn_STRING
		End If
		Sql="INSERT INTO [JianKong]([RiQi],[ShiJianDuan],[ClassNum],[YongTu],[ShenQingRen],[CaoZuoYuan],[BeiZhu]) VALUES ('"& RiQi &"','"& ShiJianDuan &"','"& ClassNum &"','"& YongTu &"','" & ShenQingRen &"','" & CaoZuoYuan &"','" & Beizhu &"')"
		conn.execute(Sql)
		Response.Redirect "?Action=ShowSheBei"
'			Response.End
	End If

	MyRs.Close
End If
If Action = "AddCheck" Then
'添加检查
	If Session("Admin")="" then
	'判断是否登陆
		Response.Redirect "jiankong.asp"
		Response.End
	End If
ID=htmlencode(Request.form("id"))
XiaoGuo=htmlencode(Request.form("XiaoGuo"))
	If (Len(ID)>0 And Len(XiaoGuo)>0) then
		If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
		If isNull(Conn) Then
			Set Conn=Server.CreateObject("ADODB.Connection")
			My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
			Conn.Open My_conn_STRING
		End If
			Sql="UPDATE [JianKong] SET [XiaoGuo] = '"& XiaoGuo &"' WHERE [id] ='"& ID &"'"
			conn.execute(Sql)
			Response.Redirect "?Action=ShowJieci"
	End If
End If
%>
<div id="Right_Content" style="align:left;float:left">
<%If Session("Admin")<>"" then%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>添加监控设备使用记录</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Add_Area">
<form name="AddNewShiYong" id="AddNewShiYong" method="post" Action="?Action=AddRecord" onSubmit="return My_CheckFields(this);">
<span style="white-space: nowrap"><label for="Riqi">日期：</label><input type="text" name="Riqi" id="Riqi" size="10" readonly="readonly" onclick="choose_date_czw('Riqi')"/></span>
<span style="white-space: nowrap"><label for="ShiJianDuan">时间段：</label><input type="text" name="ShiJianDuan" value="" id="ShiJianDuan" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="ClassNum">教室编号：</label><input type="text" name="ClassNum" id="ClassNum" value="H101-H412" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="YongTu">用途：</label><input type="text" name="YongTu" value="" id="YongTu" size="40" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="ShenQingRen">申请人：</label><input type="text" name="ShenQingRen" value="" id="ShenQingRen" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="CaoZuoYuan">操作员：</label><input type="text" name="CaoZuoYuan" value=<%=Session("ShowName")%> id="CaoZuoYuan" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="XiaoGuo">效果：</label><input type="text" name="XiaoGuo" value="" id="XiaoGuo" size="10"/></span>
<span style="white-space: nowrap"><label for="Beizhu">备注：</label><input type="text" name="Beizhu" value="" id="Beizhu" size="30"/></span>
<input type="submit" value="添加" onClick="return My_CheckFields(this);"/>
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
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>监控设备使用情况一览表</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Search_Area">
<form id="SearchSheBei" name="SearchSheBei" method="post" Action="?Action=Search" onSubmit="return My_CheckSearchDates(this);">
<span style="white-space: nowrap"><label for="S_Riqi">日期：从</label><input type="text" name="S_Riqi" id="S_Riqi" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/><label for="S_Riqi_2">到</label><input type="text" name="S_Riqi_2" id="S_Riqi_2" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/></span>
<span style="white-space: nowrap"><label for="S_ClassNum">教室编号：</label><input type="text" name="S_ClassNum" id="S_ClassNum" value="" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="S_YongTu">用途：</label><input type="text" name="S_YongTu" id="S_YongTu" value="" size="40" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_ShenQingRen">申请人：</label><input type="text" name="S_ShenQingRen" id="S_ShenQingRen" value="" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_CaoZuoYuan">操作员：</label><input type="text" name="S_CaoZuoYuan" id="S_CaoZuoYuan" value="" size="5" onblur="return My_CheckField(this);"/></span>
<input type="submit" value="搜索" name="S_Submit" id="S_Submit"/>
</form>
</div>
<%
'搜索记录

S_Riqi=htmlencode(Request.form("S_Riqi"))
S_Riqi_2=htmlencode(Request.form("S_Riqi_2"))
S_ClassNum=htmlencode(Request.form("S_ClassNum"))
S_ShenQingRen=htmlencode(Request.form("S_ShenQingRen"))
S_CaoZuoYuan=htmlencode(Request.form("S_CaoZuoYuan"))
S_YongTu=htmlencode(Request.form("S_YongTu"))

SQL="select * from JianKong where 1=1"
If Len(S_Riqi)<>0 AND Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi &"' and '" & S_Riqi_2 &"'"
If Len(S_Riqi)<>0 XOR Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi & S_Riqi_2 &"' and '" & S_Riqi & S_Riqi_2 &"'"
If Len(S_ClassNum)<>0 Then SQL = SQL & " and ClassNum Like '%" & S_ClassNum &"%'"
'ClassNum的判断方式需要增加数值判断
If Len(S_ShenQingRen)<>0 Then SQL = SQL & " and ShenQingRen Like '" & S_ShenQingRen &"'"
If Len(S_CaoZuoYuan)<>0 Then SQL = SQL & " and CaoZuoYuan Like '" & S_CaoZuoYuan &"'"
If Len(S_YongTu)<>0 Then SQL = SQL & " and YongTu Like '%" & S_YongTu &"%'"
SQL = SQL & " order by Riqi desc, ClassNum desc, ShenQingRen desc"

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
			response.Write "<th><b>" & "日期" & "</b></th>"
		Case "SHIJIANDUAN":
			response.Write "<th><b>" & "时间段" & "</b></th>"
		Case "CLASSNUM":
			response.Write "<th><b>" & "教室编号" & "</b></th>"
		Case "YONGTU":
			response.Write "<th class='NeiRong'><b>" & "用途" & "</b></th>"
		Case "SHENQINGREN":
			response.Write "<th><b>" & "申请人" & "</b></th>"
		Case "CAOZUOYUAN":
			response.Write "<th><b>" & "操作员" & "</b></th>"
		Case "XIAOGUO":
			response.Write "<th><b>" & "效果" & "</b></th>"
		Case "BEIZHU":
			response.Write "<th class='BeiZhu'><b>" & "备注" & "</b></th>"
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
	for i_c = 1 to howmanyfields '不显示Id字段
		ThisRecord = MyRs(i_c)
		If IsNull(ThisRecord) Then
			ThisRecord = ""
		end if
		If Ucase(MyRs(i_c).Name)="XIAOGUO" Then
			If ThisRecord = "" Then
				Response.write("<form name='AddCheck' id='AddCheck' method='post' Action='?Action=AddCheck'><td><input type='hidden' name='id' value='" & MyRs(0).Value & "'/><input type='text' name='XiaoGuo' value='使用状况良好' id='XiaoGuo' size='20'/><input type='submit' value='检查'/></td>")
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
response.write "结果页码："
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

response.write "<a href=""jiankong.asp?page=1"">第一页</a> "
for i=StartPage to EndPage
	response.write "<a href=""jiankong.asp?page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""jiankong.asp?page=" & PageCount & """>最后页</a> "
Else
	Response.Write "<h1>没有找到任何结果，请更改关键词，并重新搜索。</h1>"
End If
'End Select
'MyRs.close
'Set MyRs= Nothing
Conn.Close
set Conn=nothing
%>
<p align="left"><a href="download.asp?Action=jiankong">下载监控设备使用数据</a></p>
</div>
</div>
<br clear=all>
<br clear=all>
<!--#include file="../include/bottom.asp"-->
<script language="javascript">
/*将获取的值存到变量里*/
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