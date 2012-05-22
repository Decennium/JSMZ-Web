<!--#include file="top.asp"-->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>领取录音机管理系统</title>
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
function My_IsInt(val) {
	if ((parseInt(val) == val) && (parseInt(val) >= 0)) {
		return true;
	}
		return false;
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
	if (the.JiaoShi.value=="") {
		document.getElementById('Tips').innerHTML = '领取录音机的教师不能为空';
		the.JiaoShi.focus();
		return false;
	}
	if (the.XueKe.value=="") {
		document.getElementById('Tips').innerHTML = '教师所在学科不能为空';
		the.XueKe.focus();
		return false;
	}
	if (the.FaFangRen.value=="") {
		document.getElementById('Tips').innerHTML = '录音机发放人不能为空';
		the.FaFangRen.focus();
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
			If DateDiff("m",MyRs(1).Value,Now())<35 Then
				Response.Write "<script language='javascript'>document.getElementById(""Tips"").innerHTML = '" & MyRs(2).Value & "老师上一台录音机领取日期为" & FormatDateTime(MyRs(1).Value,1) & "，还差" & 36-DateDiff("m",MyRs(1).Value,Now()) & "个月才满三年。';</script>"
			Else
				Sql="INSERT INTO [LuYinJi]([RiQi],[JiaoShi],[XueKe],[FaFangRen],[BeiZhu]) " & _
					"VALUES ('"& RiQi &"','"& JiaoShi &"','"& XueKe &"','"& FaFangRen &"','" & Beizhu &"')"
				conn.execute(Sql)
				'Response.Redirect "?Action=ShowJieci&page=" & Currentpage
			End If
		Else
			Sql="INSERT INTO [LuYinJi]([RiQi],[JiaoShi],[XueKe],[FaFangRen],[BeiZhu]) " & _
				"VALUES ('"& RiQi &"','"& JiaoShi &"','"& XueKe &"','"& FaFangRen &"','" & Beizhu &"')"
			conn.execute(Sql)
			'Response.Redirect "?Action=ShowJieci&page=" & Currentpage
		End If
	End If

	MyRs.Close
End If
%>
<div id="Right_Content" style="align:left;float:left">
<%If Session("Admin")<>"" then%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>添加领取录音机记录</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Add_Area">
<form name="AddNewLingQu" id="AddNewLingQu" method="post" Action="?Action=AddRecord" onSubmit="return My_CheckFields(this);">
<span style="white-space: nowrap"><label for="Riqi">日期：</label><input type="text" name="Riqi" id="Riqi" size="10" readonly="readonly" onclick="choose_date_czw('Riqi')"/></span>
<span style="white-space: nowrap"><label for="JiaoShi">教师：</label><input type="text" name="JiaoShi" value="" id="JiaoShi" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="XueKe">学科：</label>
<select name="XueKe" id="XueKe">
	<option value="英语">英语</option>
	<option value="音乐">音乐</option>
	<option value="舞蹈">舞蹈</option>
	<option value="语文">语文</option>
	<option value="其他学科">其他学科</option>
</select></span>
<span style="white-space: nowrap"><label for="FaFangRen">发放人：</label><input type="text" name="FaFangRen" value='<%=Session("ShowName")%>' id="FaFangRen" size="10" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="Beizhu">备注：</label><input type="text" name="Beizhu" value="" id="Beizhu" size="30"/></span>
<input type="submit" value="添加"/>
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
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>领取录音机情况一览表</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Search_Area">
<form id="SearchSheBei" name="SearchSheBei" method="post" Action="?Action=Search" onSubmit="return My_CheckSearchDates(this);">
<span style="white-space: nowrap"><label for="S_Riqi">日期：从</label><input type="text" name="S_Riqi" id="S_Riqi" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/><label for="S_Riqi_2">到</label><input type="text" name="S_Riqi_2" id="S_Riqi_2" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/></span>
<span style="white-space: nowrap"><label for="S_JiaoShi">教师：</label><input type="text" name="S_JiaoShi" value="" id="S_JiaoShi" size="40" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_XueKe">学科：</label>
<select name="S_XueKe" id="S_XueKe">
	<option value="%">所有学科</option>
	<option value="英语">英语</option>
	<option value="音乐">音乐</option>
	<option value="舞蹈">舞蹈</option>
	<option value="语文">语文</option>
	<option value="其他学科">其他学科</option>
</select></span>
<span style="white-space: nowrap"><label for="S_DaoQi">是否到期：</label>
<select name="S_DaoQi" id="S_DaoQi">
	<option value="-1">不论</option>
	<option value="0">到期</option>
	<option value="1">未到期</option>
</select></span>
<span style="white-space: nowrap"><label for="S_FaFangRen">发放人：</label><input type="text" name="S_FaFangRen" value="" id="S_FaFangRen" size="10" onblur="return My_CheckField(this);"/></span>
<input type="submit" value="搜索" name="S_Submit" id="S_Submit"/>
</form>
</div>
<%
'搜索记录

S_Riqi=htmlencode(Request.form("S_Riqi"))
S_Riqi_2=htmlencode(Request.form("S_Riqi_2"))
S_JiaoShi=htmlencode(Request.form("S_JiaoShi"))
S_XueKe=htmlencode(Request.form("S_XueKe"))
S_FaFangRen=htmlencode(Request.form("S_FaFangRen"))
S_DaoQi=htmlencode(Request.form("S_DaoQi"))

SQL="select * from LuYinJi where 1=1"
If Len(S_Riqi)<>0 AND Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi &"' and '" & S_Riqi_2 &"'"
If Len(S_Riqi)<>0 XOR Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi & S_Riqi_2 &"' and '" & S_Riqi & S_Riqi_2 &"'"
If Len(S_JiaoShi)<>0 Then SQL = SQL & " and JiaoShi Like '" & S_JiaoShi &"'"
If Len(S_XueKe)<>0 Then SQL = SQL & " and XueKe Like '" & S_XueKe &"'"
If Len(S_FaFangRen)<>0 Then SQL = SQL & " and FaFangRen Like '" & S_FaFangRen &"'"
Select Case S_DaoQi
	Case "0"
		SQL = SQL & " and DATEDIFF(month,GETDATE(),DATEADD(month,36,RiQi)) <=1"
	Case "1"
		SQL = SQL & " and DATEDIFF(month,GETDATE(),DATEADD(month,36,RiQi)) > 1"
End Select
SQL = SQL & " order by Riqi asc, JiaoShi asc"
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
			response.Write "<th><b>" & "日期" & "</b></th>"
		Case "JIAOSHI":
			response.Write "<th><b>" & "教师" & "</b></th>"
		Case "XUEKE":
			response.Write "<th><b>" & "学科" & "</b></th>"
		Case "FAFANGREN":
			response.Write "<th><b>" & "发放人" & "</b></th>"
		Case "BEIZHU":
			response.Write "<th><b>" & "距离更换" & "</b></th>"
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
		ThisRecord = MyRs(i_c).value
		If IsNull(ThisRecord) Then
			ThisRecord = "&nbsp;"
		end if
		If Ucase(MyRs(i_c).Name)="BEIZHU" Then
			ShengYu = datediff("m",dateadd("m",36,MyRs(1).value),Now())
			If ShengYu < 0 Then
				ShengYu = "还有" & ABS(ShengYu) & "个月"
			Else
				ShengYu = "已经报废"
			End If
			Response.write("<td>" & ShengYu & "</td>")
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
PageCount=Int(ResultCount/(PageSize+1))+1
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

response.write "<a href=""luyinji.asp?page=1"">第一页</a> "
for i=StartPage to EndPage
	response.write "<a href=""luyinji.asp?page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""luyinji.asp?page=" & PageCount & """>最后页</a> "
Else
	Response.Write "<h1>没有找到任何结果，请更改关键词，并重新搜索。</h1>"
End If
'End Select
Conn.Close
set Conn=nothing
%>
<p align="left"><a href="download.asp?Action=luyinji">下载领取录音机数据</a></p>
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