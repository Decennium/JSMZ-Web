<!--#include file="../include/top.asp"-->
<html>
<head>
<meta name="viewport" content="width=device-width,minimum-scale=1.0, maximum-scale=2.0"/>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>图书室开放登记系统</title>
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
		document.getElementById('Tips').innerHTML = '教师不能为空';
		the.JiaoShi.focus();
		return false;
	}
	if (the.QingKuang.value=="") {
		document.getElementById('Tips').innerHTML = '开放情况不能为空';
		the.QingKuang.focus();
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
<!--#include file="../include/second_banner.asp"-->
<%
Action=Request.Querystring("Action")
If Action = "AddRecord" Then
'添加记录
	If Session("Admin")="" then
	'判断是否登陆
		Response.Redirect "tushushi.asp"
		Response.End
	End If

RiQi=htmlencode(Request.form("RiQi"))
JiaoShi=htmlencode(Request.form("JiaoShi"))
ROOM1=htmlencode(Request.form("ROOM1"))
ROOM2=htmlencode(Request.form("ROOM2"))
ROOM3=htmlencode(Request.form("ROOM3"))
STUDENT1=htmlencode(Request.form("STUDENT1"))
STUDENT2=htmlencode(Request.form("STUDENT2"))
STUDENT3=htmlencode(Request.form("STUDENT3"))
QingKuang=htmlencode(Request.form("QingKuang"))
Beizhu=htmlencode(Request.form("Beizhu"))

	If (Len(RiQi)>0 And Len(JiaoShi)>0 And Len(ROOM1)>0 And Len(ROOM2)>0 And Len(ROOM3)>0 And Len(STUDENT1)>0 And Len(STUDENT2)>0 And Len(STUDENT3)>0 And Len(QingKuang)>0) Then
		If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
		If isNull(Conn) Then
		Set Conn=Server.CreateObject("ADODB.Connection")
	'		If Conn.State=adStateOpen Then Conn.Close
			My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
			Conn.Open My_conn_STRING
		End If
		Sql="Select * from tushushi where RiQi='" & RiQi & "'"
		MyRs.open Sql,Conn,3,2
		If MyRs.recordcount>0 then
			Response.Write "<script>document.getElementById('Tips').innerHTML = '今天的阅览室开放情况已经登记。';</SCRIPT>"
		Else
			Sql="INSERT INTO [TUSHUSHI]([RiQi],[JiaoShi],[ROOM1],[ROOM2],[ROOM3],[STUDENT1],[STUDENT2],[STUDENT3],[QingKuang],[BeiZhu]) VALUES ('"& RiQi &"','"& JiaoShi &"','"& QingKuang &"','" & Beizhu &"')"
			conn.execute(Sql)
			Response.Redirect "?Action=ShoweReading"
	'		Response.End
		End If
	End If
	MyRs.Close
	'Conn.Close
End If
%>
<div id="Content" style="align:left;float:left">
<%If Session("Admin")<>"" then%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>添加图书室阅读记录</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Add_Area">
<form name="AddNewShiYong" id="AddNewShiYong" method="post" Action="?Action=AddRecord" onSubmit="return My_CheckFields(this);">
<span style="white-space: nowrap"><label for="Riqi">日期：</label><input type="text" name="Riqi" id="Riqi" size="10" readonly="readonly" onclick="choose_date_czw('Riqi')"/></span>
<span style="white-space: nowrap"><label for="JiaoShi">教师：</label><input type="text" name="JiaoShi" value=<%=Session("ShowName")%> id="JiaoShi" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="ROOM1">一号阅览室人数：</label><input type="text" name="ROOM1" value="" id="ROOM1" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="ROOM2">二号阅览室人数：</label><input type="text" name="ROOM2" value="" id="ROOM2" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="ROOM3">三号阅览室人数：</label><input type="text" name="ROOM3" value="" id="ROOM3" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="STUDENT1">第一组志愿者：</label><input type="text" name="STUDENT1" value="" id="STUDENT1" size="20" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="STUDENT2">第二组志愿者：</label><input type="text" name="STUDENT2" value="" id="STUDENT2" size="20" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="STUDENT3">第三组志愿者：</label><input type="text" name="STUDENT3" value="" id="STUDENT3" size="20" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="QingKuang">开放情况：</label><input type="text" name="QingKuang" value="良好" id="QingKuang" size="20" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="Beizhu">备注：</label><input type="text" name="Beizhu" value="" id="Beizhu" size="5"/></span>
<input type="submit" value="添加" onClick="return My_CheckFields(this);"/>
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
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>阅览室开放情况一览表</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Search_Area">
<form id="SearchSheBei" name="SearchSheBei" method="post" Action="?Action=Search" onSubmit="return My_CheckSearchDates(this);">
<span style="white-space: nowrap"><label for="S_Riqi">日期：从</label><input type="text" name="S_Riqi" id="S_Riqi" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/><label for="S_Riqi_2">到</label><input type="text" name="S_Riqi_2" id="S_Riqi_2" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/></span>
<span style="white-space: nowrap"><label for="S_JiaoShi">教师：</label><input type="text" name="S_JiaoShi" value="" id="S_JiaoShi" size="10" title="请输入教师姓名" onblur="return My_CheckField(this);"/></span>
<input type="submit" value="搜索" name="S_Submit" id="S_Submit"/>
</form>
</div>
<%
'搜索记录

S_Riqi=htmlencode(Request.form("S_Riqi"))
S_Riqi_2=htmlencode(Request.form("S_Riqi_2"))
S_JiaoShi=htmlencode(Request.form("S_JiaoShi"))

SQL="select * from tushushi where 1=1"
If Len(S_Riqi)<>0 AND Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi &"' and '" & S_Riqi_2 &"'"
If Len(S_Riqi)<>0 XOR Len(S_Riqi_2)<>0 Then SQL = SQL & " and Riqi between '" & S_Riqi & S_Riqi_2 &"' and '" & S_Riqi & S_Riqi_2 &"'"
If Len(S_JiaoShi)<>0 Then SQL = SQL & " and JiaoShi = '" & S_JiaoShi &"'"
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
			response.Write "<th><b>" & "日期" & "</b></th>"
		Case "JIAOSHI":
			response.Write "<th><b>" & "教师" & "</b></th>"
		Case "ROOM1":
			response.Write "<th><b>" & "一号阅览室人数" & "</b></th>"
		Case "ROOM2":
			response.Write "<th><b>" & "二号阅览室人数" & "</b></th>"
		Case "ROOM3":
			response.Write "<th><b>" & "三号阅览室人数" & "</b></th>"
		Case "STUDENT1":
			response.Write "<th><b>" & "第一组志愿者" & "</b></th>"
		Case "STUDENT2":
			response.Write "<th><b>" & "第二组志愿者" & "</b></th>"
		Case "STUDENT3":
			response.Write "<th><b>" & "第三组志愿者" & "</b></th>"
		Case "QINGKUANG":
			response.Write "<th><b>" & "开放情况" & "</b></th>"
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
response.write "结果页码："
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

response.write "<a href=""tushushi.asp?page=1"">第一页</a> "
for i=StartPage to EndPage
	response.write "<a href=""tushushi.asp?page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""tushushi.asp?page=" & PageCount & """>最后页</a> "
Else
	Response.Write "<h1>没有找到任何结果，请更改关键词，并重新搜索。</h1>"
End If
'End Select
MyRs.close
Set MyRs= Nothing
Conn.Close
set Conn=nothing
%>
<p align="left"><a href="download.asp?Action=tushushi">下载阅览室开放数据</a></p>
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