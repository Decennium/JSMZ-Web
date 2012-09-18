<!--#include file="top.asp"-->
<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>设备购入管理系统</title>
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
function My_IsNature(val) {
	return /^[1-9][0-9]*$/.test(val);
}
function My_IsFloat(val) {
	return /^\d+(\.\d+)?$/.test(val);
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
	if (the.SheBei.value=="") {
		document.getElementById('Tips').innerHTML = '设备名称不能为空';
		the.SheBei.focus();
		return false;
	}
	if (the.PinPai.value=="") {
		document.getElementById('Tips').innerHTML = '设备品牌不能为空';
		the.PinPai.focus();
		return false;
	}
	if (the.ShuLiang.value=="") {
		document.getElementById('Tips').innerHTML = '数量不能为空';
		the.ShuLiang.focus();
		return false;
	}
	if (!My_IsNature(the.ShuLiang.value)){
		document.getElementById('Tips').innerHTML = '数量必须为自然数';
		the.ShuLiang.focus();
		return false;
	}
	if (the.DanJia.value=="") {
		document.getElementById('Tips').innerHTML = '单价不能为空';
		the.DanJia.focus();
		return false;
	}
	if (!My_IsFloat(the.DanJia.value)){
		document.getElementById('Tips').innerHTML = '单价必须为正数';
		the.DanJia.focus();
		return false;
	}
	if (the.YongTu.value=="") {
		document.getElementById('Tips').innerHTML = '用途不能为空';
		the.YongTu.focus();
		return false;
	}
	if (the.JingShouRen.value=="") {
		document.getElementById('Tips').innerHTML = '经手人不能为空';
		the.JingShouRen.focus();
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
				Response.Write "<script>document.getElementById('Tips').innerHTML = '这款设备的购买情况已经登记。';</SCRIPT>"
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
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>添加设备购入记录</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Add_Area">
<form name="AddNewGouRu" id="AddNewGouRu" method="post" Action="?Action=AddRecord" onSubmit="return My_CheckFields(this);">
<span style="white-space: nowrap"><label for="Riqi">日期：</label><input type="text" name="Riqi" id="Riqi" size="10" readonly="readonly" onclick="choose_date_czw('Riqi')"/></span>
<span style="white-space: nowrap"><label for="SheBei">设备：</label><input type="text" name="SheBei" value="" id="SheBei" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="PinPai">品牌：</label><input type="text" name="PinPai" value="无" id="PinPai" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="XingHao">型号：</label><input type="text" name="XingHao" value="无" id="XingHao" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="XuLieHao">序列号：</label><input type="text" name="XuLieHao" value="无" id="XuLieHao" size="20" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="DanWei">单位：</label>
<select name="DanWei" id="DanWei">
	<option value="套">套</option>
	<option value="盒">盒</option>
	<option value="台">台</option>
	<option value="只">只</option>
	<option value="支">支</option>
	<option value="米">米</option>
	<option value="瓶">瓶</option>
	<option value="包">包</option>
	<option value="片">片</option>
	<option value="条">条</option>
	<option value="粒">粒</option>
	<option value="个">个</option>
</select></span>
<span style="white-space: nowrap"><label for="ShuLiang">数量：</label><input type="text" name="ShuLiang" value="1" id="ShuLiang" size="5" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="DanJia">单价：</label><input type="text" name="DanJia" value="" id="DanJia" size="5" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="JingShouRen">经手人：</label><input type="text" name="JingShouRen" value=<%=Session("ShowName")%> id="JingShouRen" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="YongTu">用途：</label><input type="text" name="YongTu" value="" id="YongTu" size="20" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="OS">操作系统：</label><input type="text" name="OS" value="无" id="OS" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="OSXuLieHao">系统序列号：</label><input type="text" name="OSXuLieHao" value="无" id="OSXuLieHao" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="Beizhu">备注：</label><input type="text" name="Beizhu" value="" id="Beizhu" size="30"/></span>
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
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>设备购入情况一览表</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Search_Area">
<form id="SearchSheBei" name="SearchSheBei" method="post" Action="?Action=Search" onSubmit="return My_CheckSearchDates(this);">
<span style="white-space: nowrap"><label for="S_Riqi">日期：从</label><input type="text" name="S_Riqi" id="S_Riqi" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/><label for="S_Riqi_2">到</label><input type="text" name="S_Riqi_2" id="S_Riqi_2" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/></span>
<span style="white-space: nowrap"><label for="S_SheBei">设备：</label><input type="text" name="S_SheBei" value="" id="S_SheBei" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="S_PinPai">品牌：</label><input type="text" name="S_PinPai" value="" id="S_PinPai" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="S_XingHao">型号：</label><input type="text" name="S_XingHao" value="" id="S_XingHao" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="S_XuLieHao">序列号：</label><input type="text" name="S_XuLieHao" value="" id="S_XuLieHao" size="20" onblur="return My_CheckField(this);"></span>
<!--  -->
<span style="white-space: nowrap"><label for="S_JingShouRen">经手人：</label><input type="text" name="S_JingShouRen" value="" id="S_JingShouRen" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="S_YongTu">用途：</label><input type="text" name="S_YongTu" value="" id="S_YongTu" size="20" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="S_OS">操作系统：</label><input type="text" name="S_OS" value="" id="S_OS" size="10" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="S_OSXuLieHao">系统序列号：</label><input type="text" name="S_OSXuLieHao" value="" id="S_OSXuLieHao" size="10" onblur="return My_CheckField(this);"></span>
<input type="submit" value="搜索" name="S_Submit" id="S_Submit"/>
</form>
</div>
<%
'搜索记录

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
			response.Write "<th width='70px'><b>" & "日期" & "</b></th>"
		Case "SHEBEI":
			response.Write "<th><b>" & "设备" & "</b></th>"
		Case "PINPAI":
			response.Write "<th><b>" & "品牌" & "</b></th>"
		Case "XINGHAO":
			response.Write "<th><b>" & "型号" & "</b></th>"
		Case "XULIEHAO":
			response.Write "<th><b>" & "序列号" & "</b></th>"
		Case "DANWEI":
			response.Write "<th width='30px'><b>" & "单位" & "</b></th>"
		Case "SHULIANG":
			response.Write "<th width='30px'><b>" & "数量" & "</b></th>"
		Case "DANJIA":
			response.Write "<th><b>" & "单价" & "</b></th>"
			response.Write "<th><b>" & "总价" & "</b></th>"
		Case "JINGSHOUREN":
			response.Write "<th width='50px'><b>" & "经手人" & "</b></th>"
		Case "YONGTU":
			response.Write "<th><b>" & "用途" & "</b></th>"
		Case "OS":
			response.Write "<th width='60px'><b>" & "操作系统" & "</b></th>"
		Case "OSXULIEHAO":
			response.Write "<th width='80px'><b>" & "系统序列号" & "</b></th>"
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
'	If MyRs.EOF Then Exit For
	if 1 = i_s mod 2 then
		response.write("<tr id='Data' class='odd'>")
	else
		response.write("<tr id='Data'>")
	end if
	for i_c = 1 to howmanyfields '不显示Id字段
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
response.write "结果页码："
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

response.write "<a href=""gouru.asp?page=1"">第一页</a> "
for i=StartPage to EndPage
	response.write "<a href=""gouru.asp?page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""gouru.asp?page=" & PageCount & """>最后页</a> "
Else
	Response.Write "<h1>没有找到任何结果，请更改关键词，并重新搜索。</h1>"
End If
'End Select
'MyRs.close
'Set MyRs= Nothing
Conn.Close
set Conn=nothing
%>
<p align="left"><a href="download.asp?Action=gouru">下载设备购入数据</a></p>
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