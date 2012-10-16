<!--#include file="top.asp"-->
<html>
<head>
<meta name="viewport" content="width=device-width,minimum-scale=1.0, maximum-scale=2.0"/>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>设备外借管理系统</title>
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
	if (the.JieQi.value=="") {
		document.getElementById('Tips').innerHTML = '借期不能为空';
		the.JieQi.focus();
		return false;
	}
	if (!My_IsInt(the.JieQi.value)){
		document.getElementById('Tips').innerHTML = '借期必须为自然数';
		the.JieQi.focus();
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
'		Sql="Select * from WaiJie where Riqi='" & Riqi & "' and ShenQingRen ='" & ShenQingRen& "' And SheBei='" & SheBei & "' And GuiHuan='" & "未归还'"
'		MyRs.open Sql,Conn,3,2
'		If MyRs.recordcount>0 then
'			Response.Write "<script>document.getElementById('Tips').innerHTML = '这个设备这个人已经借走，尚未归还。';</SCRIPT>"
''			MyRs.close
'		Else
			Sql="INSERT INTO [WaiJie] ([RiQi],[ShenQingRen],[SheBei],[JieQi],[MiaoShu],[FaFangRen],[GuiHuan],[BeiZhu]) VALUES ('"& RiQi &"','"& ShenQingRen &"','"& SheBei &"','"& JieQi &"','"& MiaoShu &"','" & FaFangRen &"','未归还','" & Beizhu &"')"
			conn.execute(Sql)
			'Response.Redirect "?Action=ShowJieci&page=" & Currentpage
'		End If
	End If

	'MyRs.Close
End If
If Action = "AddCheck" Then
'添加检查
	If Session("Admin")="" then
	'判断是否登陆
		Response.Redirect "waijie.asp"
		Response.End
	End If
	ID=htmlencode(Request.form("id"))
	GuiHuan=htmlencode(Request.form("GuiHuan"))
	GuiHuanRiQi=htmlencode(Request.form("GuiHuanRiQi"))
	GuiHuanRiQi=htmlencode(Request.form(GuiHuanRiQi))
	ZhuangKuang=htmlencode(Request.form("ZhuangKuang"))
	QianShouRen=htmlencode(Request.form("QianShouRen"))
	If (Len(ID)>0 And Len(GuiHuan)>0 And GuiHuan<>"未归还" And Len(GuiHuanRiQi)>0 And Len(QianShouRen)>0) then
		If MyRS.State = adStateClosed Then Set MyRs = Server.CreateObject("ADODB.RecordSet")
		If isNull(Conn) Then
			Set Conn=Server.CreateObject("ADODB.Connection")
			My_conn_STRING = "Provider=SQLOLEDB;server=S21;database=BOS;uid=sa;pwd="
			Conn.Open My_conn_STRING
		End If
			Sql="UPDATE [WaiJie] SET [GuiHuan] = '"& GuiHuan &"',[GuiHuanRiQi] ='"& GuiHuanRiQi &"',[QianShouRen] ='"& QianShouRen &"',[ZhuangKuang] ='"& ZhuangKuang &"' WHERE [id] ='"& ID &"'"
			conn.execute(Sql)
			Response.Redirect "?Action=ShowJieci&page=" & Currentpage
	End If
End If
%>
<div id="Right_Content" style="align:left;float:left">
<%If Session("Admin")<>"" then%>
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>添加设备外借记录</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Add_Area">
<form name="AddNewJieci" id="AddNewJieci" method="post" Action="?Action=AddRecord" onSubmit="return My_CheckFields(this);">
<span style="white-space: nowrap"><label for="Riqi">日期：</label><input type="text" name="Riqi" id="Riqi" size="10" readonly="readonly" onclick="choose_date_czw('Riqi')"/></span>
<span style="white-space: nowrap"><label for="ShenQingRen">申请人：</label><input type="text" name="ShenQingRen" value="" id="ShenQingRen" size="5" onblur="return My_CheckField(this);"></span>
<span style="white-space: nowrap"><label for="SheBei">设备：</label><input type="text" name="SheBei" value="" id="SheBei" size="20" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="JieQi">借期：</label><input type="text" name="JieQi" value="" id="JieQi" size="3" onblur="return My_CheckField(this);"/><label for="JieQi">天，</label></span>
<span style="white-space: nowrap"><label for="MiaoShu">设备描述：</label><input type="text" name="MiaoShu" value="一切正常" id="MiaoShu" size="20" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="FaFangRen">发放人：</label><input type="text" name="FaFangRen" value='<%=Session("ShowName")%>' id="FaFangRen" size="5" onblur="return My_CheckField(this);"/></span>
<span style="white-space: nowrap"><label for="Beizhu">备注：</label><input type="text" name="Beizhu" value="" id="Beizhu" size="10"/></span>
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
<div align="left" style="clear:left;float:left;nowrap;width:200px;margin:5px 100px 5px 100px"><strong>设备外借情况一览表</strong></div>
<br clear="all"/>
<div align="left" clear="all" id="Search_Area">
<form id="SearchJieci" name="SearchJieci" method="post" Action="?Action=Search" onSubmit="return My_CheckSearchDates(this);">
<span style="white-space: nowrap"><label for="S_Riqi">日期：从</label><input type="text" name="S_Riqi" id="S_Riqi" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/><label for="S_Riqi_2">到</label><input type="text" name="S_Riqi_2" id="S_Riqi_2" size="10" readonly="readonly" onclick="choose_date_czw(this.id)"/></span>
<span style="white-space: nowrap"><label for="S_ShenQingRen">申请人：</label><input name="S_ShenQingRen" id="S_ShenQingRen" type="text" value="" size="5"/></span>
<span style="white-space: nowrap"><label for="S_SheBei">设备：</label><input name="S_SheBei" id="S_SheBei" type="text" value="" size="10"/></span>
<span style="white-space: nowrap"><label for="S_GuiHuan">是否归还：</label>
<select name="S_GuiHuan" id="S_GuiHuan">
	<option value="%">不论</option>
	<option value="未归还" Selected="Selected">未归还</option>
	<option value="已归还">已归还</option>
	<option value="遗失">遗失</option>
	<option value="被盗">被盗</option>
</select></span>
<span style="white-space: nowrap"><label for="S_FaFangRen">发放人：</label><input name="S_FaFangRen" id="S_FaFangRen" type="text" value="" size="5"/></span>
<input type="submit" value="搜索" name="S_Submit" id="S_Submit"/>
</form>
</div>
<%
'搜索记录
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
If Len(S_GuiHuan)<>0 Then SQL = SQL & " and GuiHuan Like '" & S_GuiHuan &"'"
If Len(S_FaFangRen)<>0 Then SQL = SQL & " and FaFangRen = '" & S_FaFangRen &"'"
SQL = SQL & " order by Riqi desc"
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
		Case "SHENQINGREN":
			response.Write "<th><b>" & "申请人" & "</b></th>"
		Case "SHEBEI":
			response.Write "<th><b>" & "设备" & "</b></th>"
		Case "JIEQI":
			response.Write "<th width='50px'><b>" & "借期/天" & "</b></th>"
		Case "MIAOSHU":
			response.Write "<th class='NeiRong'><b>" & "描述" & "</b></th>"
		Case "FAFANGREN":
			response.Write "<th width='60px'><b>" & "发放人" & "</b></th>"
		Case "GUIHUAN":
			response.Write "<th width='60px'><b>" & "归还与否" & "</b></th>"
		Case "GUIHUANRIQI":
			response.Write "<th width='60px'><b>" & "归还日期" & "</b></th>"
		Case "ZHUANGKUANG":
			response.Write "<th class='NeiRong'><b>" & "归还时状况" & "</b></th>"
		Case "QIANSHOUREN":
			response.Write "<th width='60px'><b>" & "签收人" & "</b></th>"
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
		ThisRecord = MyRs(i_c).Value
		If IsNull(ThisRecord) Then
			ThisRecord = ""
		End if

		If Session("Admin")="" then
		'判断是否登陆
			Select Case Ucase(MyRs(i_c).Name)
			Case "JIEQI"
				Response.write("<td style='text-align:right;padding-right:10px'>" & ThisRecord & "</td>")
			Case "GUIHUAN"
				If ThisRecord = "未归还" Then
					Response.write("<td style='color:red'>" & ThisRecord & "</td>")
				Else
					Response.write("<td>" & ThisRecord & "</td>")
				End If
			Case Else
				Response.write("<td>" & ThisRecord & "</td>")
			End Select
		Else
			Select Case Ucase(MyRs(i_c).Name)
			Case "GUIHUAN"
				If ThisRecord = "未归还" Then
					Response.write("<form name='AddCheck' id='AddCheck' method='post' Action='?Action=AddCheck'><td><input type='hidden' name='page' value='" & Currentpage & "'/><select name=""GuiHuan"" id=""GuiHuan""><option value=""未归还"" Selected=""Selected"">未归还</option><option value=""已归还"">已归还</option><option value=""遗失"">遗失</option><option value=""被盗"">被盗</option></select></td>")
				Else
					Response.write("<td>" & ThisRecord & "</td>")
				End If
			Case "GUIHUANRIQI"
				If MyRs(7).Value = "未归还" Then
					GuiHuanRiQi = "GuiHuanQiRi" & replace(replace(replace(MyRs(0).Value,"-",""),"{",""),"}","")
					Response.write("<td><input type='hidden' name='id' value='" & MyRs(0).Value & "'/><input type='hidden' name='GuiHuanRiQi' value='" & GuiHuanRiQi & "'/><input type=""text"" name='" & GuiHuanRiQi & "' id='" & GuiHuanRiQi & "' value=""" & year(now()) &"-" & month(now()) & "-" & day(now()) & """ size=""10"" readonly=""readonly"" onclick=""choose_date_czw(this.id)""/></td>")
				Else
					Response.write("<td>" & ThisRecord & "</td>")
				End If
			Case "JIEQI"
				Response.write("<td style='text-align:right;padding-right:10px'>" & ThisRecord & "</td>")
			Case "ZHUANGKUANG"
				If MyRs(7).Value = "未归还" Then
					Response.write("<td><input type='text' name='ZhuangKuang' value='一切正常' size='20'/></td>")
				Else
					Response.write("<td>" & ThisRecord & "</td>")
				End If
			Case "QIANSHOUREN"
				If MyRs(7).Value = "未归还" Then
					Response.write("<td><input type='text' name='QianSHouRen' value='" & Session("ShowName") & "'  size='5'/><input type='submit' value='检查'/></form></td>")
				Else
					Response.write("<td>" & ThisRecord & "</td>")
				End If
			Case "BEIZHU"
				Response.write("<td class='BeiZhu'>" & ThisRecord & "</td>")
			Case Else
				Response.write("<td>" & ThisRecord & "</td>")
			End Select
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

response.write "<a href=""waijie.asp?page=1"">第一页</a> "
for i=StartPage to EndPage
	response.write "<a href=""waijie.asp?page=" & i & """>" & i & "</a> "
next 'i
response.write "<a href=""waijie.asp?page=" & PageCount & """>最后页</a> "
Else
	Response.Write "<h1>没有找到任何结果，请更改关键词，并重新搜索。</h1>"
End If
'End Select
'MyRs.close
'Set MyRs= Nothing
Conn.Close
set Conn=nothing
%>
<p align="left"><a href="download.asp?Action=waijie">下载设备外借数据</a></p>
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