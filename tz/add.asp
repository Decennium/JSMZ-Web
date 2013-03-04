<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>发布通知</title>
<meta name="viewport" content="width=device-width,minimum-scale=1.0, maximum-scale=2.0"/>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<link href="../css/css.css" rel="stylesheet">
</head>

<body>
<!--#include file="../include/banner.asp"-->
<div style="float:left;">
<!--#include file="../include/left_banner.asp"-->
</div>
<div id="Right_Content">
<div class="ShowTips" style="vertical-align:top; ">
<form name="form1" method="post" action="save.asp">
标　题：<input name="BiaoTi" type="text" id="BiaoTi" size="30" maxlength="50"> <br>
发布者：<input name="ZuoZhe" type="text" id="ZuoZhe" size="30" maxlength="30" value=<%=Session("ShowName")%> readonly><br>
关键词：<input name="GuanJianCi" type="text" id="GuanJianCi" size="30" maxlength="50"> <br>
内　容：<textarea name="NeiRong" cols="60" rows="8" id="NeiRong"></textarea> <br> 
<br><input type="submit" name="Submit" value="提交通知">
</form>
</div>
</div>
<br clear="all"><br><br><br><br><br><br><br><br><br>
<!--#include file="../include/bottom.asp"-->
</body>
</html>
