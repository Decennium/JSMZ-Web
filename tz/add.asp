<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>����֪ͨ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../css/css.css" rel="stylesheet">
</head>

<body>
<!--#include file="../include/banner.asp"-->
<div style="float:left;">
<!--#include file="../include/left_banner.asp"-->
</div>
<div style="width="100%";float:left">
<div class="ShowTips">
<strong>[�ύ֪ͨ]</strong>����<a href="index.asp">�鿴֪ͨ</a> 
<form name="form1" method="post" action="save.asp">
�����ߣ�<input name="ZuoZhe" type="text" id="ZuoZhe" size="30" maxlength="30" value=<%=Session("ShowName")%> readonly><br>
�ꡡ�⣺<input name="BiaoTi" type="text" id="BiaoTi" size="30" maxlength="50"> <br>
�ؼ��ʣ�<input name="GuanJianCi" type="text" id="GuanJianCi" size="30" maxlength="50"> <br>
�ڡ��ݣ�<textarea name="NeiRong" cols="60" rows="8" id="NeiRong"></textarea> <br> 
<br><input type="submit" name="Submit" value="�ύ֪ͨ">
</form>
</div>
</div>
<br clear="all"><br><br><br><br><br><br><br><br><br>
<!--#include file="../include/bottom.asp"-->
</body>
</html>