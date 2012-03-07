<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>校内资源系统</title>
<link href="../css/css.css" rel="stylesheet">
<link rel="shortcut icon" type="image/x-icon" href="favicon.ico">
</head>
<body>
<!--#include file="../include/banner.asp"-->
<div name="Left_Banner" id="Left_Banner" align="left" style="clear:left;float:left;margin-right:10px" width="150px">
<div style="clear:left;float:left" width="150px">
	<a href="http://192.168.2.1/">
	<img src="../images/jsmz_logo_mini.png" alt="JSMZ" border="0" align="left">
	</a>
</div>
<br clear="both"/>
<div style="clear:left;float:left" width="150px">
	<p><a href="/Resources/ClassWare/" target='Neirong' onclick="javascript:document.getElementById('area').value='Classware';">校内课件资源</a></p>
	<p><a href="/yuanjiao/" target='Neirong' onclick="javascript:document.getElementById('area').value='Yuanjiao';">初中远教资源</a></p>
	<p><a href="/Resources/ZheDa/" target='Neirong' onclick="javascript:document.getElementById('area').value='ZheDa';">浙江大学课件资源</a></p>
	<p><%Response.Write "<a href=" & Chr(34) & "http://" & Request.ServerVariables("Local_Addr") & ":8080/" & Chr(34) & " target='Neirong' onclick='javascript:document.getElementById(" & Chr(34) & "area" & Chr(34) & ").value=Intranet;'>内部共享资源</a>"%></p>
</div>
</div>
<form action="searchresources.asp" id="sr" name="sr" method="post" target="Neirong">
<div id="SearchBar" name="Searchbar" align="left">
<input type="hidden" value="Classware" id="area" name="area">
<input maxlength="250" id="q" name="q" size="55" value="" onmouseover="this.select()">
<input name="btnS" type="submit" value="开始搜索">
</div>
</form>
<iframe src="/Resources/ClassWare/"name="Neirong" id="Neirong" style="border:10px" width="90%" height="90%">
</iframe>
<!--#include file="../include/bottom.asp"-->
<script language="javascript">
c_Width =document.body.offsetWidth - parseInt(document.getElementById('Left_Banner').style.width)-40;
document.getElementById('Neirong').style.width="" +c_Width +"px";
document.getElementById('Neirong').style.maxWidth="" + c_Width  + "px";
c_Height=document.body.offsetHeight;// - parseInt(document.getElementById('TopBanner').style.height);
document.getElementById('Neirong').style.height="" +c_Height -50 +"px";
document.getElementById('Neirong').style.maxHeight="" + c_Height  -50 + "px";
//document.write('<p>' + document.getElementById('Neirong').style.height +'</p>');
</script>
</body>
</html>