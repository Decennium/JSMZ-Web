<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>校内资源系统</title>
<link href="../css/css.css" rel="stylesheet">
<link rel="shortcut icon" type="image/x-icon" href="favicon.ico">
</head>
<body>
<!--#include file="../include/banner.asp"-->
<div name="Left_Banner" id="Left_Banner" align="left" style="clear:left;float:left;margin-right:0px" width="150px">
<div style="clear:left;float:left" width="150px">
	<a href="http://192.168.2.1/">
	<img src="../images/jsmz_logo_mini.png" alt="JSMZ" border="0" align="left">
	</a>
</div>

<div style="clear:left;float:left" width="150px">
	<p><a href="/Resources/ClassWare/" target='Neirong' onclick="javascript:document.getElementById('area').value='Classware';">校内课件资源</a></p>
	<p><%Response.Write "<a href=" & Chr(34) & "http://" & Request.ServerVariables("Local_Addr") & ":8080/" & Chr(34) & " target='Neirong' onclick='javascript:document.getElementById(" & Chr(34) & "area" & Chr(34) & ").value=" & Chr(34) & "Intranet" & Chr(34) & ";'>内部共享资源</a>"%></p>
	<p><a href="http://192.168.2.6/" target='Neirong' onclick="javascript:document.getElementById('area').value='Video';">校内视频资源</a></p>
<hr width="150px">
	<p><a href="/Resources/XiaoXueSuCai/" target='Neirong' onclick="javascript:document.getElementById('area').value='XiaoXueSuCai';">小学课件素材</a></p>
	<p><a href="/yuanjiao/" target='Neirong' onclick="javascript:document.getElementById('area').value='Yuanjiao';">初中远教资源</a></p>
	<p><a href="/Resources/ZheDa/" target='Neirong' onclick="javascript:document.getElementById('area').value='ZheDa';">浙江大学课件资源</a></p>
	<p><a href="/Resources/JHTJ/" target='Neirong' onclick="javascript:document.getElementById('area').value='JHTJ';">江海天骄资源库</a></p>
	<p><a href="http://192.168.2.5/Res/xj.asp" target='Neirong' onclick="javascript:document.getElementById('area').value='XJ';">鑫剑资源搜索系统</a></p>
	<p><a href="/Resources/JSGZ/" target='Neirong' onclick="javascript:document.getElementById('area').value='JSGZ';">江苏高中优质资源</a></p>
	<p><a href="/Resources/GZHX/" target='Neirong' onclick="javascript:document.getElementById('area').value='GZHXSY';">高中化学演示实验</a></p>
	<p><a href="/Resources/CJWX/" target='Neirong' onclick="javascript:document.getElementById('area').value='CJKJ';">长郡高中课件全集</a></p>
<!--
	<p><a href="/Resources/ZheDa/" target='Neirong' onclick="javascript:document.getElementById('area').value='ZheDa';">浙江大学课件资源</a></p>
-->
</div>
</div>
<iframe src="/Resources/ClassWare/"name="Neirong" id="Neirong" frameBorder="0" style="border:none" width="80%" height="90%" align="left">
</iframe>
<br clear="both">
<!--#include file="../include/bottom.asp"-->
<script language="javascript">
c_Width =document.body.offsetWidth - parseInt(document.getElementById('Left_Banner').style.width)-150;
document.getElementById('Neirong').style.width="" +c_Width +"px";
document.getElementById('Neirong').style.maxWidth="" + c_Width  + "px";
c_Height=document.body.offsetHeight;// - parseInt(document.getElementById('TopBanner').style.height);
document.getElementById('Neirong').style.height="" +c_Height -50 +"px";
document.getElementById('Neirong').style.maxHeight="" + c_Height  -50 + "px";
//document.write('<p>' + document.getElementById('Neirong').style.height +'</p>');
</script>
</body>
</html>