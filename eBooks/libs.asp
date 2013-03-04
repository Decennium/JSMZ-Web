<html>
<head>
<meta http-equiv="content-type" content="text/html; charset=gb2312">
<title>校内其他图书库</title>
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
<p><a href="/elib/classic/menu/index.htm" target="Neirong">中国历史文化专辑</a></p>
<p><a href="/elib/10000Novels.HTML.2002/" target="Neirong">壹万册中文书库</a></p>
<p><a href="/elib/台大资工BBS精华区/" target="Neirong">台大资工BBS精华区</a></p>
<p><a href="/elibs/WenShiZhiShi/" target="Neirong">文史知识（1-234）</a></p>
</div>
</div>
<iframe src="/elib/classic/menu/index.htm"name="Neirong" id="Neirong" style="border:10px" width="85%" height="90%">
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