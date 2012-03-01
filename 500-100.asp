<%@ language="VBScript" @codepage=936%>
<%
  Option Explicit
  response.charset="gb2312"
  Const lngMaxFormBytes = 200

  Dim objASPError, blnErrorWritten, strServername, strServerIP, strRemoteIP
  Dim strMethod, lngPos, datNow, strQueryString, strURL

  If Response.Buffer Then
    Response.Clear
    Response.Status = "500 Internal Server Error"
    Response.ContentType = "text/html"
    Response.Expires = 0
  End If

  Set objASPError = Server.GetLastError
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2 Final//EN">

<html dir=ltr>

<head>
<style>
a:link			{font:9pt 宋体; color:FF0000}
a:visited		{font:9pt 宋体; color:#4e4e4e}
</style>

<META NAME="ROBOTS" CONTENT="NOINDEX">

<title>本页无法显示</title>

<META HTTP-EQUIV="Content-Type" Content="text-html; charset=gb2312">
<META NAME="MS.LOCALE" CONTENT="ZH-CN">
</head>

<script> 
function Homepage(){
<!--
// in real bits, urls get returned to our script like this:
// res://shdocvw.dll/http_404.htm#http://www.DocURL.com/bar.htm 

	//For testing use DocURL = "res://shdocvw.dll/http_404.htm#https://www.microsoft.com/bar.htm"
	DocURL=document.URL;
	
	//this is where the http or https will be, as found by searching for :// but skipping the res://
	protocolIndex=DocURL.indexOf("://",4);
	
	//this finds the ending slash for the domain server 
	serverIndex=DocURL.indexOf("/",protocolIndex + 3);

	//for the href, we need a valid URL to the domain. We search for the # symbol to find the begining 
	//of the true URL, and add 1 to skip it - this is the BeginURL value. We use serverIndex as the end marker.
	//urlresult=DocURL.substring(protocolIndex - 4,serverIndex);
	BeginURL=DocURL.indexOf("#",1) + 1;
	urlresult=DocURL.substring(BeginURL,serverIndex);
		
	//for display, we need to skip after http://, and go to the next slash
	displayresult=DocURL.substring(protocolIndex + 3 ,serverIndex);
	InsertElementAnchor(urlresult, displayresult);
}

function HtmlEncode(text)
{
    return text.replace(/&/g, '&amp').replace(/'/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function TagAttrib(name, value)
{
    return ' '+name+'="'+HtmlEncode(value)+'"';
}

function PrintTag(tagName, needCloseTag, attrib, inner){
    document.write( '<' + tagName + attrib + '>' + HtmlEncode(inner) );
    if (needCloseTag) document.write( '</' + tagName +'>' );
}

function URI(href)
{
    IEVer = window.navigator.appVersion;
    IEVer = IEVer.substr( IEVer.indexOf('MSIE') + 5, 3 );

    return (IEVer.charAt(1)=='.' && IEVer >= '5.5') ?
        encodeURI(href) :
        escape(href).replace(/%3A/g, ':').replace(/%3B/g, ';');
}

function InsertElementAnchor(href, text)
{
    PrintTag('A', true, TagAttrib('HREF', URI(href)), text);
}

//-->
</script>

<body bgcolor="FFFFFF">

<table width="410" cellpadding="3" cellspacing="5">

  <tr>    
    <td align="left" valign="middle" width="360">
	<h1 style="COLOR:000000; FONT: 9pt 宋体"><!--Problem-->本页无法显示</h1>
    </td>
  </tr>
  
  <tr>
    <td width="400" colspan="2">
	<font style="COLOR:000000; FONT: 9pt 宋体">试图访问的网页出现问题，无法显示。</font></td>
  </tr>
  
  <tr>
    <td width="400" colspan="2">
	<font style="COLOR:000000; FONT: 9pt 宋体">

	<hr color="#C0C0C0" noshade>
	
    <p>请尝试以下方法：</p>

	<ul>
      <li id="instructionsText1">单击 
      <a href="javascript:location.reload()">
      刷新</a>按钮或者梢候再试。<br>
      </li>
	  
      <li>打开 
	  
	  <script>
	  <!--
	  if (!((window.navigator.userAgent.indexOf("MSIE") > 0) && (window.navigator.appVersion.charAt(0) == "2")))
	  {
	  	 Homepage();
	  }
	  //-->
	  </script>

	  主页，然后查找与所需信息相关的链接。 </li>
    </ul>
	
    <h2 style="FONT: 9pt 宋体; color:000000">HTTP 500.100 - 内部服务器错误 - ASP 错误<br>
    Internet 信息服务</h2>

	<hr color="#C0C0C0" noshade>
	
	<p>技术信息（适用于支持人员）</p>

<ul>
<li>错误类型：<br>
<%
  Dim bakCodepage
  bakCodepage = Session.Codepage
  Session.Codepage = 936
  Response.Write Server.HTMLEncode(objASPError.Category)
  If objASPError.ASPCode > "" Then Response.Write Server.HTMLEncode(", " & objASPError.ASPCode)
  Response.Write Server.HTMLEncode(" (0x" & Hex(objASPError.Number) & ")" ) & "<br>"

  If objASPError.ASPDescription > "" Then Response.Write Server.HTMLEncode(objASPError.ASPDescription) & "<br>"

  blnErrorWritten = False

  ' Only show the Source if it is available and the request is from the same machine as IIS
  If objASPError.Source > "" Then
    strServername = LCase(Request.ServerVariables("SERVER_NAME"))
    strServerIP = Request.ServerVariables("LOCAL_ADDR")
    strRemoteIP =  Request.ServerVariables("REMOTE_ADDR")
    If (strServername = "localhost" Or strServerIP = strRemoteIP) And objASPError.File <> "?" Then
      Response.Write Server.HTMLEncode(objASPError.File)
      If objASPError.Line > 0 Then Response.Write ", 第 " & objASPError.Line & " 行"
      If objASPError.Column > 0 Then Response.Write ", 第 " & objASPError.Column & " 列"
      Response.Write "<br>"
      Response.Write "<font style=""COLOR:000000; FONT: 9pt 宋体""><b>"
      Response.Write Server.HTMLEncode(objASPError.Source) & "<br>"
      If objASPError.Column > 0 Then Response.Write String((objASPError.Column - 1), "-") & "^<br>"
      Response.Write "</b></font>"
      blnErrorWritten = True
    End If
  End If

  If Not blnErrorWritten And objASPError.File <> "?" Then
    Response.Write "<b>"
    Response.Write Server.HTMLEncode(objASPError.File)
    If objASPError.Line > 0 Then Response.Write Server.HTMLEncode(", 第 " & objASPError.Line & " 行")
    If objASPError.Column > 0 Then Response.Write ", 第 " & objASPError.Column & " 列"
    Response.Write "</b><br>"
  End If
%>
</li>
<p>
<li>浏览器类型：<br>
<%= Server.HTMLEncode(Request.ServerVariables("HTTP_USER_AGENT")) %>
</li>
<p>
<li>页：<br>
<%
  strMethod = Request.ServerVariables("REQUEST_METHOD")

  Response.Write strMethod & " "

  If strMethod = "POST" Then
    Response.Write Request.TotalBytes & " bytes to "
  End If

  Response.Write Request.ServerVariables("SCRIPT_NAME")

  lngPos = InStr(Request.QueryString, "|")

  If lngPos > 1 Then
    Response.Write "?" & Server.HTMLEncode(Left(Request.QueryString, (lngPos - 1)))
  End If

  Response.Write "</li>"

  If strMethod = "POST" Then
    Response.Write "<p><li>POST 数据:<br>"
    If Request.TotalBytes > lngMaxFormBytes Then
       Response.Write Server.HTMLEncode(Left(Request.Form, lngMaxFormBytes)) & " . . ."
    Else
      Response.Write Server.HTMLEncode(Request.Form)
    End If
    Response.Write "</li>"
  End If

%>
<p>
<li>时间：<br>
<%
  datNow = Now()

  Response.Write Server.HTMLEncode(FormatDateTime(datNow, 1) & ", " & FormatDateTime(datNow, 3))
  Session.Codepage = bakCodepage
%>
</li>
</p>
<p>
<li>详细信息：<br>
 <%  strQueryString = "prd=iis&sbp=&pver=5.0&ID=500;100&cat=" & Server.URLEncode(objASPError.Category) & _
    "&os=&over=&hrd=&Opt1=" & Server.URLEncode(objASPError.ASPCode)  & "&Opt2=" & Server.URLEncode(objASPError.Number) & _
    "&Opt3=" & Server.URLEncode(objASPError.Description) 
       strURL = "http://www.microsoft.com/ContentRedirect.asp?" & _
    strQueryString
%>
<a href="<%= strURL %>">Microsoft 支持</a>
</li>
</p>

    </font></td>
  </tr>
  
</table>
</body>
</html>
