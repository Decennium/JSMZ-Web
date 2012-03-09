<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>吉首市民族中学资源库</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../CSS/ebooks.css" rel="stylesheet" type="text/css">
</head>

<body style="margin-left:10px;margin-right:10px;">
<%
Response.Charset = "gb2312"
Response.Buffer = True

StartTime=Timer()

q=request("q")
area=ucase(request("area"))
Response.Write "<p>目前仅返回" & area & "资源内符合条件的前20条记录</p>"
Const MaxResult = 20
select case area
case "CLASSWARE"
	ResultCount = 0
	Response.Write "<p>校内课件资源</p>"
	Response.Write "<ol>"
	bianli "D:\ClassWare",q 'http://192.168.2.1/Resources/ClassWare/
	Response.Write "</ol>"
Case "ZHEDA"
	ResultCount = 0
	Response.Write "<p>搜索浙江大学课件资源请到电子图书库</p>"
	' Response.Write "<ol>"
	' bianli "E:\fix_source",q 'http://192.168.2.1/Resources/ZheDa/
	' Response.Write "</ol>"
Case "INTRANET"
	ResultCount = 0
	Response.Write "<p>内部共享资源</p>"
	Response.Write "<ol>"
	bianli "D:\内部共享",q 'http://192.168.2.1:8080/
	Response.Write "</ol>"
Case "YUANJIAO"
	ResultCount = 0
	Response.Write "<p>初中远教资源不支持搜索，请直接访问</p>"
	' Response.Write "<ol>"
	' bianli "D:\农村远程教育资源初中",q 'http://192.168.2.1/yuanjiao/
	' Response.Write "</ol>"
Case Else
End Select

Function bianli(path,str)
	'set fso=server.CreateObject( "scripting.filesystemobject")
	on error resume next
	set fso=CreateObject( "scripting.filesystemobject")
	set objFolder=fso.GetFolder(path)
	set objSubFolders=objFolder.Subfolders

	nowpath=path
	set objFiles=objFolder.Files
	for each objFile in objFiles
		CheckAndOutput objFile.Name,str,nowpath
		if ResultCount >= MaxResult Then Exit Function
'		ResultCount = ResultCount +1
	next

	for each objSubFolder in objSubFolders
		nowpath=path + "\" + objSubFolder.name
		set objFiles=objSubFolder.Files
		for each objFile in objFiles
			CheckAndOutput objFile.Name,str,nowpath
			if ResultCount >= MaxResult Then Exit Function
'			ResultCount = ResultCount +1
		next
		bianli nowpath,str '递归
	next
	set objFolder=nothing
	set objSubFolders=nothing
	set fso=nothing
end function
Sub CheckAndOutput(byval FileName,byval strSearch,byval strPath)
	if instr(ucase(strPath & "\" & FileName),ucase(strSearch))>0 then
		URL = "<li><a href='" & strPath & "\" & FileName & "' title='" & strPath & "\" & FileName & "'>"  & strPath & "\" & FileName &  "</a></li>"
		URL=replace(URL,"D:\ClassWare","/Resources/ClassWare")
		URL=replace(URL,"E:\fix_source","/Resources/ZheDa")
		URL=replace(URL,"D:\内部共享","http://192.168.2.1:8080")
		URL=replace(URL,"D:\农村远程教育资源初中","/yuanjiao")
		URL=Replace(URL,"\","/")
		Response.Write URL
		ResultCount = ResultCount +1
	end if
End Sub
'http://topic.csdn.net/u/20070724/20/aa99c396-cdf4-47f9-b100-0405fbb38ec0.html
%>
</body>
</html>