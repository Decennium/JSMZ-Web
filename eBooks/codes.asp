<%
Response.Charset = "gb2312"
Response.Buffer = True

StartTime=Timer()

q=request("q")
PageSize=20
Currentpage=request("page")

Set MyConn=Server.CreateObject("ADODB.Connection")
MyConn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;" & _
"DATA SOURCE=" & server.mappath("data/codes.mdb")
'MyConn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ="& _
'server.mappath("data/codes.mdb") 

MySQL="SELECT COUNT(*) As ResultCount FROM codes WHERE codes.code = '" & q & "' "
Set MyRs=MyConn.Execute(MySQL)
ResultCount=MyRs("ResultCount")

MySQL="SELECT * FROM codes WHERE codes.code = '" & q & "' "

'response.write MySQL
Set MyRs=MyConn.Execute(MySQL)
%>
<!--#include file="result_top.asp"-->
<%If ResultCount>0 then%>
<table style="border-style:solid; border-width:1px;" align="left" >
<thead>
<% 'Put Headings On The Table of Field Names
howmanyfields=MyRs.fields.count -1 %>
<% for i=0 to howmanyfields %>
	<td style="border-style:solid; border-width:1px; "><b><font color=""><%=MyRs(i).name %></font> </b></td>
<% next %>
</thead>
<tbody>
<% ' Get all the records
do while not MyRs.eof %>
<tr id='Data'>
<%
for i = 0 to howmanyfields
	ThisRecord = MyRs(I)
	If IsNull(ThisRecord) Then
		ThisRecord = "&nbsp;"
	end if
%>
	<td valign=top style="border-style:solid; border-width:1px; "><font color=""><%=Thisrecord%></font></td>
<% next %>
</tr>
<% MyRs.movenext
loop %>
</tbody>
</table>
<br clear=both>
<div id="navbar" align="left">

<div>
<%Else%>
<pre>
���屨������ȫ���� 

��ͬ�����屨������ͬ,����ǲο�˵����,��ȻҲ���Բο�����ͨ�õġ�

Award BIOS: 

1�̣�ϵͳ�������� 
2�̣�������󡣽������������BIOS 
1��1�̣�RAM��������� 
1��2�̣���ʾ������ʾ������ 
1��3�̣����̿��������� 
1��9�̣�����FLASH RAM ��EPROM����BIOS�� 
��ͣ���죨���������ڴ���δ������� 
�������죺��Դ����ʾ��δ���Կ����Ӻ� 
�ظ����죺��Դ������ 
����������ʾ����Դ������ 

AWI BIOS�� 

1�̣��ڴ�ˢ��ʧ�ܡ���������������ڴ��� 
2�̣��ڴ�ECCЧ����󡣽������������CMOS���ã���ECCЧ��ر� 
3�̣�ϵͳ�����ڴ棨��һ��64KB�����ʧ�� 
4�̣�ϵͳʱ�ӳ��� 
5�̣�CPU���� 
6�̣����̿��������� 
7�̣�ϵͳʵģʽ���󣬲����л�������ģʽ 
8�̣��Դ���� 
9�̣�ROM BIOS����ʹ��� 
1��3�̣��ڴ���� 
1��8�̣���ʾ���Դ��� 

Phoenix BIOS 

1�̣�ϵͳ�������� 
1��1��1�̣�ϵͳ�ӵ��Լ��ʼ��ʧ�� 
1��1��2�̣�������� 
1��1��3�̣�CMOS���ش��� 
1��1��4�̣�ROM BIOSЧ��ʧ�� 
1��2��1�̣�ϵͳʱ�Ӵ��� 
1��2��2�̣�DMA��ʼ��ʧ�� 
1��2��3�̣�DMAҳ�Ĵ������� 
1��3��1�̣�RAMˢ�´��� 
1��3��2�̣������ڴ���� 
1��4��1�̣������ڴ��ַ�ߴ��� 
1��4��2�̣������ڴ�Ч����� 
1��4��3�̣�EISAʱ�������� 
1��4��4�̣�EASA NMI�ڴ��� 
2��1��2�̵�2��4��4�̣������п�ʼΪ2�̵���������ϣ��������ڴ���� 
3��1��1�̣���DMA�Ĵ������� 
3��1��2�̣���DMA�Ĵ������� 
3��1��3�̣����жϴ���Ĵ������� 
3��1��4�̣����жϴ���Ĵ������� 
3��2��4�̣����̿��������� 
3��3��4�̣���ʾ���ڴ���� 
3��4��2�̣���ʾ���� 
3��4��3�̣�δ������ʾֻ���洢�� 
4��2��1�̣�ʱ�Ӵ��� 
4��2��2�̣��ػ����� 
4��2��3�̣�A20�Ŵ��� 
4��2��4�̣�����ģʽ�жϴ��� 
4��3��1�̣��ڴ���� 
4��3��3�̣�ʱ��2���� 
4��3��4�̣�ʵʱ�Ӵ��� 
4��4��1�̣����пڴ��� 
4��4��2�̣����пڴ��� 
4��4��3�̣�����Э���������� 

����BIOS�� 

1�̣�ϵͳ���� 
2�̣�ϵͳ�ӵ��Լ죨POST��ʧ�� 
1������Դ�����������ʾ����Ϊ��ʾ������ 

1��1�̣�������� 
1��2�̣��Կ����� 
1��1��1�̣���Դ���� 
3��1�̣����̴���
</pre>
<%
End If
If 1=2 Then '����result_bottom.asp�еĴ���
%>
<!--#include file="result_bottom.asp"-->
<%
MyRs.close
Set MyRs= Nothing
MyConn.Close
set MyConn=nothing
%>
	</body>
</html>

