<!--#include file="inc/conn.inc"-->
<html>
<head>
<title>ʵ����ԤԼϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
</head>
<body>
<table border=0 cellSpacing=1 cellPadding=3 bgcolor=#000000>
<tbody bgcolor=#ffffff>
<tr><th>�γ�����</th><th width=70>�γ̱��</th><th width=70>�ο�<br>��ʦ</th><th width=100>�༶</th><th width=20>ѧʱ</th><th width=30>����</th><th width=90>����һ</th><th width=90>���ڶ�</th><th width=90>������</th><th width=90>������</th><th width=90>������</th><th width=90>������</th><th width=90>������</th><th width=90>ʵ����</th></tr>
<%
set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select * from kc order by �γ���"
Rs.Open Sqlstr,conn,2,3  
if rs.eof then response.write "no"
do while not rs.eof
response.write "<tr><td> " & rs(1)&"</td><td> " & rs(2)&"</td><td> "&rs(3)&"</td><td> "& rs(5)&"</td><td>&nbsp;"&rs(4)&"</td><td> "& rs(6)&"</td><td> "&rs("����һ")&"</td><td> "&rs("���ڶ�")&"</td><td> "&rs("������")&"</td><td>"&rs("������")&" </td><td>"&rs("������")&" </td><td>"&rs("������")&" </td><td>"&rs("������")&" </td><td>"&rs("ʵ���ҵص�")&" </td></tr>"

rs.movenext
loop

%>
</tbody>
</table>
</body>
</html>