<!--#include file="inc/conn.inc"-->
<html>
<head>
<title>ʵ����ԤԼϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
</head>
<body>
<table border=0 cellSpacing=1 cellPadding=4 bgcolor=#000000>
<tbody bgcolor=#ffffff>
<tr><th width=50>ʵ���ұ��</th><th width=230>ʵ��������</th><th width=150>��ַ</th><th>����</th><th width=360>�е��γ�</th><th>���</th><th width=200>��ע</th><th>����</th></tr>
<%
set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select * from lab where lab_id<22"
Rs.Open Sqlstr,conn,2,3  
if rs.eof then response.write "no"
do while not rs.eof
response.write "<tr><td>" & rs("lab_id")&"</td><td>" & rs("lab_name")&"</td><td>"&rs("lab_address")&"</td><td>"& rs("lab_capacity")&"</td><td>"&rs("lab_course")&"</td><td>"&rs("lab_area")&"</td><td>"&rs("lab_remarks")&"</td><td></td></tr>"

rs.movenext
loop

%>
</tbody>
</table>
</body>
</html>