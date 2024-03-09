<!--#include file="inc/conn.inc"-->
<html>
<head>
<title>实验室预约系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
</head>
<body>
<table border=0 cellSpacing=1 cellPadding=4 bgcolor=#000000>
<tbody bgcolor=#ffffff>
<tr><th width=50>实验室编号</th><th width=230>实验室名称</th><th width=150>地址</th><th>容量</th><th width=360>承担课程</th><th>面积</th><th width=200>备注</th><th>操作</th></tr>
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