<!--#include file="inc/conn.inc"-->
<html>
<head>
<title>实验室预约系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
</head>
<body>
<table border=0 cellSpacing=1 cellPadding=3 bgcolor=#000000>
<tbody bgcolor=#ffffff>
<tr><th>课程名称</th><th width=70>课程编号</th><th width=70>任课<br>老师</th><th width=100>班级</th><th width=20>学时</th><th width=30>人数</th><th width=90>星期一</th><th width=90>星期二</th><th width=90>星期三</th><th width=90>星期四</th><th width=90>星期五</th><th width=90>星期六</th><th width=90>星期天</th><th width=90>实验室</th></tr>
<%
set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select * from kc order by 课程名"
Rs.Open Sqlstr,conn,2,3  
if rs.eof then response.write "no"
do while not rs.eof
response.write "<tr><td> " & rs(1)&"</td><td> " & rs(2)&"</td><td> "&rs(3)&"</td><td> "& rs(5)&"</td><td>&nbsp;"&rs(4)&"</td><td> "& rs(6)&"</td><td> "&rs("星期一")&"</td><td> "&rs("星期二")&"</td><td> "&rs("星期三")&"</td><td>"&rs("星期四")&" </td><td>"&rs("星期五")&" </td><td>"&rs("星期六")&" </td><td>"&rs("星期日")&" </td><td>"&rs("实验室地点")&" </td></tr>"

rs.movenext
loop

%>
</tbody>
</table>
</body>
</html>