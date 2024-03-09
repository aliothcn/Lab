<%if session("admin")="" or session("adminid")<>"ALIOTH@126.COM" then
  response.redirect "login.asp"
response.end
else
  adminname=session("admin")
end if%>
<!--#include file="inc/conn.inc"-->
<html>
<head>
<title>实验室预约系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"/>
<script LANGUAGE="JavaScript"> 
<!-- 
function load(form) { 
var url = form.list.options[form.list.selectedIndex].value; 
if (url != "") open(url, "main"); 
return false; 
} 

// --> 
</script> 
</head>
<body>
<table cellSpacing=1 cellPadding=6 border=0 bgcolor=#000000 align=center width=1250>
<tbody bgcolor=#ffffff align=center>
<tr><form NAME="menu" onSubmit="return load(this)">
<td bgcolor=#ffffff align=left><%=adminname%>&nbsp;&nbsp;你好!&nbsp;&nbsp;请选择教师：
<%
set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select distinct 教师名 from kc"
%>
	<select name="list" id="selectcourse" onChange="load(this.form)">
<option selected value="">请选择</option>
<%
Rs.Open Sqlstr,conn,2,3  
Do While Not Rs.EOF 
Response.Write "<option value='teacher.asp?teacher="& rs("教师名") &"'>" & rs("教师名") & "</option>"
Rs.MoveNext
Loop 
rs.close
set rs=nothing%>
</select>
</td></form>
</tr>
<%if request("teacher")<>"" then%>
<tr>
<td bgcolor=#ffffff>
<%
sqlstr="select * from kc where 教师名='"&request("teacher")&"'"
'response.write sqlstr
Dim kb(6,4) '第一维表示周几,第二维表示第几堂课
	set rs=conn.execute(sqlstr)
	do while not Rs.eof
'response.write rs(0)&"$$"&rs(1)&"$$"&rs(2)&"$$"&rs(3)&"$$"&rs(4)&"$$"&rs(5)&"$$"&rs(6)&"$$"&rs(7)&"<br>"

for m=0 to 6
if m=0 then n="星期一"
if m=1 then n="星期二"
if m=2 then n="星期三"
if m=3 then n="星期四"
if m=4 then n="星期五"
if m=5 then n="星期六"
if m=6 then n="星期日"
if rs(n)<>"" then
str=rs("课程名")&"<br>"&rs("教师名")&"<br>"&rs("上课班级")&"<br>"&mid(rs(n),1)&"<br>"
x=int(left(rs(n),1))
y=int(mid(rs(n),3,2))
if x<=1 and y>=1 then kb(m,0)=kb(m,0)&str
if x<=3 and y>=3 then kb(m,1)=kb(m,1)&str
if x<=5 and y>=5 then kb(m,2)=kb(m,2)&str
if x<=7 and y>=7 then kb(m,3)=kb(m,3)&str
if x<=9 and y>=9 then kb(m,4)=kb(m,4)&str

end if
next

rs.movenext
loop
rs.close

Response.Write "<table cellSpacing=1 cellPadding=6 border=0 bgcolor=#000000 align=center width=1240><tbody bgcolor='#ffffff'>"  '输出一个表格和表头
Response.Write "<tr align=center><td> </td><td>周一</td><td>周二</td><td>周三</td><td>周四</td><td>周五</td><td>周六</td><td>周日</td></tr>"
Dim i, j
For i = 0 To 4   '用变量i做第X堂课的循环变量
   Response.Write "<tr><td>" & 2*i + 1 & "-" & 2*i + 2 & "</td>"  '把"第X堂"字样输出到每行第一个单元格中.
   For j = 0 To 6  '用变量j做周几的循环变量
	if kb(j,i)<>"" then
	      Response.Write "<td>" & kb(j,i)  & "</td>"   '依次输出周一到周日的第i堂课
	else
	      Response.Write "<td bgcolor=#aaffff align=center></td>"   '预约课程连接
	end if
   Next
   Response.Write "</tr>"  '一行输出完毕,输出行结束标签
Next
Response.Write "</tbody></table>" '全部输出完比,输出表格结束标签
%>

</td>
</tr>
<%end if%>
</tbody>
</table>
</body>
</html>