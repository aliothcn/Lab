<%  adminname=session("admin")
%>
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
<td bgcolor=#ffffff align=left><%=adminname%>&nbsp;&nbsp;你好!&nbsp;&nbsp;请选择班级：
<%
if request("show")="" then
	show=0
	else
	show=request("show")
end if
if request("classname")<>"" then
	classname=request("classname")
end if
aa=year(now())-4
bb=split("测控 自动化 电子 通信 软件 软件双 计算机 数字媒体 机械")
%>
	<select name="list" id="selectclass" onChange="load(this.form)">
<option selected value="">请选择</option>
<%for i=0 to ubound(bb)-1%>
<option value='class.asp?classname=<%=right(aa,2)&bb(i)%>'><%=right(aa,2)&bb(i)%></option>
<%next
aa=aa+1
for i=0 to ubound(bb)-1%>
<option value='class.asp?classname=<%=right(aa,2)&bb(i)%>'><%=right(aa,2)&bb(i)%></option>
<%next
aa=aa+1
for i=0 to ubound(bb)-1%>
<option value='class.asp?classname=<%=right(aa,2)&bb(i)%>'><%=right(aa,2)&bb(i)%></option>
<%next
aa=aa+1
for i=0 to ubound(bb)-1%>
<option value='class.asp?classname=<%=right(aa,2)&bb(i)%>'><%=right(aa,2)&bb(i)%></option>
<%next
aa=aa+1
for i=0 to ubound(bb)-1%>
<option value='class.asp?classname=<%=right(aa,2)&bb(i)%>'><%=right(aa,2)&bb(i)%></option>
<%next
%>
</select>
</td></form>
</tr>

<tr>
<td bgcolor=#ffffff>
<a href=class.asp?classname=<%=request("classname")%>&show=<%=show+14%>>上两周</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href=class.asp?classname=<%=request("classname")%>&show=0>本周</a>&nbsp;&nbsp;|&nbsp;&nbsp;<a href=class.asp?classname=<%=request("classname")%>&show=<%=show-14%>>下两周</a>
<%	Set rs = Server.CreateObject("ADODB.Recordset")
	sqlstr="select * from Curriculum order by Curriculum_id ASC"
	set rs=conn.execute(sqlstr)
%>
  <table cellspacing=0 bordercolordark=#ffffff cellpadding=0 width="1220" align=center border=1>
    <tr>
    <td width=80 height=30 bgcolor=#99ccff align=center>日 期</td>
    <td width="65" bgcolor="#99ccff" align="center">星 期</td>
	<%
	do while not rs.eof%>
	<td bgcolor="#99ccff" align="center"><%=rs("Curriculum_name")%></td>
	<%rs.movenext
    loop%>
    </tr>
<%
	week=weekday(date-show)
	for a=0 to 13
%>
<%if weekend and(weekday(dateadd("d",date-show,a-week))=1 or weekday(dateadd("d",date-show,a-week))=7) then
%>
<tr>
<td width="80" height="38" align="center" bgcolor="#99ccff">
	<%if dateadd("d",date-show,a-week)=date then
	response.write "<font color=#550055><b>"&dateadd("d",date-show,a-week)&"</b></font>"
	'elseif dateadd("d",date-show,a-week)<date then
	'response.write "<font color=red>"&dateadd("d",date-show,a-week)&"</font>"
	else'if dateadd("d",date-show,a-week)>date then
	response.write dateadd("d",date-show,a-week)
	end if%></td>
    <td width="65" align="center" bgcolor="#99ccff">
    <%if a mod 7=0 or a mod 7=1 then
	response.write "<font color=red>"&weekdayname(weekday(dateadd("d",date-show,a-week)))&"</font>"
	elseif a = 1 then
	response.write "<font color=red>"&weekdayname(weekday(dateadd("d",date-show,a-week)))&"</font>"
	else
	response.write ""&weekdayname(weekday(dateadd("d",date-show,a-week)))&""
	end if%>
    </td>
</tr>
<%else%>
  <tr>
    <td width="80" height="38" align="center" bgcolor="#99ccff">
	<%if dateadd("d",date-show,a-week)=date then
	response.write "<font color=#550055><b>"&dateadd("d",date-show,a-week)&"<b></font>"
	'elseif dateadd("d",date-show,a-week)<date then
	'response.write "<font color=red>"&dateadd("d",date-show,a-week)&"</font>"
	else'if dateadd("d",date-show,a-week)>date then
	response.write dateadd("d",date-show,a-week)
	end if%>
	</td>
    <td width="65" align="center" bgcolor="#99ccff">
    <%if (a mod 7)= 0 or a mod 7 =1 then
	response.write "<font color=red>"&weekdayname(weekday(dateadd("d",date-show,a-week)))&"</font>"
	elseif a = 1 then
	response.write "<font color=red>"&weekdayname(weekday(dateadd("d",date-show,a-week)))&"</font>"
	else
	response.write ""&weekdayname(weekday(dateadd("d",date-show,a-week)))&""
	end if%>
    </td>
	<%set rs=server.createobject("adodb.recordset")
	rs.open "select count(*) from Curriculum",conn,1,1
	num=rs(0)
	rs.close
	for b=1 to num%>
	<%if request("classname")<>"" then

        sqlstr="select * from course left join pk on course.course_id=pk.pk_course where (course.course_class like '%"&classname&"%') and pk_day=#"&dateadd("d",date-show,a-week)&"# and pk_time="&b

	set rs=conn.execute(sqlstr)
        if Rs.eof or Rs.bof then
	  if dateadd("d",date-show,a-week)<date then
	  response.write "<td align=center><font color=red><font size=3>过期</font></font></td>"
	  else
          response.write "<td align=center bgcolor=#EEffEE></td>"
 	  end if
	end if
set rs1=Server.CreateObject("ADODB.Recordset")
sqlstr1="select * from lab where lab_id="&rs("course_lab")
Rs1.Open Sqlstr1,conn,2,3 

end if

    response.write "<td align=center bgcolor=#ccccff><font size=2>"&rs("course_name") & "&nbsp;" &rs("course_teacher") &"<br>"& rs("course_class")& "&nbsp;" &rs("course_number")&"人<br>" &rs1("lab_name")& "</font></a>"
response.write "</td>"
next%>
  </tr>
  <%
end if
next
  conn.close
	   set conn=nothing%>
</table>
<%rs.close
    set rs=nothing%>

  </table>
</td>
</tr>
</tbody>
</table>
</body>
</html>