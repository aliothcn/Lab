<%if session("admin")="" or session("adminid")<>"ALIOTH@126.COM" then
  response.redirect "login.asp"
response.end
else
  adminname=session("admin")
end if%>
<!--#include file="inc/conn.inc"-->
<html>
<head>
<title>ʵ����ԤԼϵͳ</title>
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
<td bgcolor=#ffffff align=left><%=adminname%>&nbsp;&nbsp;���!&nbsp;&nbsp;��ѡ��γ̣�
<%
if request("labid")="" then
	labid=0
	else
	labid=request("labid")
end if
if request("course")<>"" then'��ѡ���γ�
set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select * from kc where id="&request("course")
set rs=conn.execute(sqlstr)
courseid=rs(0)'�ڲ����
coursename=rs("�γ���")'�γ���
courseno=rs(2)'�γ̺�
coursehours=rs("��ʵ��ѧʱ")'��ʵ��ѧʱ
courseclass=rs("�Ͽΰ༶")'�Ͽΰ༶
coursenumber=rs("����")'����
courseremarks=rs("ʵ���ҵص�")
set rs=nothing
end if

set rs=Server.CreateObject("ADODB.Recordset")
if adminname="admin" then
sqlstr="select * from kc where ʵ���ҵص�<>'' order by ��ʦ��"
else
sqlstr="select * from kc where ��ʦ��='"&adminname&"'and ʵ���ҵص�<>''"
end if
%>
	<select name="list" id="selectcourse" onChange="load(this.form)">
<option selected value="">��ѡ��</option>
<%
Rs.Open Sqlstr,conn,2,3  
Do While Not Rs.EOF 
Response.Write "<option value='main.asp?labid=" & rs("ʵ���ҵص�") &"&course="& rs("id") &"'>" & rs("��ʦ��") & "|" & Rs("�γ���") & "</option>"
Rs.MoveNext
Loop 
rs.close
set rs=nothing%>
</select>
</td></form>
</tr>
<%if request("course")<>"" then%>
<tr>
<td bgcolor=#ffffff align=left>��ǰ�γ�:<%=coursename%>&nbsp;|&nbsp;���:<%=courseno%>&nbsp;|&nbsp;ѧʱ:<%=coursehours%>&nbsp;|&nbsp;�༶:<%=courseclass%>&nbsp;|&nbsp;����:<%=coursenumber%>&nbsp;|&nbsp;��ע��<%=courseremarks%>&nbsp;&nbsp;&nbsp;<a href=del.asp?labid=<%=request("labid")%>&course=<%=courseid%>>�����ѡ</a>
</td>
</tr>

<tr>
<td bgcolor=#ffffff>
<%
sqlstr="select * from kc where ("
aaa=split(courseclass)
j=ubound(aaa)
sqlstr=sqlstr&"�Ͽΰ༶ like '%" & left(aaa(0),4) &"%'" 
if j>0 then
for i=1 to j+1
sqlstr=sqlstr&" or �Ͽΰ༶ like '%" & left(aaa(i),4) &"%'"  
next
end if
sqlstr=sqlstr&" or ʵ���ҵص� like '%"&courseremarks&"%')"
'response.write sqlstr
Dim kb(6,4) '��һά��ʾ�ܼ�,�ڶ�ά��ʾ�ڼ��ÿ�
	set rs=conn.execute(sqlstr)
	do while not Rs.eof
'response.write rs(0)&"$$"&rs(1)&"$$"&rs(2)&"$$"&rs(3)&"$$"&rs(4)&"$$"&rs(5)&"$$"&rs(6)&"$$"&rs(7)&"<br>"

for m=0 to 6
if m=0 then n="����һ"
if m=1 then n="���ڶ�"
if m=2 then n="������"
if m=3 then n="������"
if m=4 then n="������"
if m=5 then n="������"
if m=6 then n="������"
if rs(n)<>"" then
str=rs("�γ���")&"<br>"&rs("��ʦ��")&"<br>"&rs("�Ͽΰ༶")&"<br>"&mid(rs(n),1)&"<br>"
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

Response.Write "<table cellSpacing=1 cellPadding=6 border=0 bgcolor=#000000 align=center width=1240><tbody bgcolor='#ffffff'>"  '���һ�����ͱ�ͷ
Response.Write "<tr align=center><td> </td><td>��һ</td><td>�ܶ�</td><td>����</td><td>����</td><td>����</td><td>����</td><td>����</td></tr>"
Dim i, j
For i = 0 To 4   '�ñ���i����X�ÿε�ѭ������
   Response.Write "<tr><td>" & 2*i + 1 & "-" & 2*i + 2 & "</td>"  '��"��X��"���������ÿ�е�һ����Ԫ����.
   For j = 0 To 6  '�ñ���j���ܼ���ѭ������
	if kb(j,i)<>"" then
	      Response.Write "<td>" & kb(j,i)  & "</td>"   '���������һ�����յĵ�i�ÿ�
	else
	      Response.Write "<td bgcolor=#aaffff align=center>��<a href=yuyue.asp?j="&j&"&i="&i&"&course="&request("course")&"&labid="&labid&">ԤԼ</a></td>"   'ԤԼ�γ�����
	end if
   Next
   Response.Write "</tr>"  'һ��������,����н�����ǩ
Next
Response.Write "</tbody></table>" 'ȫ��������,�����������ǩ
%>

</td>
</tr>
<%end if%>
</tbody>
</table>
</body>
</html>