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
<td bgcolor=#ffffff align=left><%=adminname%>&nbsp;&nbsp;���!&nbsp;&nbsp;��ѡ���ʦ��
<%
set rs=Server.CreateObject("ADODB.Recordset")
sqlstr="select distinct ��ʦ�� from kc"
%>
	<select name="list" id="selectcourse" onChange="load(this.form)">
<option selected value="">��ѡ��</option>
<%
Rs.Open Sqlstr,conn,2,3  
Do While Not Rs.EOF 
Response.Write "<option value='teacher.asp?teacher="& rs("��ʦ��") &"'>" & rs("��ʦ��") & "</option>"
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
sqlstr="select * from kc where ��ʦ��='"&request("teacher")&"'"
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
	      Response.Write "<td bgcolor=#aaffff align=center></td>"   'ԤԼ�γ�����
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