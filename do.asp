<%@ language="vbscript"%>
<%response.Expires = 0%>
<%
if request.servervariables("http_referer")="" then
response.write "非法进入"
response.end
end if


if session("admin")="" or session("adminid")=""  then
  response.redirect "login.asp"
else
  adminname=session("admin")
end if%>

<!--#include file="inc/conn.inc"-->
<%
course=trim(Request("course"))'课程内部号
labid=trim(Request("labid"))'实验室地点
i=trim(Request("i"))*2 + 1'节次
m=trim(Request("j"))'星期
cc=0
if m=0 then j="星期一"
if m=1 then j="星期二"
if m=2 then j="星期三"
if m=3 then j="星期四"
if m=4 then j="星期五"
if m=5 then j="星期六"
if m=6 then j="星期日"
if course="" then
 response.write "<script language=JavaScript>" & chr(13) & "alert('未选择课程');" & "history.back()" & "</script>"
  response.End
end if

set rs=Server.CreateObject("ADODB.Recordset")
sqltext="select * from kc where 实验室地点='"&labid&"' and "&j&"<>'' order by "&j
'response.write sqltext
rs.open sqltext,conn,1,1
do while not rs.eof
k=int(left(rs(j),1))
l=int(mid(rs(j),3,2))

	if i>=k and i<=l then
	cc=1
	end if
rs.movenext
loop
rs.close
if cc<>0 then
  response.write "<script language=JavaScript>" & chr(13) & "alert('预定失败，您所选的这节课已被预定。');" & "history.back()" & "</script>"
  response.End
else
sqltext="select * from kc where id="&course
'response.write sqltext
rs.open sqltext,conn,3,3
if rs(j)<>"" then
  if i<int(left(rs(j),1)) then
    ii=int(mid(rs(j),3,2))
  else
    ii=i+1
    i=int(left(rs(j),1))
  end if
else
ii=i+1
end if
'response.write i&"$"&ii
rs(j)=i&"-"&ii&" "&labid
rs("person")=adminname
rs("ttt")=now()
rs.update
Response.Write"<script language=javascript>location.replace('main.asp?labid="&labid&"&course=" & course & "')</script>"
end if
%>

<%
rs.close
conn.close
%>