<!--#include file="inc/conn.inc"-->
<%
session("admin")=""
session("adminid")=""  
       s=Trim(Request.Form("s"))
	   s2=Request.Form("s2")
	   If s2<>s Then
	   	   Response.Write("<script language=javascript>alert('��������ȷ����֤�룡');history.back()</script>")
		   Response.End
	   end if
	   username=replace(trim(request.form("username")),"'","''")
       password=replace(trim(request.form("password")),"'","''")
	   'password=MD5(password)

	   if instr(username,"%") or instr(username,"#") or instr(username,"?") or instr(username,"|") then
	   response.write "<script language=javascript>alert('�����������зǷ��ַ���');history.back()</script>"
	   response.end 
	   end if                               '====================����������Ƿ��зǷ��ַ�
	   if instr(password,"%") or instr(password,"#") or instr(password,"?") or instr(password,"|") then
	   response.write "<script language=javascript>alert('�������뺬�зǷ��ַ���');history.back()</script>"
	   response.end 
	   end if                              '====================����������Ƿ��зǷ��ַ�
	   sql="select * from person where person_name='"&username&"' and person_password='"&password&"'"
	   set rs=conn.execute(sql)
	   if rs.eof then
	   	   Response.Write("<script language=javascript>alert('�û������������');history.back()</script>")
		   Response.End
       else'if username<>"admin" then
set rs1=server.createobject("adodb.recordset")
sqltext1="select * from log"
rs1.open sqltext1,conn,3,3
rs1.addnew
rs1("time")=now()
rs1("ip")=Request.ServerVariables("REMOTE_ADDR")
rs1("name")=username
rs1("kind")="��½"
rs1.update
rs1.close


	   Session("admin")=username
           session("adminid")="ALIOTH@126.COM"
	   Response.Redirect("main.asp")         '=============�����֤�ɹ����������Աҳ��
           'Response.Write"<script language=javascript>location.replace('main.asp')</script>"
response.end

       end if
	   conn.close
	   set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
</head>

<body>

</body>
</html>