<!--#include file="inc/conn.inc"-->
<% 
set rs1=server.createobject("adodb.recordset")
sqltext1="select * from log"
rs1.open sqltext1,conn,3,3
rs1.addnew
rs1("time")=now()
rs1("ip")=Request.ServerVariables("REMOTE_ADDR")
rs1("name")=session("admin")
rs1("kind")="离开"
rs1.update
rs1.close
   session("admin")=""
   Session.Abandon()
   Response.Redirect "login.asp"
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FFFFFF" text="#000000">

</body>
</html>