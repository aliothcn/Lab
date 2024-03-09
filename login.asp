<%@ language=vbscript codepage="936"  %>
<%
if session("admin")<>"" then response.redirect("main.asp")
 dim s
randomize timer
s=Int((8999)*Rnd +1000)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>::实验室预约系统::用户登陆</title>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>    
BODY {
	MARGIN: 0px;
    }
.STYLE5 {color: #FFFFFF}
.style8 {
	font-size: 14px;
	font-family: "Times New Roman", Times, serif;
	color: #FFFFff;
}
.zhweng {
	font-size: 14px;
	line-height: 120%;
	color: #000000;
	text-decoration: none;
	text-align: left;
	vertical-align: bottom;
	height: auto;
	cursor: w-resize;
	font-family: "宋体";
	text-indent: 2pc;
}
.style22 {font-size: 14px; font-family: "Times New Roman", Times, serif; color: #000000; }
body,td,th {
	font-family: Times New Roman, Times, serif;
	font-size: 12px;
	color: #000000;
}
a {
	font-family: Times New Roman, Times, serif;
	font-size: 12px;
}
a:visited {8
	color: #0000FF;
}
a:active {
	color: #FF0000;
}
.style24 {
	font-size: 18px;
	color: #EEEEFF;
	font-weight: bold;
}
</STYLE>
</style>
<script language=javascript>
  function xxg()
  {
      if (document.form1.username.value==""){
	      alert("您的姓名？")
		  document.form1.username.focus();
		  return false
		    }
	  if (document.form1.password.value==""){
	      alert("您的密码？");
		  document.form1.password.focus();
		  return false
		  }
		  return true
		}
</script>

</head>
<body>
<%
session("admin")=""
session("adminid")=""%>
      <TABLE width="100%" height="600" border=0 align="center" cellPadding=0 cellSpacing=0 bordercolor="0">
        <TBODY>
        <TR>
          <TD valign="middle" id=td0><%if request.querystring("action")="" then%>	<form name="form1" method="post" action="loginchk.asp" onSubmit="return xxg()">
            <TABLE border=1 align="center" cellPadding=1 cellSpacing=1 bordercolor="#0000FF" bordercolordark=#ffffff>
              <TBODY>
              <TR>
                <TD align=center vAlign=middle><TABLE width="326" border=0 align="center" 
                  cellPadding=0 cellSpacing=0 id=table1 
                  valign="middle">
                      <TBODY>
                      <TR>
                        <TD align=middle colSpan=4></TD></TR>
                      <TR align="center" bgcolor="#336699">
                        <TD height="45" colspan="4" class="style24">用户登陆</TD>
                        </TR>
                      <TR>
                        <TD width="47" height="36" align=right class=style8> </TD>
                        <TD width="76" align=right class=style22>帐&nbsp;&nbsp;&nbsp;号：</TD>
                        <TD colSpan=2>
                        <input name="username"  type="text"  id="UserName5" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; WIDTH: 110px; BORDER-BOTTOM: 1px solid" 
 maxlength="20" class="login"></TD></TR>
                      <TR>
                        <TD height="36" align=right class=style8>&nbsp;</TD>
                        <TD align=right class=style22>密&nbsp;&nbsp;&nbsp;码：</TD>
                        <TD colSpan=2><input name="password"  type="password" id="Password3" maxlength="20" class="login" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; WIDTH: 110px; BORDER-BOTTOM: 1px solid">
                        </TD></TR>
                      <TR>
                        <TD height="36" align=right class=style8>&nbsp;</TD>
                        <TD align=right class=style22>验证码：</TD>
                        <TD width=76>
                        <input name="s" type="text" class="login" id="DreamhomeSys_CheckCode" size=12 maxlength="20" style="BORDER-RIGHT: 1px solid; BORDER-TOP: 1px solid; BORDER-LEFT: 1px solid; WIDTH: 70px; BORDER-BOTTOM: 1px solid"></TD>
                        <TD width=127 class=style22><font size=3><b><%=s%></b></font>
                <input maxlength=20 name="s2" size=18 type=hidden value="<%=s%>"></TD></TR>
                      <TR align=center vAlign=center>
                        <TD height=45 colSpan=4>
                          <input type="submit" name="Submit3" value=" 登陆 ">
&nbsp;&nbsp;&nbsp;&nbsp;                      
  <input type="reset" name="Submit2" value=" 重置 "></TD>
                      </TR></TBODY>
            </TABLE></TD></TR></TBODY></TABLE></FORM><%end if%></TD>
</TR></TBODY></TABLE>
</BODY></HTML>