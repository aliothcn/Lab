<!--#include file="inc/conn.inc"-->
<%if session("admin")="" and session("adminid")=""  then
  response.redirect "login.asp"
else
  adminname=session("admin")
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户密码修改</title>
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript> 
//CharMode函数
//测试某个字符是属于哪一类. 
function CharMode(iN){ 
if (iN>=48 && iN <=57) //数字 
return 1; 
if (iN>=65 && iN <=90) //大写字母 
return 2; 
if (iN>=97 && iN <=122) //小写 
return 4; 
else 
return 8; //特殊字符 
} 
//bitTotal函数 
//计算出当前密码当中一共有多少种模式 
function bitTotal(num){ 
modes=0; 
for (i=0;i<4;i++){ 
if (num & 1) modes++; 
num>>>=1; 
} 
return modes; 
} 
//checkStrong函数 
//返回密码的强度级别 
function checkStrong(sPW){ 
if (sPW.length<=4) 
return 0; //密码太短 
Modes=0; 
for (i=0;i<sPW.length;i++){ 
//测试每一个字符的类别并统计一共有多少种模式. 
Modes|=CharMode(sPW.charCodeAt(i)); 
} 
return bitTotal(Modes); 
} 
//pwStrength函数 
//当用户放开键盘或密码输入框失去焦点时,根据不同的级别显示不同的颜色 
function pwStrength(pwd){ 
O_color="#eeeeee"; 
L_color="#FF0000"; 
M_color="#FF9900"; 
H_color="#33CC00"; 
if (pwd==null||pwd==''){ 
Lcolor=Mcolor=Hcolor=O_color; 
} 
else{ 
S_level=checkStrong(pwd); 
switch(S_level) { 
case 0: 
Lcolor=Mcolor=Hcolor=O_color; 
case 1: 
Lcolor=L_color; 
Mcolor=Hcolor=O_color; 
break; 
case 2: 
Lcolor=Mcolor=M_color; 
Hcolor=O_color; 
break; 
default: 
Lcolor=Mcolor=Hcolor=H_color; 
} 
}
document.getElementById("strength_L").style.background=Lcolor; 
document.getElementById("strength_M").style.background=Mcolor; 
document.getElementById("strength_H").style.background=Hcolor; 
return; 
} 
</script> 
<script language=javascript>
   function check_edit(){
	 if (form1.password.value==""){
	      alert("密码？");
		  form1.password.focus();
		  return false  }
	  if (form1.new_pass.value==""){
	      alert("新密码？");
		  form1.new_pass.focus();
		  return false  }
	  if (form1.new_pass2.value==""){
	      alert("请重复输入新密码");
		  form1.new_pass2.focus();
		  return false  }
	  if (form1.new_pass.value!=form1.new_pass2.value){
	      alert("第一次输入密码和第二次输入的密码不一致！");
		  form1.new_pass2.focus();

		  return false  }
		  return true  }
</script>
<style type="text/css">
<!--
.STYLE1 {
	font-size: 16px;
	color: #FF0000;
	font-weight: bold;
}
.style3 {
	font-size: 24px;
	font-weight: bold;
	color: #0033FF;
}
-->
</style>
</head>
<body>
<%
IF Request.QueryString("action")="edit" then
Call UserEdit()

End IF

Sub UserEdit()
dim username
dim password
dim new_pass

   username    =Replace(Trim(Request.Form("username")),"'","''")
   password    =Replace(Trim(Request.Form("password")),"'","''")
   
   new_pass=Replace(Trim(Request.Form("new_pass")),"'","''")   
   new_pass2=Replace(Trim(Request.Form("new_pass2")),"'","''")
   
if new_pass<>new_pass2 then
response.Write "<script>alert('两次密码不一致！');history.go(-1);</script>"
Response.End

ElseIf Len(Request.Form("new_pass")) < 6 Or Len(Request.Form("new_pass2")) < 6 Then  
Response.Write "<Script>alert('密码项或重复密码项少于6位！');history.go(-1);</Script>"
Response.End
else
Set Rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where person_name='"&username&"' and person_password='"&password&"' and person_name='"&session("admin")&"'"
Rs.Open Sql,conn,3,3
if rs.eof then
response.write "<script>alert('对不起！原密码错误！');history.go(-1);</script>"
else
   
   sql="update person set person_password='"&new_pass&"' where person_name='"&session("admin")&"'"
   conn.execute(sql)
   Response.Write("<script language=javascript>alert('更改密码成功！！！');location.replace('login.asp')</script>")
   connclose()
End If
 End If 
End Sub%>
<table width=467 height="314" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="10" align="center"><br>
    </td>
    <td width="750" align="center" valign="top"> 
      <form name="form1" method="post" action="password.asp?action=edit" onSubmit="return check_edit()" autocomplete="off">
        <br>
        <br>
                <br>
        <br>
        <table width="445" bordercolor=#CCCCCC cellspacing=0 bordercolordark=#ffffff cellpadding=0 
align=center border=1" class="sft">

          <tr align="center">
            <td height="68" colspan="3" bgcolor="#FFF0FF"><span class="pt4 style3">用户密码修改</span></td>
          <tr> 
            <td width="93" height="30" align="right" bgcolor="#FFF0FF">用户名称：</td>
            <td width="140" bgcolor="#FFFFFF"> <input type=text name="username" size=19 value="<%=session("admin")%>" readonly></td>
            <td width="186" bgcolor="#FFF0FF">&nbsp;</td>
          <tr> 
            <td width="93" height="30" align="right" bgcolor="#FFF0FF">登陆密码：</td>
            <td> <input type=password name="password" size=20> </td>
            <td bgcolor="#FFF0FF">(旧密码)</td>
          </tr>
          <tr> 
            <td width="93" height="30" align="right" bgcolor="#FFF0FF">新 密 码：</td>
            <td> <input type=password name="new_pass" size=20 onBlur=pwStrength(this.value) onKeyUp=pwStrength(this.value)> </td>
            <td align="left" bgcolor="#FFF0FF"><table width=132 height="23" border=1 align="left" cellpadding=0 cellspacing=0 bordercolor=#CCCCCC bordercolordark=#ffffff> 
<tr align="center" bgcolor="#eeeeee"> 
<td width="33%" valign="middle" bgcolor="#eeeeee" id="strength_L">弱</td> 
<td width="33%" valign="middle" id="strength_M">中</td> 
<td width="33%" valign="middle" id="strength_H">强</td> 
</tr>
</table></td>
          </tr>
          <tr> 
            <td width="93" height="30" align="right" bgcolor="#FFF0FF">重复密码：</td>
            <td> <input type=password name="new_pass2" size=20> </td>
            <td bgcolor="#FFF0FF">(要求新密码大于6位)</td>
          </tr>
          <tr bgcolor="#FFF0FF"> 
            <td height="58" colspan=3 align=center> <input type=submit value="确 定" name="submit"> 
            &nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" name="Submit" value="重置"></td>
          </tr>
        </table>
        <br>
        <br>
      </form></td>
  </tr>
</table>
</body>
</html>
