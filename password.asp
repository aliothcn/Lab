<!--#include file="inc/conn.inc"-->
<%if session("admin")="" and session("adminid")=""  then
  response.redirect "login.asp"
else
  adminname=session("admin")
end if%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�û������޸�</title>
<link href="style.css" rel="stylesheet" type="text/css">
<script language=javascript> 
//CharMode����
//����ĳ���ַ���������һ��. 
function CharMode(iN){ 
if (iN>=48 && iN <=57) //���� 
return 1; 
if (iN>=65 && iN <=90) //��д��ĸ 
return 2; 
if (iN>=97 && iN <=122) //Сд 
return 4; 
else 
return 8; //�����ַ� 
} 
//bitTotal���� 
//�������ǰ���뵱��һ���ж�����ģʽ 
function bitTotal(num){ 
modes=0; 
for (i=0;i<4;i++){ 
if (num & 1) modes++; 
num>>>=1; 
} 
return modes; 
} 
//checkStrong���� 
//���������ǿ�ȼ��� 
function checkStrong(sPW){ 
if (sPW.length<=4) 
return 0; //����̫�� 
Modes=0; 
for (i=0;i<sPW.length;i++){ 
//����ÿһ���ַ������ͳ��һ���ж�����ģʽ. 
Modes|=CharMode(sPW.charCodeAt(i)); 
} 
return bitTotal(Modes); 
} 
//pwStrength���� 
//���û��ſ����̻����������ʧȥ����ʱ,���ݲ�ͬ�ļ�����ʾ��ͬ����ɫ 
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
	      alert("���룿");
		  form1.password.focus();
		  return false  }
	  if (form1.new_pass.value==""){
	      alert("�����룿");
		  form1.new_pass.focus();
		  return false  }
	  if (form1.new_pass2.value==""){
	      alert("���ظ�����������");
		  form1.new_pass2.focus();
		  return false  }
	  if (form1.new_pass.value!=form1.new_pass2.value){
	      alert("��һ����������͵ڶ�����������벻һ�£�");
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
response.Write "<script>alert('�������벻һ�£�');history.go(-1);</script>"
Response.End

ElseIf Len(Request.Form("new_pass")) < 6 Or Len(Request.Form("new_pass2")) < 6 Then  
Response.Write "<Script>alert('��������ظ�����������6λ��');history.go(-1);</Script>"
Response.End
else
Set Rs = Server.CreateObject("ADODB.Recordset")
sql="select * from person where person_name='"&username&"' and person_password='"&password&"' and person_name='"&session("admin")&"'"
Rs.Open Sql,conn,3,3
if rs.eof then
response.write "<script>alert('�Բ���ԭ�������');history.go(-1);</script>"
else
   
   sql="update person set person_password='"&new_pass&"' where person_name='"&session("admin")&"'"
   conn.execute(sql)
   Response.Write("<script language=javascript>alert('��������ɹ�������');location.replace('login.asp')</script>")
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
            <td height="68" colspan="3" bgcolor="#FFF0FF"><span class="pt4 style3">�û������޸�</span></td>
          <tr> 
            <td width="93" height="30" align="right" bgcolor="#FFF0FF">�û����ƣ�</td>
            <td width="140" bgcolor="#FFFFFF"> <input type=text name="username" size=19 value="<%=session("admin")%>" readonly></td>
            <td width="186" bgcolor="#FFF0FF">&nbsp;</td>
          <tr> 
            <td width="93" height="30" align="right" bgcolor="#FFF0FF">��½���룺</td>
            <td> <input type=password name="password" size=20> </td>
            <td bgcolor="#FFF0FF">(������)</td>
          </tr>
          <tr> 
            <td width="93" height="30" align="right" bgcolor="#FFF0FF">�� �� �룺</td>
            <td> <input type=password name="new_pass" size=20 onBlur=pwStrength(this.value) onKeyUp=pwStrength(this.value)> </td>
            <td align="left" bgcolor="#FFF0FF"><table width=132 height="23" border=1 align="left" cellpadding=0 cellspacing=0 bordercolor=#CCCCCC bordercolordark=#ffffff> 
<tr align="center" bgcolor="#eeeeee"> 
<td width="33%" valign="middle" bgcolor="#eeeeee" id="strength_L">��</td> 
<td width="33%" valign="middle" id="strength_M">��</td> 
<td width="33%" valign="middle" id="strength_H">ǿ</td> 
</tr>
</table></td>
          </tr>
          <tr> 
            <td width="93" height="30" align="right" bgcolor="#FFF0FF">�ظ����룺</td>
            <td> <input type=password name="new_pass2" size=20> </td>
            <td bgcolor="#FFF0FF">(Ҫ�����������6λ)</td>
          </tr>
          <tr bgcolor="#FFF0FF"> 
            <td height="58" colspan=3 align=center> <input type=submit value="ȷ ��" name="submit"> 
            &nbsp;&nbsp;&nbsp;&nbsp; <input type="reset" name="Submit" value="����"></td>
          </tr>
        </table>
        <br>
        <br>
      </form></td>
  </tr>
</table>
</body>
</html>
