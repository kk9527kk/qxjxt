<!--#include file="conn.asp"-->
<%
username=request.form("username")
oldpass=request.form("oldpass")
newpass=request.form("newpass")
confirmpass=request.form("confirmpass")
if username<>"" and newpass=confirmpass then
sql="select * from zg where id='"&username&"'"
set rs=server.createobject("adodb.recordset")
rs.open sql,connstr,3,2
if not rs.eof then
	'�ж��Ƿ��޸ı�������,sws+
	uid=request.Cookies("adminuser")
	if username<>uid then
		response.write("����ֻ���޸ı��˵����룡")
	else
	
	if oldpass=rs("mm") then
	rs("mm")=newpass
	rs.update
	response.redirect "ok.asp"
end if
	end if
	end if
end if
%>
<html>
<head>
<STYLE type=text/css>A:link {
	COLOR: #000000; FONT-FAMILY: ����; TEXT-DECORATION: none
}
A:visited {
	COLOR: #000000; FONT-FAMILY: ����; TEXT-DECORATION: none
}
A:active {
	FONT-FAMILY: ����; TEXT-DECORATION: underline
}
A:hover {
	COLOR: #84bd6b; TEXT-DECORATION: underline overline
}
BODY {
	COLOR: #000000; FONT-FAMILY: ����; FONT-SIZE: 9pt
}
TABLE {
	COLOR: #000000; FONT-FAMILY: ����; FONT-SIZE: 9pt
}
.f24 {
	COLOR: #ff0000; FONT-FAMILY: ����; FONT-SIZE: 9pt; TEXT-DECORATION: underline overline
}
.l17 {
	LINE-HEIGHT: 170%
}
.f18 {
	FONT-SIZE: 18px
}
</STYLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�޸Ŀ���</title>
</head>

<body topmargin="0" background="images/header_1.gif">

<div align="center"> ��</div>
<div align="left">
  <table   align="center"  border="0" width="441">
    <tr>
      <td width="89%" height="8" align="center"></td>
    </tr>
    <tr>
      <td width="89%" align="center">
        <div align="center">
          <center>
          <table border="0"  align="center"  width="300" cellspacing="1" cellpadding="0" bgcolor="#000000">
          <form action="" method="post">
            <tr>
              <td width="100%" background="images/bg1.gif" height="23">
                <p align="center"><b>����Ա�����޸�</b></td>
            </tr>
          </center>
            <tr>
              <td width="100%" bgcolor="#FFFFFF">
                <table border="0" width="100%">
                  <tr>
                    <td width="43%" valign="bottom">
                      <p align="right">����Ա�ʺţ�</td>
          <td width="57%"><input type="text" name="username" size="17" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%"></td>
  </tr>
                  <tr>
                    <td width="43%" align="right" valign="bottom">
                      ԭʼ���룺</td>
          <td width="57%"><input type="text" name="oldpass" size="27" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%"></td>
  </tr>
                  <tr>
                    <td width="43%" align="right" valign="bottom">
                      �������룺</td>
          <td width="57%"><input type="text" name="newpass" size="17" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%"></td>
  </tr>
                  <tr>
                    <td width="43%" align="right" valign="bottom">
                      ȷ�����룺</td>
          <td width="57%"><input type="text" name="confirmpass" size="17" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%"></td>
  </tr>
  <tr>
    <td width="100%" colspan="2">
      <p align="center"><input type="submit" value=" �� ��  " name="edit" style="background-color: #DDDDDD; background-repeat: repeat; background-attachment: scroll; font-size: 9pt; height: 20; width: 71; background-image: url('images/bg1.gif'); border: 1px groove #000000; background-position: 0% 50%">&nbsp;&nbsp;   
      <input type="reset" value=" ȡ �� " name="reset" style="background-color: #DDDDDD; background-repeat: repeat; background-attachment: scroll; font-size: 9pt; height: 20; width: 71; background-image: url('images/bg1.gif'); border: 1px groove #000000; background-position: 0% 50%"></p>              
    </td>              
  </tr></form>              
                    </table>              
                  </td>              
            </tr>              
          </table>              
        </div>              
      </td>              
    </tr>              
  </table>              
</div>              
              
<div align="left">              
  <table align="center" border="0" width="439" cellspacing="0" cellpadding="0">              
    <tr>              
      <td height="8"></td>              
    </tr>              
    <tr>              
      <td>         
        <p align="center"></td>              
    </tr>              
    <tr>              
      <td height="8"></td>              
    </tr>              
    <tr>              
        <td height="35" valign="bottom"> 
          <p align="center"><br>
          
        </td>                    
    </tr>               
  </table>               
</div>               
               
</body>               
               
</html>               
