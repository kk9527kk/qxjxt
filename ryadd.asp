<!--#include file="conn.asp"-->
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='login.asp';}</script>"
response.end
end if
uid=request.Cookies("adminuser")
lid=request.QueryString("id")
'response.write(lid)
if  request.Cookies("adminuser")="" then
   response.write "��Դδ֪�����ݶ�ʧ!"
   response.end
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
<title>�����ع���˰�����Ա¼��</title>
</head>

<body topmargin="0" background="images/header_1.gif">

<div align="center"> ��</div>
<div align="center">
  <center>
  <table border="0" width="306">
    <tr>
      <td width="66%" height="8"></td>
    </tr>
    <tr>
      <td width="66%">
        <div align="left">
          <table border="0" width="300" cellspacing="1" cellpadding="0" bgcolor="#000000">
          <form action="rysave.asp" method="post">
            <tr>
              <td width="100%" background="images/bg1.gif" height="23">
                <p align="center"><b>�û���Ϣ���</b></td>
            </tr>
          </center>
            <tr>
              <td width="100%" bgcolor="#FFFFFF">
                <table border="0" width="100%" height="147">
                  <tr> 
                    <td width="30%" valign="bottom" height="25"> 
                      <p align="center">&nbsp;&nbsp; ���ڵ�λ��                
                    </td>
                    <td width="70%" height="25"> 
						<select size="1" name="dw">
                         <option value>��ѡ��λ</option>
<% sql200="select *  from dw order by dwbm"
set rs200=server.createobject("adodb.recordset")
rs200.open sql200,connstr,1,1           
do while not rs200.eof
%>
<option value="<%=rs200("dw")%>"><%=rs200("dw")%></option>
<%rs200.movenext
loop    
rs200.close
%>  
                        </select>
                    </td>
                  </tr>
                  <tr>
                    <td width="30%" align="right" valign="bottom" height="16">
                    ������</td>
                    <td width="70%" height="16"> 
                      <input type="text" name="zg" size="17" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">
                    </td>
                  </tr>
                  <tr> 
                    <td width="30%" align="right" valign="bottom" height="16">
                    �˺ţ�</td>
                    <td width="70%" height="16"> 
                      <input type="text" name="ID" size="17" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">
                    </td>
                  </tr>
                  <tr> 
                    <td width="30%" align="right" valign="bottom" height="16">
                    ְ��
                    </td>
                    <td width="70%" height="16"> 
                    
                    <p><select size="1" name="zw">
                      <option value="1">һ��ְ��</option>
                      <option value="2">��������ְ</option>
                      <option value="3">������</option>
                      <option value="4">�ֹܾ��쵼</option>
                      <option value="5">�ֳ�</option>
                      <option value="6">�˽̿ƿƳ�</option>
                      <option value="7">�˽̿ƹ���Ա</option>
                    </select>
                    </td>
                  </tr>
                  <tr> 
                    <td width="30%" valign="top" height="25"> 
                      <p align="right">��ţ�</td>
                    <td width="70%" valign="top" height="25"> 
                      <input type="text" name="xh" size="17" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%">
                    </td>
                  </tr>
                  <tr> 
                    <td width="30%" valign="top" height="55"> &nbsp;</td>
                    <td width="70%" valign="top" height="55"> 
                                                       ��
                                                       <p> 
                                                       <input type="submit" value=" �� �� " name="add" style="background-color: #DDDDDD; background-repeat: repeat; background-attachment: scroll; font-size: 9pt; height: 20; width: 71; background-image: url('images/bg1.gif'); border: 1px groove #000000; background-position: 0% 50%"><input type="reset" value=" ȡ �� " name="reset" style="background-color: #DDDDDD; background-repeat: repeat; background-attachment: scroll; font-size: 9pt; height: 20; width: 71; background-image: url('images/bg1.gif'); border: 1px groove #000000; background-position: 0% 50%"> 
						</form>                        
                                                       ��</p>
                    </td>
                  </tr>
                  <tr> 
                    <td width="100%" colspan="2" height="14"> 
                      <p align="center"> 
                        &nbsp;&nbsp;&nbsp;&nbsp;
                      </p>
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
              
<div align="center">              
  <center>              
  <table border="0" width="441" cellspacing="0" cellpadding="0">              
    <tr>              
      <td height="8">
        <p align="center">˵��:��ʼ�û�������Ϊ1,�ʺŲ����غ�</td>              
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
  </center>               
</div>               
               
</body>               
               
</html>               




