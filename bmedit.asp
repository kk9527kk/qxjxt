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

<%
lid=request.QueryString("id")
sql="select * from dw  where id="&lid
set editrs=server.createobject("adodb.recordset")
editrs.open sql,connstr,3,2
dw=editrs("dw")
fgld=editrs("fgld")
dwbm=editrs("dwbm")
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
<title>�����ع���˰��ֲ���ά��</title>
</head>

<body topmargin="0" background="IMAGES/header_1.gif">

<div align="center"> ��</div>
<div align="center">
  <center>
  <table border="0" width="279">
    <tr>
      <td width="353" height="8"></td>
    </tr>
    <tr>
      <td width="353">
        <div align="left">
          <table border="0" width="268" cellspacing="1" cellpadding="0" bgcolor="#000000" height="96">
          <form action="bmsave1.asp?id=<%=lid%> "  method="post">
            <tr>
              <td width="281" background="images/bg1.gif" height="21">
                <p align="center"><b>������Ϣ�޸�</b></td>
            </tr>
          </center>
            <tr>
              <td width="281" bgcolor="#FFFFFF" height="71">
                <table border="0" width="261" height="26">
                  <tr>
                    <td width="30%" valign="bottom" height="1"> 
                      <p align="right">&nbsp; ��λ��               
                    </td>
                    <td width="70%" height="1"> 
                    <%=dw%> 
                    </td>
                  </tr>
                </table>             
                <table border="0" width="100%" height="60">
                  <tr>
                    <td width="30%" align="right" valign="bottom" height="1">�ֹ��쵼��</td>
                    <td width="70%" height="1"><select size="1" name="fgld">
                        <option value="<%=fgld%>" selected><%=fgld%></option>
<% sql200="select distinct fgld  from dw"
set rs200=server.createobject("adodb.recordset")
rs200.open sql200,connstr,1,1           
do while not rs200.eof
%>
<option value="<%=rs200("fgld")%>"><%=rs200("fgld")%></option>
<%rs200.movenext
loop    
rs200.close
%>  
                      </select></td>
                  </tr>
                  <tr>
                    <td width="30%" align="right" valign="bottom" height="1">��ţ�</td>
                    <td width="70%" height="1"><input type="text" name="dwbm" size="17" style="background-attachment: scroll; background-repeat: repeat; font-size: 9pt; height: 18; width: 101; background-color: #FFFFFF; border: 1 solid #000000; background-position: none 0%" value="<%=dwbm%>"></td>
                  </tr>
                </table>
                <table border="0" width="261" height="68">
                  <tr> 
                    <td width="76" valign="top" height="1">
                      <p align="right">&nbsp;</p>
                    </td>
                    <td width="171" height="1" valign="top"> 
                                              
                        ��
                        <p><input type="submit" value=" �� �� " name="edit" style="background-color: #DDDDDD; background-repeat: repeat; background-attachment: scroll; font-size: 9pt; height: 20; width: 64; background-image: url('images/bg1.gif'); border: 1px groove #000000; background-position: 0% 50%"><input type="reset" value=" ȡ �� " name="reset" style="background-color: #DDDDDD; background-repeat: repeat; background-attachment: scroll; font-size: 9pt; height: 20; width: 71; background-image: url('images/bg1.gif'); border: 1px groove #000000; background-position: 0% 50%"> 
                      </form>                    
                    </td>
                  </tr>
                  <tr> 
                    <td width="253" colspan="2" height="14"> 
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
             
</body>              
              
</html>              



















