<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='index.htm';}</script>"
response.end
end if
%>
<html>
<head>
<title>管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312"><meta http-equiv="keywords" content="">
<LINK href="teacher/css.css" type=text/css rel=stylesheet>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>

<body background="teacher/img/bg.gif" topmargin="0" leftmargin="0" onLoad="MM_preloadImages('img/InfoInput_B.JPG','img/InfoModifyDel_B.JPG','img/InfoQuery_B.JPG','img/InfoPrint_B.JPG','img/ClassManage_B.jpg')">
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#99CC00">
    <tr bgcolor="#99CC00">
    <td height="50">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2">　</td>
  </tr>
  <tr>
    <td width="93" bgcolor="#CEE785" height="490" valign="top">　</td>
    <td width="685" valign="top"><table width="100%"  border="0" cellpadding="0" cellspacing="0" bgcolor="#99CC00">
      <tr>
        <td width="12%" height="30">&nbsp;</td>
        <td width="22%">&nbsp;</td>
        <td width="6%">&nbsp;</td>
        <td width="22%">&nbsp;</td>
        <td width="6%">&nbsp;</td>
        <td width="22%">&nbsp;</td>
        <td width="10%">&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><a href="teacher/admin_writeto.asp" target="_blank" onMouseOver="MM_swapImage('Image3','','img/InfoInput_B.JPG',1)" onMouseOut="MM_swapImgRestore()"><img src="teacher/img/InfoInput_A.JPG" alt="录入信息" name="Image3" width="150" height="113" border="0"></a></td>
        <td>&nbsp;</td>
        <td><a href="teacher/admin_modify.asp" target="_blank" onMouseOver="MM_swapImage('Image4','','img/InfoModifyDel_B.JPG',1)" onMouseOut="MM_swapImgRestore()"><img src="teacher/img/InfoModifyDel_A.JPG" alt="修改信息" name="Image4" width="150" height="113" border="0"></a></td>
        <td>&nbsp;</td>
        <td><a href="teacher/admin_search.asp" target="_blank" onMouseOver="MM_swapImage('Image5','','img/InfoQuery_B.JPG',1)" onMouseOut="MM_swapImgRestore()"><img src="teacher/img/InfoQuery_A.JPG" alt="查询信息" name="Image5" width="150" height="113" border="0"></a></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><a href="teacher/admin_print.asp" target="_blank" onMouseOver="MM_swapImage('Image6','','img/InfoPrint_B.JPG',1)" onMouseOut="MM_swapImgRestore()"><img src="teacher/img/InfoPrint_A.JPG" alt="打印信息" name="Image6" width="150" height="113" border="0"></a></td>
        <td>&nbsp;</td>
        <td><a href="teacher/admin_admin.asp" target="_blank" onMouseOver="MM_swapImage('Image7','','img/ClassManage_B.jpg',1)" onMouseOut="MM_swapImgRestore()"><img src="teacher/img/ClassManage_A.jpg" alt="帐号管理" name="Image7" width="150" height="113" border="0"></a></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
