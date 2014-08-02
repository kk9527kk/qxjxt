<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('网络超时或你还未登录，请重新登陆!');window.location.href='index.htm';}</script>"
response.end
end if

uid=request.Cookies("adminuser")
%>


<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>云阳县国家税务局计算机耗材及配件申请审批表</title>
</head>

<body>

<p align="center"><font size="5" face="楷体_GB2312"><b>云阳县国家税务局计算机耗材及配件申请审批表</b></font></p>
<p align="center">　</p>
<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;            
申请单位:</p><form method="POST" action="save.asp?id=<%=uid%>" name="form1">
<table align="center" border="1" width="68%" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolor="#000000" height="261">
  <tr>
    <td width="25%" bordercolorlight="#000000" height="36">所需耗材名:<select size="1" name="hcmc">
        <option value="硒鼓">硒鼓</option>
        <option value="墨粉">墨粉</option>
        <option value="色带">色带</option>
        <option value="色带架">色带架</option>
        <option value="优盘">优盘</option>
        <option value="移动硬盘">移动硬盘</option>
        <option value="电脑配件">电脑配件</option>
        <option value="打印纸">打印纸</option>
        <option value="其他">其他</option>
                    
        &nbsp;            
        </select>
     
    </td>
    <td width="36%" bordercolorlight="#000000" height="36">耗材型号:<input type="text" name="XH" size="24"></td>
    <td width="24%" bordercolorlight="#000000" height="36">数量:<input type="text" name="SL" size="9" value="1"></td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#000000" height="45" colspan="3">申请事项:<textarea rows="2" name="sqsx" cols="81"> </textarea></td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#000000" height="15" colspan="3">单位负责人签字:            
      同意<input type="radio" value="1" name="DWFZR" checked>，不同意<input type="radio" name="DWFZR" value="0"><textarea rows="2" name="S3" cols="65"></textarea>
      <p>&nbsp;签名：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;            
      时间：</p>
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#000000" height="49" colspan="3">信息中心意见:同意<input type="radio" value="1" name="ZXYJ" checked>，不同意<input type="radio" name="ZXYJ" value="0">&nbsp;&nbsp;            
      <textarea rows="2" name="zxyj" cols="65"></textarea>
      <p>&nbsp;签名：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;            
      &nbsp;&nbsp;&nbsp;&nbsp; 时间：</p>           
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#000000" height="16" colspan="3">局领导意见:同意<input type="radio" value="1" name="JLDYJ" checked>，不同意<input type="radio" name="JLDYJ" value="0">&nbsp;&nbsp;&nbsp;&nbsp;            
      <textarea rows="2" name="jldyj" cols="65"></textarea>
      <p>&nbsp;签名：&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      时间：</p>
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#000000" height="16" colspan="3">
      <p align="center">　<input type="submit" name="B1" value="提交审批">&nbsp; 
      <input type="reset" value="全部重写" name="B2"></p>
      　
      <p>　</p>
    </td>
  </tr>
  <tr>
    <td width="25%" bordercolorlight="#000000" height="16">　</td>
    <td width="36%" bordercolorlight="#000000" height="16">　</td>
    <td width="24%" bordercolorlight="#000000" height="16">　</td>
  </tr>
</table>
 </form>
</body>

</html>