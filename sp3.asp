<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
if request.Cookies("flag")<>"loginok" then
response.write "<script language=JavaScript>{window.alert('���糬ʱ���㻹δ��¼�������µ�½!');window.location.href='index.htm';}</script>"
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
<title>�����ع���˰��ּ�����Ĳļ��������������</title>
</head>

<body>

<p align="center"><font size="5" face="����_GB2312"><b>�����ع���˰��ּ�����Ĳļ��������������</b></font></p>
<p align="center">��</p>
<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;            
���뵥λ:</p><form method="POST" action="save.asp?id=<%=uid%>" name="form1">
<table align="center" border="1" width="68%" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolor="#000000" height="261">
  <tr>
    <td width="25%" bordercolorlight="#000000" height="36">����Ĳ���:<select size="1" name="hcmc">
        <option value="����">����</option>
        <option value="ī��">ī��</option>
        <option value="ɫ��">ɫ��</option>
        <option value="ɫ����">ɫ����</option>
        <option value="����">����</option>
        <option value="�ƶ�Ӳ��">�ƶ�Ӳ��</option>
        <option value="�������">�������</option>
        <option value="��ӡֽ">��ӡֽ</option>
        <option value="����">����</option>
                    
        &nbsp;            
        </select>
     
    </td>
    <td width="36%" bordercolorlight="#000000" height="36">�Ĳ��ͺ�:<input type="text" name="XH" size="24"></td>
    <td width="24%" bordercolorlight="#000000" height="36">����:<input type="text" name="SL" size="9" value="1"></td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#000000" height="45" colspan="3">��������:<textarea rows="2" name="sqsx" cols="81"> </textarea></td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#000000" height="15" colspan="3">��λ������ǩ��:            
      ͬ��<input type="radio" value="1" name="DWFZR" checked>����ͬ��<input type="radio" name="DWFZR" value="0"><textarea rows="2" name="S3" cols="65"></textarea>
      <p>&nbsp;ǩ����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;            
      ʱ�䣺</p>
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#000000" height="49" colspan="3">��Ϣ�������:ͬ��<input type="radio" value="1" name="ZXYJ" checked>����ͬ��<input type="radio" name="ZXYJ" value="0">&nbsp;&nbsp;            
      <textarea rows="2" name="zxyj" cols="65"></textarea>
      <p>&nbsp;ǩ����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;            
      &nbsp;&nbsp;&nbsp;&nbsp; ʱ�䣺</p>           
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#000000" height="16" colspan="3">���쵼���:ͬ��<input type="radio" value="1" name="JLDYJ" checked>����ͬ��<input type="radio" name="JLDYJ" value="0">&nbsp;&nbsp;&nbsp;&nbsp;            
      <textarea rows="2" name="jldyj" cols="65"></textarea>
      <p>&nbsp;ǩ����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      ʱ�䣺</p>
    </td>
  </tr>
  <tr>
    <td width="85%" bordercolorlight="#000000" height="16" colspan="3">
      <p align="center">��<input type="submit" name="B1" value="�ύ����">&nbsp; 
      <input type="reset" value="ȫ����д" name="B2"></p>
      ��
      <p>��</p>
    </td>
  </tr>
  <tr>
    <td width="25%" bordercolorlight="#000000" height="16">��</td>
    <td width="36%" bordercolorlight="#000000" height="16">��</td>
    <td width="24%" bordercolorlight="#000000" height="16">��</td>
  </tr>
</table>
 </form>
</body>

</html>