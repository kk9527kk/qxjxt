function load()
{
	var names=navigator.appName
	var vers=navigator.appVersion
	if(names=='Netscape')
	{
		window.location.reload(true)
	}else
	{
		history.go(0)
	}
}
//����Ϊ��������ģ���JS./////////////////////////////////////////////////////
function adminadd()
{
        if  (document.typeadmin.typeadmin.value=="")
        {
            alert("�ܰ������Ʋ���Ϊ��!");
            document.typeadmin.typeadmin.focus();
            return false ;
        } 
var user=document.typeadmin.typeadmin.value
if( confirm("��ȷ��Ҫ��("+user+")����Ϊ�ܰ�����\n"))
{ window.location = "admin_typeadminadd.asp?Id=2&user="+user;
}
 return false;
}

function badminadd()
{
        if  (document.btypeadmin.btypeadmin.value=="")
        {
            alert("��������Ʋ���Ϊ��!");
            document.btypeadmin.btypeadmin.focus();
            return false ;
        } 
var user=document.btypeadmin.btypeadmin.value
var room=document.btypeadmin.broom.value
var roomname=document.btypeadmin.broom.options[document.btypeadmin.broom.options.selectedIndex].text
if( confirm("��ȷ��Ҫ��("+user+")����Ϊ("+roomname+")�Ĵ������\n"))
{ window.location = "admin_typeadminadd.asp?Id="+room+"&user="+user;
}
 return false;
}

function sadminadd()
{
        if  (document.stypeadmin.stypeadmin.value=="")
        {
            alert("С�������Ʋ���Ϊ��!");
            document.stypeadmin.stypeadmin.focus();
            return false ;
        }        
var user=document.stypeadmin.stypeadmin.value
var room=document.stypeadmin.sroom.value
var roomname=document.stypeadmin.sroom.options[document.stypeadmin.sroom.options.selectedIndex].text
if( confirm("��ȷ��Ҫ��("+user+")����Ϊ("+roomname+")��С������\n"))
{ window.location = "admin_typeadminadd.asp?Id="+room+"&user="+user;
}
 return false;
}
function admincl()
{
        if  (document.typeadmin.typeadmin.value=="")
        {
            alert("�ܰ������Ʋ���Ϊ��!");
            document.typeadmin.typeadmin.focus();
            return false ;
        }
var user=document.typeadmin.typeadmin.value
if( confirm("��ȷ��Ҫ������("+user+")���ܰ���ְ����\n"))
{ window.location = "admin_typeadmincl.asp?Id=2&user="+user;
}
 return false;
}

function badmincl()
{
        if  (document.btypeadmin.btypeadmin.value=="")
        {
            alert("��������Ʋ���Ϊ��!");
            document.btypeadmin.btypeadmin.focus();
            return false ;
        }
var user=document.btypeadmin.btypeadmin.value
var room=document.btypeadmin.broom.value
var roomname=document.btypeadmin.broom.options[document.btypeadmin.broom.options.selectedIndex].text
if( confirm("��ȷ��Ҫ����("+user+")("+roomname+")�Ĵ�����ʸ���\n"))
{ window.location = "admin_typeadmincl.asp?Id="+room+"&user="+user;
}
 return false;
}
function sadmincl()
{
        if  (document.stypeadmin.stypeadmin.value=="")
        {
            alert("С�������Ʋ���Ϊ��!");
            document.stypeadmin.stypeadmin.focus();
            return false ;
        } 
var user=document.stypeadmin.stypeadmin.value
var room=document.stypeadmin.sroom.value
var roomname=document.stypeadmin.sroom.options[document.stypeadmin.sroom.options.selectedIndex].text
if( confirm("��ȷ��Ҫ����("+user+")("+roomname+")��С�����ʸ���\n"))
{ window.location = "admin_typeadmincl.asp?Id="+room+"&user="+user;
}
 return false;
}
//����Ϊ��������ģ���JS./////////////////////////////////////////////////////
function deluser()
{
if( confirm("��ȷ��Ҫɾ�������Ա��\n"))
{ return true;
}
 return false;
}
//����Ϊ������ģ���JS./////////////////////////////////////////////////////
function benableit()
{
document.btypeadd.broom.disabled=false;
document.btypeadd.roomname.disabled=false;
document.btypeadd.s1.disabled=false;
}
function btypeadds()
{
        if  (document.btypeadd.roomname.value=="")
        {
            alert("������Ʋ���Ϊ��!");
            document.btypeadd.roomname.focus();
            return false ;
        }
var j=document.btypeadd.roomname.value
if( confirm("��ȷ��Ҫ���("+j+")�����\n"))
{ return true;
}
 return false;
}

function deltype(id,j)
{	
if( confirm("��ȷ��Ҫɾ��("+j+")�����\n\nɾ������,�ð���µ�����(С���)\n��(����)����ɾ��,����ɾ��ǰ������!\n"))
{ window.location = "admin_typedelete.asp?Id="+id;
}
 return false;
}

function updatetype(i)
{
if( confirm("��ȷ��Ҫ��("+i+")����������ĸ�����\n"))
{ return true;
}
 return false;
}
//����Ϊ������ģ���JS./////////////////////////////////////////////////////
//����Ϊ���ӹ���ģ���JS./////////////////////////////////////////////////////
function delallarticle(tip,url)
{   var i;
    var truei;
    truei=0;
    for(i=0;i<document.all.del.length;i++)
    { if (document.all.del[i].checked==1)
      {
       truei=truei+1;
      }
    }
    if (truei <= 0)
     {
      alert("��û��ѡ���޷�����!!");
      return false;
     }
if( !confirm(tip+"\n"))
{ 
	return true;   
}
     frmpost.action=url;
     frmpost.submit();
 return true;
}

function dels(tips)
{
if( confirm(tips+"\n"))
{ return true;
}
 return false;
}

function  checkallarticle()
{
	var i;
	for(i=0;i<document.all.del.length;i++)
	{
		if (document.all.del[i].value != 0)
		{
		document.all.del[i].checked = document.all.Select.checked
		}
	}
}
//����Ϊ���ӹ���ģ���JS./////////////////////////////////////////////////////