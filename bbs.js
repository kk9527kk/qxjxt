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
//以下为版主设置模块的JS./////////////////////////////////////////////////////
function adminadd()
{
        if  (document.typeadmin.typeadmin.value=="")
        {
            alert("总版主名称不能为空!");
            document.typeadmin.typeadmin.focus();
            return false ;
        } 
var user=document.typeadmin.typeadmin.value
if( confirm("你确定要将("+user+")设置为总版主吗？\n"))
{ window.location = "admin_typeadminadd.asp?Id=2&user="+user;
}
 return false;
}

function badminadd()
{
        if  (document.btypeadmin.btypeadmin.value=="")
        {
            alert("大版主名称不能为空!");
            document.btypeadmin.btypeadmin.focus();
            return false ;
        } 
var user=document.btypeadmin.btypeadmin.value
var room=document.btypeadmin.broom.value
var roomname=document.btypeadmin.broom.options[document.btypeadmin.broom.options.selectedIndex].text
if( confirm("你确定要将("+user+")设置为("+roomname+")的大版主吗？\n"))
{ window.location = "admin_typeadminadd.asp?Id="+room+"&user="+user;
}
 return false;
}

function sadminadd()
{
        if  (document.stypeadmin.stypeadmin.value=="")
        {
            alert("小版主名称不能为空!");
            document.stypeadmin.stypeadmin.focus();
            return false ;
        }        
var user=document.stypeadmin.stypeadmin.value
var room=document.stypeadmin.sroom.value
var roomname=document.stypeadmin.sroom.options[document.stypeadmin.sroom.options.selectedIndex].text
if( confirm("你确定要将("+user+")设置为("+roomname+")的小版主吗？\n"))
{ window.location = "admin_typeadminadd.asp?Id="+room+"&user="+user;
}
 return false;
}
function admincl()
{
        if  (document.typeadmin.typeadmin.value=="")
        {
            alert("总版主名称不能为空!");
            document.typeadmin.typeadmin.focus();
            return false ;
        }
var user=document.typeadmin.typeadmin.value
if( confirm("你确定要将罢免("+user+")的总版主职务吗？\n"))
{ window.location = "admin_typeadmincl.asp?Id=2&user="+user;
}
 return false;
}

function badmincl()
{
        if  (document.btypeadmin.btypeadmin.value=="")
        {
            alert("大版主名称不能为空!");
            document.btypeadmin.btypeadmin.focus();
            return false ;
        }
var user=document.btypeadmin.btypeadmin.value
var room=document.btypeadmin.broom.value
var roomname=document.btypeadmin.broom.options[document.btypeadmin.broom.options.selectedIndex].text
if( confirm("你确定要罢免("+user+")("+roomname+")的大版主资格吗？\n"))
{ window.location = "admin_typeadmincl.asp?Id="+room+"&user="+user;
}
 return false;
}
function sadmincl()
{
        if  (document.stypeadmin.stypeadmin.value=="")
        {
            alert("小版主名称不能为空!");
            document.stypeadmin.stypeadmin.focus();
            return false ;
        } 
var user=document.stypeadmin.stypeadmin.value
var room=document.stypeadmin.sroom.value
var roomname=document.stypeadmin.sroom.options[document.stypeadmin.sroom.options.selectedIndex].text
if( confirm("你确定要罢免("+user+")("+roomname+")的小版主资格吗？\n"))
{ window.location = "admin_typeadmincl.asp?Id="+room+"&user="+user;
}
 return false;
}
//以上为版主设置模块的JS./////////////////////////////////////////////////////
function deluser()
{
if( confirm("你确定要删除这个会员吗？\n"))
{ return true;
}
 return false;
}
//以下为版块管理模块的JS./////////////////////////////////////////////////////
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
            alert("版块名称不能为空!");
            document.btypeadd.roomname.focus();
            return false ;
        }
var j=document.btypeadd.roomname.value
if( confirm("你确定要添加("+j+")版块吗？\n"))
{ return true;
}
 return false;
}

function deltype(id,j)
{	
if( confirm("你确定要删除("+j+")版块吗？\n\n删除版块后,该版块下的所有(小版块)\n及(贴子)将被删除,请在删除前作处理!\n"))
{ window.location = "admin_typedelete.asp?Id="+id;
}
 return false;
}

function updatetype(i)
{
if( confirm("你确定要将("+i+")版块作这样的更改吗？\n"))
{ return true;
}
 return false;
}
//以上为版块管理模块的JS./////////////////////////////////////////////////////
//以下为贴子管理模块的JS./////////////////////////////////////////////////////
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
      alert("你没有选择，无法处理!!");
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
//以上为贴子管理模块的JS./////////////////////////////////////////////////////