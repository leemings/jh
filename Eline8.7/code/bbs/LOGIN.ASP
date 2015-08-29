<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="inc/char_login.asp"-->
<!--#include file="md5.asp"-->

<%
dim regjm
randomize timer
regjm=int(rnd*8998)+1000

stats="填写登陆信息"
dim url,comeurl

call nav()
call head_var(2,0,"","")
if request("action")="chk" then
	call chklogin()
	if founderr then call dvbbs_error()
else
	url=Request.ServerVariables("HTTP_REFERER")
%>
<form action="../check.asp" method="post" name="login" onsubmit="return(check())" target="_self" ;>
    <input type="hidden" name="CookieDate" value="0" >
    <input type="hidden" name="bbs" value="1" >
    <input type=hidden name=regjm value="<%=regjm%>">
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
    <tr>
    <th valign=middle colspan=2 align=center height=25><font color=#cc0000>请完整、准确输入您的江湖资料登陆</font></td></tr>
<tr><td class=tablebody1 valign=middle align=center colspan=2> 
     <img alt="『E线江湖』" border="0" height="32" src="images/eline.gif" width="162">
    </td></tr>
    <tr><td class=tablebody1 valign=middle align=center colspan=2> 
     <script src="../online.asp"></script>
    </td></tr>
    <tr>
    <td class=tablebody1 valign=middle>请输入您的ＩＤ</td>
    <td class=tablebody1 valign=middle><input  type="password" onKeyDown="ctlent()" maxlength=8 name=id>
    </tr>
    <tr>
    <td class=tablebody1 valign=middle>请输入您的用户名</td>
    <td class=tablebody1 valign=middle><INPUT name="name" type=text> &nbsp; <a href="reg.asp">没有注册？</a></td></tr>
    <tr>
    <td class=tablebody1 valign=middle>请输入您的密码</td>
    <td class=tablebody1 valign=middle><INPUT name="pass" type=password> &nbsp; <a href="lostpass.asp">忘记密码？</a></td></tr>
    </td></tr>
	<input type=hidden name=comeurl value="<%=Request.ServerVariables("HTTP_REFERER")%>">
    <tr>
    <td class=tablebody1 valign=middle>请输入您的认证码</td>
    <td class=tablebody1 valign=middle><INPUT name="regjm1" type=text> &nbsp; <%=regjm%></td></tr>

    <tr>
    <td class=tablebody1 valign=top width=30% ><b>Cookie 选项</b><BR> 请选择你的 Cookie 保存时间，下次访问可以方便输入。</td>
    <td valign=middle class=tablebody1>                <input type=radio name=CookieDate value=0 checked>不保存，关闭浏览器就失效<br>
                <input type=radio name=CookieDate value=1>保存一天<br>
                <input type=radio name=CookieDate value=2>保存一月<br>
                <input type=radio name=CookieDate value=3>保存一年<br>                </td></tr>
    <tr>
    <td valign=top width=30% class=tablebody1><b>隐身登陆</b><BR> 您可以选择隐身登陆，论坛会员将在用户列表看不到您的信息。</td>
    <td valign=middle class=tablebody1>                <input type=radio name=userhidden value=2 checked>正常登陆<br>
                <input type=radio name=userhidden value=1>隐身登陆<br>
                </td></tr>
	<input type=hidden name=comeurl value=<%=url%>>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name=submit value="登 陆"></td></tr></table>
</form>


<script>
ie = (document.all)? true:false
if (ie){
function ctlent(eventobject){
if(window.event.keyCode==13 && check()!=false){this.document.login.submit();}
}
}
</script>

<SCRIPT language="JavaScript">     
<!--     
var targetwin="jhwindow";    
function check()    
{    
var id=document.login.id.value;    
var youname=document.login.name.value;    
var mima=document.login.pass.value;    
var regno1=document.login.regjm.value;    
var regno2=document.login.regjm1.value;    
if(id=="" || id==null){window.alert("请输入ID，ID为数字！^_^!");return false;}    
if(youname=="" || youname==null){window.alert("大侠，你要来玩，请您输入大名！^_^!");return false;}    
if( CheckIfEnglish(youname) )    
{    
window.alert("名字中不能有非法字符、英文、数字；只能使用中文片假名！^_^!");    
return false;    
}    
if(mima=="" || mima==null){window.alert("大侠，你要来玩，没有密码怎么进入呀？^_^!");return false;}    
if(regno1!=regno2){window.alert('请输入正确的认证码:'+regno1+'！');return false;}    
return true;    
}    
    
function jhdo(){    
OpenWindow=window.open('',targetwin,'resizable=no,scrollbars=yes,toolbar=no,menubar=no');    
OpenWindow.moveTo(0,0)    
OpenWindow.resizeTo(screen.availWidth,screen.availHeight)    
OpenWindow.outerWidth=screen.availWidth    
OpenWindow.outerHeight=screen.availHeight    
}    
function CheckIfEnglish( String )    
{     
    var Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890`=~!@#$%^&*()_+[]{}\\|/?.>,<;:'\-?<>/～！·＃￥％…；‘’：“”—＊（　）—＋｜－＝、／。，？《》↑↓⊙●★☆■♀ 『』◆◥◤◣ Ψ ※ →№←㊣∑⌒ 〖〗 ＠ξζω□ ∮〓※ ▓∏卐【 】▲△√ ∩¤々 ♀♂∞ ㄨ ≡↘↙ ＆◎Ю┼┏ ┓田 ┃▎○╪┗┛ ∴ ①②③④⑤⑥⑦⑧ \"";    
     var i;    
     var c;    
      if(String=="chonger" || String=="8848") {return false;}    
     for( i = 0; i < String.length; i ++ )    
     {    
          c = String.charAt( i );    
	  if (Letters.indexOf( c ) < 0)    
	     return false;    
     }    
     return true;    
}    
    
//-->    
</SCRIPT>
<%
end if
call footer()
%>