<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
%>
<html>
<script Language="javascript" src="hw.js"></script>
<script Language="javascript" src="date.js"></script>
<script Language="javascript" src="st_js.asp"></script>
<script>
ico_top=10;
ico_left=10;
ico_bottom=45;
ico_x=65;
ico_y=65;
var intInterval=20;
var intDelay=20;
</script>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="imagetoolbar" content="no">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
%>
<title>��E�߽����� - 51eline.com</title>
<style type="text/css">
<!--
.shakeimage{
position:relative
}--></style>

<script Language="javascript" src="menu.asp?id=1"></script>
<link rel="stylesheet" type="text/css" href="stylesheet.asp?id=1">
<script language="JavaScript">
document.onmousedown=click;
function chatWin() {
woiwo=window.open('chat/jhchat.asp','sjjh','resizable=no,scrollbars=no,toolbar=no,menubar=no,fullscreen=no');
woiwo.moveTo(0,0)
woiwo.resizeTo(screen.availWidth,screen.availHeight)
woiwo.outerWidth=screen.availWidth
woiwo.outerHeight=screen.availHeight
}
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
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<style type="text/css">
BODY {CURSOR: url('chat/56.ani')}
</style>
</head>
<script Language="javascript">window.name="_desktop";</script>
<body onresize=setico(); scroll=no onMouseOver="if (style.behavior==''){style.behavior='url(#default#homepage)';setHomePage('http://51eline.com')}" onmousemove="s_t();" onload="setico();" id=bglayer background="f2.gif">
<script lauguage="javascript" src="STARMOON.JS"></script>
<img id="bgimg" src="images/spacer.gif" width="100%" height="100%">

<table id=taskbar border="0" style="position: absolute;width:expression(screen.width-0);bottom:-1;left:-1" cellpadding="0" cellspacing="0"  class="up" style="padding: 1">
    <tr>
      <td width="50">
        <div align="center">
          <table border="0" onclick="startclick();" id="startmenu" cellpadding="0" cellspacing="0" width="100%" class="up">
            <tr>
              <td><a href="#" onClick="return false;" id=s_img><img border="0"  onmousemove='if(startmenu.className=="up"){event.srcElement.alt="���������￪ʼ��";}else{event.srcElement.alt="";}' src="51eline/start_menu/start.gif" width="45" height="18" align="center" vspace="1" hspace="1"></a></td>
            </tr>
          </table>
        </div>
      </td>
      <td width="120">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0" width="116">
            <tr>
              <td width="10" align="center"><img border="0" src="51eline/start_menu/ii16.gif" width="7" height="22"></td>
              <td width="24" align="center"><img title="���� Internet Explorer �����" border="0" onclick=pop("#") class=menuout onmouseover=over(this); onmouseout=out(this); onmousedown=down(this); onmouseup=up(this); src="51eline/start_menu/ie162.gif" width="20" height="20"></td>
              <td width="24" align="center"><img title="����ͼ��" onclick="setico()" border="0" class=menuout onmouseover=over(this); onmouseout=out(this); onmousedown=down(this); onmouseup=up(this); src="51eline/start_menu/desktop.gif" width="20" height="20"></td>
              <td width="24" align="center"><img title="���� Outlook Express(Eline_Email@etang.com)" onclick=location.href="mailto:Eline_Email@etang.com" border="0" class=menuout onmouseover=over(this); onmouseout=out(this); onmousedown=down(this); onmouseup=up(this); src="51eline/start_menu/oe16.gif" width="21" height="20"></td>
              <td width="24" align="center"><img title="E�߽�����Դ������" onclick=pop("explorer.htm") border="0" class=menuout onmouseover=over(this); onmouseout=out(this); onmousedown=down(this); onmouseup=up(this); src="51eline/start_menu/soft16.gif" width="20" height="20"></td>
              <td width="10" align="center"><img border="0" src="51eline/start_menu/ii16.gif" width="7" height="22"></td>
            </tr>
          </table>
        </div>
      </td>
      <td></td>
      <td width="140">
        <div align="center">
          <table border="0" cellpadding="0" cellspacing="0"  height="24" width="140"  class="down1" >
            <tr>
              <td width="20"><img border="0" src="51eline/start_menu/speak.gif" width="19" height="18"></td>
              <td width="20"><img border="0" src="51eline/start_menu/shuru.gif" width="18" height="18"></td>
              <td width="20"><img border="0" src="51eline/start_menu/netlink16.gif" ondblclick=alert("61.154.121.246") title="61.154.121.246" width="18" height="18"></td>
              <td width="20"><img title="����...."  border="0" src="51eline/start_menu/oicq16.gif" width="18" height="18"></td>
              <td width="60" style="cursor:default" id=clock onmouseover=show("date") onmouseout=hide("date")><script>tick();</script></td>
            </tr>
          </table>
        </div>
      </td>
    </tr>
  </table>
<div class=date id="date"  onmouseover=show("date") onmouseout=hide("date") style="position:absolute;left:expression(screen.width-95);bottom:30;display:none;z-index:11"><script>time();</script></div>
<div id=ico>

<div id="ico1" class=ico ondblclick=pop("zhuangtai.asp") onmousedown=m("ico1");spanico1.focus(); onblur=spanico1.blur(); title="�鿴����״̬"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/document.gif" vspace="3"><br><a  id=spanico1 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>�ҵ�״̬</a></div>

<div id="ico2" class=ico onClick="chatWin()" onmousedown=m("ico2");spanico2.focus(); onblur=spanico2.blur(); title="���뽭�������Ҵ���"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/computer.gif" vspace="3"><br><a  id=spanico2 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';><font color="#ffff00"><strong>ɱ�뽭��</strong></font></a></div>

<div id="ico3" class=ico ondblclick=pop("jhmp/index.asp") onmousedown=m("ico3");spanico3.focus(); onblur=spanico3.blur(); title="���������Ҹ����ϴ��չ�׼û��"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/nethood.gif" vspace="3"><br><a  id=spanico3 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>��������</a></div>

<div id="ico4" class=ico ondblclick=pop("yamen/laofan.asp") onmousedown=m("ico4");spanico4.focus(); onblur=spanico4.blur(); title="�����㻵��Ŷ...��Ȼ...�ٺ�..."><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/rec.gif" vspace="3"><br><a  id=spanico4 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>��Ѻ�˷�</a></div>

<div id="ico5" class=ico ondblclick=pop("jhmp/wuguan.asp") onmousedown=m("ico5");spanico5.focus(); onblur=spanico5.blur(); title="����������.�ð�ʦ���ܽ���"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/explorer.gif" vspace="3"><br><a  id=spanico5 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>��������</a></div>

<div id="ico14" class=ico ondblclick=pop("http://zhzx.jjedu.org/eline/chat/xinshou.asp") onmousedown=m("ico14");spanico14.focus(); onblur=spanico14.blur(); title="��������.�������.ϵͳ����"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/helper.gif" vspace="3"><br><a  id=spanico14 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>��������</a></div>

<div id="ico66" class=ico ondblclick=pop("jhmp/money.asp") onmousedown=m("ico66");spanico66.focus(); onblur=spanico66.blur(); title="��Ǯ�.��ͬ����һ����ؼ�"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/folder.gif" vspace="3"><br><a  id=spanico66 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>��ȡ����</a></div>

<div id="ico73" class=ico ondblclick=pop("jl/jlmain.asp") onmousedown=m("ico73");spanico73.focus(); onblur=spanico73.blur(); title="����ÿ�췢���Ĵ���"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/readme.gif" vspace="3"><br><a  id=spanico73 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>��������</a></div>

<div id="ico83" class=ico ondblclick=pop("top/top.htm") onmousedown=m("ico83");spanico83.focus(); onblur=spanico83.blur(); title="�鿴��������"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/find.gif" vspace="3"><br><a  id=spanico83 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>��������</a></div>

<div id="ico84" class=ico ondblclick=pop("wish/wish.asp") onmousedown=m("ico84");spanico84.focus(); onblur=spanico84.blur(); title="����ȥ��.��������.ף����"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/book.gif" vspace="3"><br><a  id=spanico84 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>������Ը</a></div>

<div id="icon232" class="ico" ondblclick=open("work/changework.asp","","")  onmousedown=m("icon232");spanicon232.focus(); spanicon232.focus();  title="�뻻��ְҵ�������"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/hlp.gif" vspace="3"><img class=shortcut border="0" src="51eline/sc.gif"><br><a  id=spanicon232 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>ѡ��ְҵ</a></div>

<div   id="vote" class=ico ondblclick=pop("hcjs/jhjs/hua.asp") onmousedown=m("vote")  title="Ҫ��ȫ������.�����������">
<p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/readme.gif" vspace="3">
<br>�����ʻ�</div>
<%if sjjh_grade>=10 then%>
<div id="ico92" class=ico ondblclick=pop("admin/login.asp") onmousedown=m("ico92");spanico92.focus(); onblur=spanico92.blur(); title="�������"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/qq.gif" vspace="3"><br><a  id=spanico92 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';><font color=#ff0000><b>�������</b></font></a></div>
<%end if%>
<div id="ico259" class=ico ondblclick=pop("hcjs/jhjs/cw.asp") onmousedown=m("ico259");spanico259.focus(); onblur=spanico259.blur(); title="������ϲ����С�����"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/cw.gif" vspace="3"><br><a  id=spanico259 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>�������</a></div>

<div id="ico280" class=ico ondblclick=pop("XIAOHAI/indexsheep.asp") onmousedown=m("ico280");spanico280.focus(); onblur=spanico280.blur(); title="���ӳ���ᱨ�����"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/boy.gif" vspace="3"><br><a  id=spanico280 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>����С��</a></div>

<div id="ico445" class=ico ondblclick=pop("garden/hua.asp") onmousedown=m("ico445");spanico445.focus(); onblur=spanico84.blur(); title="����Ļ������򲻵���Ŷ"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/hua.gif" vspace="3"><br><a  id=spanico445 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>���ػ�԰</a></div>

<div id="ico446" class=ico ondblclick=pop("gupiao/stock.asp") onmousedown=m("ico446");spanico446.focus(); onblur=spanico84.blur(); title="���°棬һ���ܸ���һ����ϲ��"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/gupiao.gif" vspace="3"><br><a  id=spanico446 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>������Ʊ</a></div>

<div id="ico447" class=ico ondblclick=pop("yhy/index.asp") onmousedown=m("ico447");spanico447.focus(); onblur=spanico84.blur(); title="Ҫע������Ŷ"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/qinglou.gif" vspace="3"><br><a  id=spanico447 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>������¥</a></div>

<div id="ico450" class=ico ondblclick=pop("bbs/index.asp") onmousedown=m("ico450");spanico450.focus(); onblur=spanico84.blur(); title="�����������Ա����"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/bbs.gif" vspace="3"><br><a  id=spanico450 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>������̳</a></div>

<div id="ico448" class=ico ondblclick=pop("JHMP/bentop4.asp") onmousedown=m("ico448");spanico448.focus(); onblur=spanico84.blur(); title="���ɾ�����"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/yishi.gif" vspace="3"><br><a  id=spanico448 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>��������</a></div>

<div id="ico449" class=ico ondblclick=pop("taiqiu/zh.asp") onmousedown=m("ico449");spanico449.focus(); onblur=spanico84.blur(); title="��ת���Ķ����ɶ���"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/shichang.gif" vspace="3"><br><a  id=spanico449 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>ת���г�</a></div>

<div id="ico662" class=ico ondblclick="javascript:window.showModelessDialog('http://zhzx.jjedu.org/eline/readme.asp','','dialogWidth:542px;dialogHeight:416px;scroll:1;status:0;edge:raised')" onmousedown=m("ico662");spanico662.focus(); onblur=spanico662.blur(); title="վ��˵��"><p align="center"><img onMouseOver="init(this);rattleimage()" onMouseOut="stoprattle(this)"  class="shakeimage" border="0" width="32" height="32" src="51eline/home.gif" vspace="3"><br><a  id=spanico665 href="javascript:" class=onblur onfocus=this.className='onfocus'; onblur=this.className='onblur';>���ڱ�վ</a></div>
</div>

<script>setico();</script>
<div  style="position: absolute;bottom:35; left:expression(screen.width-150) ; width: 150; height:45;color:#FFFFFF;align:right">
<!--#include file="z_showvisual.asp"-->
<!-- ��Ȩ��Ϣ�뱣�� -->
��Ȩ����E�߽�����&#8482;<br>
������һ����<br>
�汾��ELINE V8.7���ް�<br>
�����£�2004-02-24<br>
</div>

<script Language="javascript" src="startmenu.asp"></script>
<SCRIPT LANGUAGE="JavaScript">
<!--
menu.AddItem("logoff","xp/start_logoff","24","����<font class=w2kfont>(<u>L</u>)</font>","rbpm","http://51eline.com/old.asp");
menu.AddItem("shut","xp/start_shut","24","�˳�<font class=w2kfont>(<u>U</u>)</font>","rbpm","exit.asp");
document.writeln(menu.GetMenu());
//-->
</SCRIPT>

<!--[if IE]>
<DIV id="ie5menu" style="padding:2px" class=up>
<script>
function  display(){
showModelessDialog("display.asp?id=1",window,"status:no;dialogWidth:420px;dialogHeight:465px");
//window.open("display.asp","window","status=no,Width=420,Height=430");
}
link('0',' �����<font class=w2kfont>(<u>D</u>)</font>','0',"-");hr();
link('setico()',' ����ͼ��','none',"");
link('addfav()',' �����ղ�','none',"");
link('shuaxin()','ˢ��<font class=w2kfont>(<u>E</u>)</font>','none',"");
link('pop("fileadd.asp")','�½�<font class=w2kfont>(<u>W</u>)</font>','0',"�ļ�-");
hr();link("display();",' ����(R)','none',"");
</script>
</DIV><![endif]-->
<SCRIPT language=JavaScript>
if (document.all&&window.print){
document.oncontextmenu=showmenuie5
document.onkeydown = keypress
document.body.onclick=hidemenuie5
}
</SCRIPT>
</body>

<NOSCRIPT><IFRAME SRC="*.html"></IFRAME></NOSCRIPT>
</html>