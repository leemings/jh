<%@ LANGUAGE=VBScript codepage ="936" %>

<%Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if chatbgcolor="" then chatbgcolor="008888"%>
<html><head><META http-equiv='Content-Type' content='text/html; charset=gb2312'>
<style type=text/css>
body {
    CURSOR: url('yinsu.ani');
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
	}
A 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:'000000'; font-size: 9pt;}
A:hover 	{font-family:verdana,arial,helvetica,Tahoma; text-decoration: none; color:ff6600;  font-size: 9pt;CURSOR: url('caidan.cur');}
.normal{
cursor:default;
font-size: 9pt;
background-color:#006699;
border:#000000 1px solid;
padding:1px;
}
.mouseover{
cursor:default;
font-size: 9pt;
background-color:#000000;
border-top:#006699 1px solid;
border-right:#000000 1px solid;
border-bottom:#000000 1px solid;
border-left:#006699 1px solid;
padding:1px;
}
.mousedown{
cursor:default;
font-size: 9pt;
background:buttonface;
border-top:buttonshadow 1px solid;
border-right:buttonhighlight 1px solid;
border-bottom:buttonhighlight 1px solid;
border-left:buttonshadow 1px solid;
padding-top:2px;
padding-right:0px;
padding-bottom:0px;
padding-left:2px;
}
</style>
<script language="JavaScript">
function s(list){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}
function rc(list){if(list!="0"){parent.f2.document.af.sytemp.value=list;parent.f2.document.af.sytemp.focus();}}
</script>
<script language="JavaScript">
document.onmousedown=click;
function click(){if(event.button==2){if(confirm("�Ƿ���ʾ�Լ�����Ʒ��","������ʾ")){window.open('chat/wupin.asp','wupin','scrollbars=yes,resizable=yes,width=500,height=400');}}}
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
<script language="JavaScript"> 
function WM_readCookie(name)
{
if(document.cookie == '')
    return false;
else
    return unescape(WM_getCookieValue(name));
}

var Read_cookie=WM_readCookie("yescnet_openday");

function openday()
{
	Arrset=window.showModalDialog('help.asp','�Ǻ�','dialogWidth:435px;dialogHeight:297px;scroll:0;status:0;help:1;edge:raised');
	setCookie("yescnet_openday",Arrset.charAt(0),1);
	setCookie("yescnet_MIDw",Arrset.charAt(1),1);
}
</script>
</head>
<body bgcolor="#006699" leftmargin="0" topmargin="0" bgproperties="fixed" oncontextmenu=self.event.returnValue=false>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr> 
        <td width="100%" height="28"> <div align="center"><font color="#CCCCFF" size="2"><strong>�� 
        �� ��</strong></font></div></td>
      </tr>
    </table>
<div align="center">
  <center>
    <table width="140" height="565" border="1" cellpadding="1" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
      <tr> 
        <td width="100%" align="center">
		<div align="left"> 
            <table width="130" height="565" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#006699">
              <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a style="text-decoration: none" href="#" title="�ⲻ�ɲ���Ŷ^_^" target="_self" onclick="javascript:openday();"><font color="#b7d4f1">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" title="�鿴�Լ�����ϸ״̬" onMouseOver="window.status='�鿴�Լ�����ϸ״̬��';return true" onMouseOut="window.status='';return true" onclick="javascript:window.open('zt/index.asp','zhuantai','scrollbars=no,toolbar=no,menubar=no,location=no,status=no,resizable=no,width=500,height=530')" target="_self"><font color="#b7d4f1">�ҵ�״̬</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href=# target="_self" onclick="javascript:window.showModelessDialog('wnl.htm','','dialogWidth:682px;dialogHeight:320px;scroll:0;status:0;edge:raised')" title="��������ʱ�����ǲ��ó�"><font color="#b7d4f1">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" onClick="window.open('../hcjs/jhzb/index.asp','wupin','scrollbars=yes,resizable=yes,width=550,height=300')" title="�鿴һ���Լ���ʲô��Ʒ"><font color="#b7d4f1">�ҵ���Ʒ</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" title="�������书����" onClick="window.open('../top/top.htm','','scrollbars=no,resizable=yes,width=780,height=400')"><font color="#b7d4f1">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" onClick="window.open('../hcjs/jhzb/myhua.asp','myhua','scrollbars=no,resizable=no,width=430,height=340')"> 
                          <font color="#b7d4f1">�ҵ��ʻ�</font> </a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			 
              <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a target="_blank" href="../yamen/yiyuan.ASP"><font color="#b7d4f1">����ҽԺ</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../DIAOYU/Diaoyu.asp" target="_blank"><font color="#b7d4f1">������ͷ</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" onClick="window.open('../jtbc/buy.asp','','scrollbars=no,resizable=no,width=400,height=400')"><font color="#b7d4f1">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../wabao/wabao.asp','diaoyu','scrollbars=yes,resizable=yes,width=450,height=340')"><font color="#b7d4f1">�ھ򱦲�</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" onClick="window.open('../poll/poll.asp')" ><font color="#b7d4f1">����ѡ��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../dzwq/index.asp','dzwq','scrollbars=yes,resizable=yes,width=780,height=460')" title="����������"><font color="#b7d4f1">�������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../hcjs/pub/pub.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><font color="#FFFF99">������ջ</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../peiyao/peiyao.asp','peiyao','scrollbars=yes,resizable=yes,width=780,height=460')"><font color="#FFFF99">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../hcjs/ww/indexclean.asp','','scrollbars=no,resizable=no,width=700,height=400')"><font color="#FFFF99">����ɣ��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../work/famu/famumain.asp','famu','scrollbars=yes,resizable=yes,width=500,height=370,resizable=no,scrollbars=no,toolbar=no,menubar=no,fullscreen=no')"><font color="#FFFF99">��ɽ��ľ</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../hcjs/MEIYONG/MEIYONG.asp','','scrollbars=yes,resizable=yes,width=778,height=434')"><font color="#FFFF99">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../hcjs/jhjs/cw.asp','','scrollbars=no,resizable=yes,width=520,height=450')"><font color="#FFFF99">�������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../hunyin/yuelao.asp','','scrollbars=yes,resizable=yes,width=550,height=450')"><font color="#FFFF99">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../hcjs/jhjs/wuqi.asp','','scrollbars=yes,resizable=yes,width=650,height=550')"><font color="#FFFF99">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../jhmp/setwg.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><font color="#FFFF99">�书����</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../hcjs/jhjs/anqi.asp','','scrollbars=yes,resizable=yes,width=740,height=400')"><font color="#FFFF99">���Ű���</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a target="_blank" href="../yamen/Laofan1.asp"> 
                          <font color="#FFFF99">���˯��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../hunyin/killer.asp','','scrollbars=no,resizable=no,width=700,height=400')"><font color="#FFFF99">Ƹ��ɱ��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../taiqiu/yule.asp','yule','scrollbars=no,resizable=no,width=320,height=340')"><font color="#00CC66">���ֽ��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../jhmp/maidou.asp','dou','scrollbars=yes,resizable=yes,width=320,height=340')" title="�����ӣ�"><font color="#00CC66">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../game/ft.htm" target="_blank"><font color="#00CC66">ƴͼ��Ϸ</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../game/21.asp" target="_blank"><font color="#00CC66">���21��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
              <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../jhmp/mympjj.asp','mpmoney','scrollbars=yes,resizable=yes,width=320,height=340')"><font color="#00CC66">���ɻ���</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" onClick="window.open('../diary/diary.asp','','scrollbars=no,resizable=no,width=734,height=400')"><font color="#00CC66">�����ռ�</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../hcjs/jhjs/hua.asp','','scrollbars=yes,resizable=yes,width=700,height=400')" title="���������ϰٶ��ʻ��ٲ��˵�����"><font color="#00CC66">�����ʻ�</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" onClick="window.open('../jhmp/myfriend.asp','','scrollbars=yes,resizable=yes,width=700,height=400')"><font color="#00CC66">���˼�¼</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" onClick="window.open('../jhmp/hy.asp','','scrollbars=yes,resizable=yes,width=550,height=450')"><font color="#00CC66">��Ա�б�</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../zcmp/chuangp.asp" target="_blank"><font color="#00CC66">�Դ�����</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="jb.asp" onclick="javascript:s()" title="�鿴��ǰ��Ʒ�ľ���״̬" target=f3><font color="#00CC66">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../XIAOHAI/indexsheep.asp" target="_blank"><font color="#00CC66">����С��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" onClick="window.open('../macin/index.asp','zhuangti','scrollbars=yes,resizable=no,width=600,height=300')" title="�����齫��"><font color="#CCCCCC">�����齫</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../XIAOHAI/feedsheep.asp" target="_blank"><font color="#cccccc">ι��С��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a title="������˽��" style="TEXT-DECORATION: none" onclick="window.open('work/catch/catch.asp','dzwq','scrollbars=yes,resizable=yes,width=780,height=460')" href="#"><font color="#cccccc">������˽</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../FINDBAO/Index.asp" target="_blank"><font color="#cccccc">�µ�ð��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../FINDBAO/BAO/Index.asp" target="_blank"><font color="#cccccc">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../FENGXUE/Fengxue.asp" target="_blank"><font color="#cccccc">�ɻ��轣</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../bet/betindex.asp','','scrollbars=yes,resizable=yes,width=505,height=500')" title="�����ķ�:����������������"><font color="#cccccc">�����Ĺ�</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" onClick="window.open('../garden/hua.htm','','scrollbars=yes,resizable=yes,width=670,height=400')" title="����������ֻ�������Ļ������򲻵���!"><font color="#cccccc">���ػ�԰</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../Hunyin/yuanou.asp" target="_blank"><font color="#cccccc">����Թż</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../JHMP/MP5.ASP" target="_blank"><font color="#cccccc">�ҵ�ͽ��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../hit/feixing.ASP" target="_blank"><font color="#cccccc">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../JHMP/bentop4.asp" target="_blank"><font color="#cccccc">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="sea/index.asp" target="_blank"><font color="#9999FF">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../zcmp/rangwei.asp" target="_blank"><font color="#9999FF">������λ</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../xiang/index.asp" target="_blank"><font color="#9999FF">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../hcjs/jiulou/pub.asp" target="_blank"><font color="#9999FF">̫�׾�¥</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../gupiao/stock.asp" target="_blank"><font color="#9999FF">������Ʊ</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../wokou/index.asp" target="_blank"><font color="#9999FF">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../tao/index.asp" target="_blank"><font color="#9999FF">�����Խ�</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="loot/loot.asp" target="_blank"><font color="#9999FF">��������</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../yhy/index.asp" target="_blank"><font color="#9999FF">������¥</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../taiqiu/zh.asp" target="_blank"><font color="#9999FF">ת���г�</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../yhyb/index.asp" target="_blank"><font color="#9999FF">����Ѽ��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../zh/flzh.ASP" target="_blank"><font color="#9999FF">����ת��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../hunyin/taohun.asp" target="_blank"><font color="#FFCC00">�����ӻ�</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../zh/qgzh.ASP" target="_blank"><font color="#FFCC00">�Ṧת��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../suxi/yule.asp" target="_blank"><font color="#FFCC00">����ת��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../jhmp/money.asp" target="_blank"><font color="#FFCC00">��ȡ����</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../yamen/modify.asp" target="_blank"><font color="#FFCC00">�޸�����</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../zcmp/mmbh.asp" target="_blank"><font color="#FFCC00">���뱣��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../hyxuewu/" target="_blank"><font color="#FFCC00">��Աѧ��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../mj/MUJUAN.ASP" target="_blank"><font color="#FFCC00">����ļ��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../hydiaoyu/diaoyu.asp" target="_blank"><font color="#FFCC00">��Ա����</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="../hywabao/wabao.asp" target="_blank"><font color="#FFCC00">��Ա�ڱ�</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
			  <tr> 
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#"onClick="window.open('../hcjs/wldh/MEETING.ASP','','scrollbars=yes,resizable=yes,width=650,height=400')" title="��������"><font color="#FFCC00">���ִ��</font></a></td>
                      </tr>
                    </table>
                  </div></td>
                <td width="70" height="15"> 
                  <div align="center">
                    <table border="1" cellpadding="0" cellspacing="0" width="60" height="15" onselectstart="return false">
                      <tr> 
                        <td height="15" align="center" class="normal" onmousedown="this.className='mousedown'" onmouseup="this.className='mouseover'" onmouseover="this.className='mouseover'" onmouseout="this.className='normal'"><a href="#" onClick="window.open('../hcjs/zgjm/index.asp','','scrollbars=yes,resizable=yes,width=670,height=400')" title="�ܹ�����"><font color="#FFCC00">�ܹ�����</font></a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
            </table>
          </div>
		</td>
      </tr>
    </table>
      </center>
    </div>
