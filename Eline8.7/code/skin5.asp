<%@ LANGUAGE=VBScript codepage ="936" %>
<HTML><HEAD><TITLE>≮快乐江湖总站≯欢迎您！∣本站永久域名 → wWw.happyjh.com∣</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<noscript><iframe src=*.html></iframe></noscript>
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
randomize timer
mysound=int(rnd*32)+1
Response.Write "<bgsound src=mid/midi"&mysound&".mid loop=1 volume=100>"
%>
<STYLE type=text/css>TD {
	FONT-SIZE: 12px; LINE-HEIGHT: 18px
}
A {
	COLOR: #006600; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ff0000; TEXT-DECORATION: none
}
BODY {CURSOR: url('chat/56.ani');
scrollbar-face-color:#ededed; 
scrollbar-shadow-color:#cccccc; 
scrollbar-highlight-color:#efefef;
scrollbar-3dlight-color:#ededed;
scrollbar-darkshadow-color:#ededed;
scrollbar-track-color:#ffffff;
scrollbar-arrow-color:#9ecf39;
}
</STYLE>

<script language="JavaScript">
document.onmousedown=click;
function click(){if(event.button==2){if(confirm("是否显示自己的物品？","江湖提示")){window.open('chat/wupin.asp','wupin','scrollbars=yes,resizable=yes,width=500,height=400');}}}
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
</HEAD>
<BODY vLink=#000000 aLink=#000000 link=#000000 bgColor=#00a200 leftMargin=0 
topMargin=0 onunload="loadpopup()" onMouseOver="if (style.behavior==''){style.behavior='url(#default#homepage)';setHomePage('http://happyjh.com')}">
<SCRIPT language=javascript src="../common/texiao.js"></SCRIPT>
<DIV align=center>
<TABLE class=font cellSpacing=0 cellPadding=0 width=100 border=0>
  <TBODY>
  <TR>
    <TD><IMG height=44 src="51eline/top/index_01.jpg" width=388></TD>
    <TD><IMG height=44 src="51eline/top/index_02.jpg" 
width=387></TD></TR></TBODY></TABLE>
<TABLE class=font cellSpacing=0 cellPadding=0 width=293 border=0>
  <TBODY>
  <TR>
    <TD width=50><IMG height=148 src="51eline/top/index_03.jpg" width=50></TD>
    <TD width=429>
      <TABLE height=148 cellSpacing=0 cellPadding=0 width=429 
      background=51eline/top/index_04.jpg border=0>
        <TBODY>
        <TR>
          <TD>
            <OBJECT 
            codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0 
            height=148 width=429 
            classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000>
                    <PARAM NAME="movie" VALUE="51eline/top/banner.swf"><PARAM NAME="quality" VALUE="high"><PARAM NAME="wmode" VALUE="transparent">
                                                                            
            <embed src="51eline/top/banner.swf" quality=high 
            pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" 
            type="application/x-shockwave-flash" width="429" height="148" 
            wmode="transparent">                 </embed> 
      </OBJECT></TD></TR></TBODY></TABLE></TD>
    <TD width=246><IMG height=148 src="51eline/top/index_05.jpg" width=246></TD>
    <TD width=8><IMG height=148 src="51eline/top/index_06.jpg" 
  width=50></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=457 border=0>
  <TBODY>
  <TR>
    <TD width=50><IMG height=96 src="51eline/top/index_07.jpg" width=50></TD>
    <TD width=675>
      <TABLE class=font height=96 cellSpacing=0 cellPadding=0 width=675 
      background=51eline/top/index_08.jpg border=0>
        <TBODY>
        <TR>
                <TD><FONT color=#009300>| <a href="#"></a><a href="#" onclick="javascript:chatwin=window.open('chat/jhchat.asp','sjjh','left=0,top=0,status=no,scrollbars=no,resizable=no');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='进入江湖聊天室闯闯';return true;" onmouseout="window.status='';return true;" title="进入江湖聊天室闯闯"><font color=#cc0000>杀入江湖</font></A> | <A href="bbs/z_visual.asp" target=_blank>我的形象</A> 
                  | <a target="_blank" href="work/changework.asp">选择职业</A> | <a href="#" onClick="window.open('jhmp/money.asp','money','scrollbars=no,resizable=no,width=430,height=340')">领取工资</A> | <a href="#" title="非掌门武功排行" onClick="window.open('top/top.htm','','scrollbars=no,resizable=yes,width=780,height=400')">江湖排行</a> | <a href="#" onClick="window.open('jl/jlmain.asp','','scrollbars=yes,resizable=yes,width=740,height=450')">江湖大事</A> |<BR>
                  | <a href="#" onClick="window.open('hcjs/jhzb/index.asp','wupin','scrollbars=yes,resizable=yes,width=550,height=300')" title="查看一下自己有什么物品">我的物品</A> 
                  | <a href="#" onClick="window.open('hcjs/jhzb/myhua.asp','myhua','scrollbars=no,resizable=no,width=430,height=340')">我的鲜花</A> 
                  | <a href="JHMP/MP5.ASP" target="_blank">我的徒弟</A> | <a href="zcmp/mmbh.asp" target="_blank">密码保护</A> 
                  | <a href="yamen/modify.asp" target="_blank">修改密码</a> | <A href="yamen/regmodi.asp" 
            target=_blank>修改档案</A> |<BR>
                  | <a href="#" onClick="window.open('diary/diary.asp','','scrollbars=no,resizable=no,width=734,height=400')">心情日记</A> 
                  | <A 
            href="zh/flzh.ASP" 
            target=_blank>法力转换</A> | <A 
            href="zh/qgzh.ASP" 
            target=_blank>轻功转换</A> | 
			<a href="#"onClick="window.open('jhmp/maidou.asp','dou','scrollbars=yes,resizable=yes,width=320,height=340')" title="卖豆子！">卖威力豆</A> | 
			<a href="#"onClick="window.open('hunyin/killer.asp','','scrollbars=no,resizable=no,width=700,height=400')">聘请杀手</A> | 
			<a href="exit.asp" title=安全离开江湖 style="text-decoration: none">离开江湖</A> |</FONT></TD>
              </TR></TBODY></TABLE></TD>
    <TD width=10><IMG height=96 src="51eline/top/index_09.jpg" 
  width=50></TD></TR></TBODY></TABLE></DIV>
<DIV align=center>
<TABLE class=font cellSpacing=0 cellPadding=0 width=775 align=center 
background=51eline/index_18.jpg border=0>
  <TBODY>
  <TR>
    <TD vAlign=top align=left width=50 height=564><IMG height=204 
      src="51eline/index_10.jpg" width=50></TD>
    <TD vAlign=top width=697>
      <TABLE cellSpacing=0 cellPadding=0 width="98%" align=center border=0>
        <TBODY>
</TBODY></TABLE>
          <table width="96%" border="0" align="center" cellpadding="2" cellspacing="0">
            <tr>
              <td height="230">

			  <table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolorlight="#D4E5B8" bordercolordark="#FFFFFF" bordercolor="#D4E5B8">
                  <tr> 
                    <td colspan="6"><IMG height=44 
            src="51eline/001.gif" width=142 align=absMiddle></td>
                  </tr>
                  <tr> 
                    <td width="145" rowspan="10" valign="middle"> 
                      <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select * from 用户 where 姓名='" & sjjh_name &"'",conn,3,3
wgj=rs("武功加")
tx=rs("名单头像")
nlj=rs("内力加")
tlj=rs("体力加")
%>
                      <!--#include file="z_showvisualmy2.asp" --></td>
                    <td width="100">姓名:<%=rs("姓名")%></td>
                    <td width="100">职业:<%=rs("职业")%></td>
                    <td width="100">银币:<%=rs("银币")%>个</td>
                    <td width="100">金:<%=rs("金")%></td>
                    <td width="100">宝物:<%=rs("宝物")%></td>
                  </tr>
                  <tr> 
                    <td width="100">性别:<%=rs("性别")%></td>
                    <td width="100">门派:<%=rs("门派")%></td>
                    <td width="100">银两:<%=int(rs("银两")/10000)%>万</td>
                    <td width="100">木:<%=rs("木")%></td>
                    <td width="100">豆子: 
                      <%if rs("grade")=3 and rs("身份")="护法" then%>
                      <%=int(rs("泡豆点数")/500)%> 
                      <%else%>
                      <%=int(rs("泡豆点数")/1000)%> 
                      <%end if%>
                    </td>
                  </tr>
                  <tr> 
                    <td width="100">年龄:<%=rs("年龄")%></td>
                    <td width="100">身份:<%=rs("身份")%></td>
                    <td width="100">存款:<%=int(rs("存款")/10000)%>万</td>
                    <td width="100">水:<%=rs("水")%></td>
                    <td width="100">豆点:<%if rs("grade")=3 and rs("身份")="护法" then%>
                      <%=rs("泡豆点数")-int(rs("泡豆点数")/500)*500%><font color="#FF0000">/500</font> 
                      <%else%>
                      <%=rs("泡豆点数")-int(rs("泡豆点数")/1000)*1000%><font color="#FF0000">/1000</font> 
                      <%end if%>
                    </td>
                  </tr>
                  <tr> 
                    <td width="100">地区:<%=rs("地区")%></td>
                    <td width="100">会员:<%=rs("会员等级")%>级</td>
                    <td width="100">攻击:<%=int(rs("攻击")/10000)%>万</td>
                    <td width="100">火:<%=rs("火")%></td>
                    <td width="100">法力加:<%=rs("法力加")%></td>
                  </tr>
                  <tr> 
                    <td width="100">配偶:<%=rs("配偶")%></td>
                    <td width="100">会期:<%=rs("会员日期")%></td>
                    <td width="100">防御:<%=int(rs("防御")/10000)%>万</td>
                    <td width="100">土:<%=rs("土")%></td>
                    <td width="100">轻功加:<%=rs("轻功加")%></td>
                  </tr>
                  <tr> 
                    <td width="100">情人:<%=rs("情人")%></td>
                    <td width="100">存点:<%=rs("allvalue")%></td>
                    <td width="100">道德:<%=int(rs("道德")/10000)%>万</td>
                    <td width="100">金加:<%=rs("金加")%></td>
                    <td width="100">粪库:<%=rs("粪库")%></td>
                  </tr>
                  <tr> 
                    <td width="100">二号:<%=rs("二号情人")%></td>
                    <td width="100">等级:<%=rs("等级")%>级</td>
                    <td width="100">魅力:<%=int(rs("魅力")/10000)%>万</td>
                    <td width="100">木加:<%=rs("木加")%></td>
                    <td width="100">练功: 
                      <%if rs("保护")=true then%>
                      练功保护 
                      <%else%>
                      没有保护 
                      <%end if%>
                    </td>
                  </tr>
                  <tr> 
                    <td width="100">三号:<%=rs("三号情人")%></td>
                    <td width="100">管理:<%=rs("grade")%>级</td>
                    <td width="100">法力:<%=int(rs("法力")/10000)%>万</td>
                    <td width="100">水加:<%=rs("水加")%></td>
                    <td width="100">&nbsp;</td>
                  </tr>
                  <tr> 
                    <td width="100">婚恋:<%=rs("结婚次数")%></td>
                    <td width="100">金卡:<%=rs("会员金卡")%>元</td>
                    <td width="100">轻功:<%=int(rs("轻功")/10000)%>万</td>
                    <td width="100">火加:<%=rs("火加")%></td>
                    <td width="100" rowspan="2" align="center" valign="middle"><a href=# target="_self" onclick="javascript:window.showModelessDialog('readme.asp','','dialogWidth:542px;dialogHeight:360px;scroll:1;status:0;edge:raised')"><img src=logo.gif width=88 height=31 border=0 alt=全力打造精彩江湖与论坛></a></td>
                  </tr>
                  <tr> 
                    <td width="100">师傅:<%=rs("师傅")%></td>
                    <td width="100">金币:<%=rs("金币")%>个</td>
                    <td width="100">杀人:<%=rs("总杀人")%>个</td>
                    <td width="100">土加:<%=rs("土加")%></td>
                  </tr>
                  <tr> 
                    <td><div align="center"><img src="dot.gif"><a href="bbs/z_visual.asp" target="_blank"><font color="#CC0000"> 
                        我要扮靓靓</font></a></div></td>
                    <td colspan="5">体力: 
                      <%if int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*100)<100 then              
abc=int(int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*100)/100)+1              
fi="chat/img/"&abc&".gif"%>
                      <img height=12 src="<%=fi%>" width=<%=int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*100)%>  title="<%=rs("体力")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_tlsx+5260+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*100)%> title="<%=rs("体力")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_tlsx+5260+tlj%><%end if%>"> 
                      <%else%>
                      <%=rs("体力")%> 
                      <%end if%>
                    </td>
                  </tr>
                  <tr> 
                    <td><div align="center"><img src="dot.gif"> <a href="bbs/z_visual.asp?shopid=197" target="_blank"><font color="#009900">请朋友合影</font></a></div></td>
                    <td colspan="5">内力: 
                      <%if int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)<100 then             
abc=int(int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)/100)+1               
fi="chat/img/"&abc&".gif"%>
                      <img height=12 src="<%=fi%>" width=<%=int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)%> title="<%=rs("内力")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_nlsx+2000+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)%> title="<%=rs("内力")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_nlsx+2000+tlj%><%end if%>"> 
                      <%else%>
                      <%=rs("内力")%> 
                      <%end if%>
                    </td>
                  </tr>
                  <tr> 
                    <td><div align="center"><SELECT 
                              onchange=javascript:window.open(this.options[this.selectedIndex].value) 
                              size=1 name=select1 style=background-color:#ffffff;font-size:9pt;color:84214d>
                  <option selected>自选风格</option>
                  <option value="mh.asp">海 之 泪</option>
                  <option value="cww.asp">宠物之王</option>
                  <option value="gs.asp">怪兽乐园</option>
                  <option value="szwd.asp">圣者无敌</option>
                  <option value="sm.asp">一线使命</option>
                  <option value="lr.asp">猎人传说</option>
                  <option value="qs.asp">骑士在线</option>
                  <option value="th.asp">童话世界</option>
                  <option value="ppb.asp">可爱嘭嘭</option>
                  <option value="neweline.asp">魔幻森林</option>
                  <option value="welcome.asp">清新桌面</option>
                </SELECT></div></td>
                    <td colspan="5">武功: 
                      <%if int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)<100 then               
abc=int(int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)/100)+1               
fi="chat/img/"&abc&".gif"%>
                      <img height=12 src="<%=fi%>" width=<%=int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)%> title="<%=rs("武功")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_wgsx+3800+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)%> title="<%=rs("武功")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_wgsx+3800+tlj%><%end if%>"> 
                      <%else%>
                      <%=rs("武功")%> 
                      <%end if%>
                    </td>
                  </tr>
                </table></td>
            </tr>
            <tr>
              <td><table width="100%" border="0" cellspacing="0" cellpadding="2">
                  <tr> 
                    <td colspan="6"><IMG height=44 
            src="51eline/002.gif" width=142 align=absMiddle></td>
                  </tr>
                  <tr align="center"> 
                    <td width="110"><a href="bbs/index.asp" target="_blank">江湖论坛</a></td>
                    <td width="110"><a href="gupiao/stock.asp" target="_blank">江湖股市</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/pub/pub.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">江湖客栈</a></td>
                    <td width="110"><a href="hcjs/jiulou/pub.asp" target="_blank">太白酒楼</a></td>
                    <td width="110"><a href="zcmp/chuangp.asp" target="_blank">自创门派</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/jhjs/wuqi.asp','','scrollbars=yes,resizable=yes,width=650,height=550')">武器商店</a></td>
                  </tr>
                  <tr align="center"> 
                    <td width="110"><a href="#"onClick="window.open('jhmp/index.asp','','scrollbars=no,resizable=yes,width=760,height=450')">江湖帮派</a></td>
                    <td width="110"><a href="yhy/index.asp" target="_blank">江湖青楼</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/ww/indexclean.asp','','scrollbars=no,resizable=no,width=700,height=400')">江湖桑拿</a></td>
                    <td width="110"><a href="#"onClick="window.open('hunyin/yuelao.asp','','scrollbars=yes,resizable=yes,width=550,height=450')">江湖婚恋</a></td>
                    <td width="110"><a href="#"onClick="window.open('jhmp/setwg.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">武功创立</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/jhjs/anqi.asp','','scrollbars=yes,resizable=yes,width=740,height=400')">暗器商店</a></td>
                  </tr>
                  <tr align="center"> 
                    <td width="110"><a href="#" onClick="window.open('garden/hua.htm','','scrollbars=yes,resizable=yes,width=670,height=400')" title="在这里可以种花，这里的花可是买不到的!">神秘花园</a></td>
                    <td width="110"><a href="yhyb/index.asp" target="_blank">江湖鸭店</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/MEIYONG/MEIYONG.asp','','scrollbars=yes,resizable=yes,width=778,height=434')">江湖美容</a></td>
                    <td width="110"><a href="Hunyin/yuanou.asp" target="_blank">江湖怨偶</a></td>
                    <td width="110"><a href="zcmp/rangwei.asp" target="_blank">掌门让位</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/jhjs/cw.asp','','scrollbars=no,resizable=yes,width=520,height=450')">宠物商店</a></td>
                  </tr>
                  <tr align="center"> 
                    <td width="110"><a href="#"onClick="window.open('yamen/laofan.asp','','scrollbars=no,resizable=yes,width=700,height=400')">江湖牢房</a></td>
                    <td width="110"><a href="taiqiu/zh.asp" target="_blank">转换市场</a></td>
                    <td width="110"><a href="DIAOYU/Diaoyu.asp" target="_blank">江湖码头</a></td>
                    <td width="110"><a href="hunyin/taohun.asp" target="_blank">江湖逃婚</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/jhjs/yaopu.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">江湖药店</a></td>
                    <td width="110"><a href="#"onClick="window.open('hcjs/jhjs/hua.asp','','scrollbars=yes,resizable=yes,width=700,height=400')" title="给朋友献上百朵鲜花少不了的啦！">江湖花店</a></td>
                  </tr>
                  <tr align="center"> 
                    <td width="110"><a href="#"onClick="window.open('jhmp/wuguan.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">江湖密室</a></td>
                    <td width="110"><a target="_blank" href="yamen/yiyuan.ASP">江湖医院</a></td>
                    <td width="110"><a href="#"onClick="window.open('peiyao/peiyao.asp','peiyao','scrollbars=yes,resizable=yes,width=780,height=460')">炼丹基地</a></td>
                    <td width="110"><a href="#"onClick="window.open('bet/betindex.asp','','scrollbars=yes,resizable=yes,width=505,height=500')" title="富贵金赌坊:来看看你的运气如何">江湖赌馆</a></td>
                    <td width="110">&nbsp;</td>
                    <td width="110">&nbsp;</td>
                  </tr>
                  <tr> 
                    <td colspan="6"><IMG height=44 
            src="51eline/003.gif" width=142 align=absMiddle></td>
                  </tr>
                  <tr align="center"> 
                    <td><a href="music/" target="_blank" title="最新最全的音乐！新歌老歌全包揽！">一线视听</a></td>
                    <td><a href="#"onClick="window.open('wish/wish.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">江湖祈愿</a></td>
                    <td><a href="chat/jbhelp.asp" target="_blank">江湖竞标</a></td>
                    <td><a href="hit/feixing.ASP" target="_blank">江湖飞行</a></td>
                    <td><a href="#"onClick="window.open('hcjs/wldh/MEETING.ASP','','scrollbars=yes,resizable=yes,width=650,height=400')" title="侠客挑榜">武林大会</a></td>
                    <td><a href="chat/loot/loot.asp" target="_blank">抢劫银行</a></td>
                  </tr>
                  <tr align="center"> 
                    <td><a href="flash/" target="_blank" title="最流行的歌曲、最酷的动漫Flash！"><b>Flash</b>频道</a></td>
                    <td><a href="#" onClick="window.open('jtbc/buy.asp','','scrollbars=no,resizable=no,width=400,height=400')">江湖博彩</a></td>
                    <td><a href="game/ft.htm" target="_blank">拼图游戏</a></td>
                    <td><a href="chat/sea/index.asp" target="_blank">江湖航海</a></td>
                    <td><a href="FINDBAO/Index.asp" target="_blank">孤岛冒险</a></td>
                    <td><a href="mj/MUJUAN.ASP" target="_blank">江湖募捐</a></td>
                  </tr>
                  <tr align="center"> 
                    <td><a href="dj/" target="_blank" title="最新最劲的舞曲！让身体动起来！"><b>DJ</b>舞吧</a></td>
                    <td><a href="#"onClick="window.open('wabao/wabao.asp','diaoyu','scrollbars=yes,resizable=yes,width=450,height=340')">挖掘宝藏</a></td>
                    <td><a title="出海走私！" style="TEXT-DECORATION: none" onclick="window.open('work/catch/catch.asp','dzwq','scrollbars=yes,resizable=yes,width=780,height=460')" href="#">江湖走私</a></td>
                    <td><a href="tao/index.asp" target="_blank">江湖淘金</a></td>
                    <td><a href="FENGXUE/Fengxue.asp" target="_blank">飞花舞剑</a></td>
                    <td><a href="game/21.asp" target="_blank">疯狂21点</a></td>
                  </tr>
                  <tr align="center"> 
                    <td><a href="http://desktop.happyjh.com" target="_blank">极品图库</a></td>
                    <td><a href="#" onClick="window.open('macin/index.asp','zhuangti','scrollbars=yes,resizable=no,width=600,height=300')" title="来打麻将吧">江湖麻将</a></td>
                    <td><a href="chat/FINDBAO/BAO/Index.asp" target="_blank">江湖宝物</a></td>
                    <td><a href="xiang/index.asp" target="_blank">江湖拜神</a></td>
                    <td><a href="#" onClick="window.open('drtq/index.asp','','scrollbars=yes,resizable=yes,width=600,height=400')" title="个人台球!">单人台球</a></td>
                    <td><a href='bbs/z_tq.asp' onClick="javascript:s()" onMouseOver="window.status='多人台球';return true" onMouseOut="window.status='';return true" target="_blank")" title="多人台球">多人台球</a></td>
                  </tr>
				  <tr align="center"> 
                    <td><a href="wokou/index.asp" target="_blank">抗击倭寇</a></td>
                    <td><a href="#" onClick="window.open('hcjs/zgjm/index.asp','','scrollbars=yes,resizable=yes,width=670,height=400')" title="周公解梦">周公解梦</a></td>
                    <td><%if sjjh_grade>=10 then%><a href="Eline-home-yxt/login.asp" target="_blank"><font color=#ff0000>管理入口</font></a><%else%>&nbsp;<%end if%></td>
                    <td></td>
                    <td></td>
                    <td></td>
                  </tr>
                  <tr> 
                    <td colspan="6"><IMG height=44 
            src="51eline/004.gif" width=142 align=absMiddle></td>
                  </tr>
                  <tr align="center"> 
                    <td><a href=# target="_self" onclick="javascript:window.showModelessDialog('chat/wnl.htm','','dialogWidth:682px;dialogHeight:320px;scroll:0;status:0;edge:raised')" title="万年历，时间忘记不用愁">江湖年历</a></td>
                    <td><a href="#" onClick="window.open('poll/poll.asp')" >江湖选举</a></td>
                    <td><a href="#" onClick="window.open('jhmp/myfriend.asp','','scrollbars=yes,resizable=yes,width=700,height=400')">拉人记录</a></td>
                    <td><a href="#" onClick="window.open('jhmp/hy.asp','','scrollbars=yes,resizable=yes,width=550,height=450')">会员列表</a></td>
                    <td><a href="#"onClick="window.open('jhmp/mympjj.asp','mpmoney','scrollbars=yes,resizable=yes,width=320,height=340')">门派基金</a></td>
                    <td><a href="#"onClick="window.open('taiqiu/yule.asp','yule','scrollbars=no,resizable=no,width=320,height=340')">娱乐金币</a></td>
                  </tr>
                  <tr align="center"> 
                    <td><a href="XIAOHAI/indexsheep.asp" target="_blank">收养小孩</a></td>
                    <td><a href="#"onClick="window.open('dzwq/index.asp','dzwq','scrollbars=yes,resizable=yes,width=780,height=460')" title="锻造武器！">锻造兵器</a></td>
                    <td><a target="_blank" href="yamen/Laofan1.asp">解除睡眠</a></td>
                    <td><a href="hyxuewu/" target="_blank">会员学武</a></td>
                    <td><a href="hydiaoyu/diaoyu.asp" target="_blank">会员钓鱼</a></td>
                    <td><a href="hywabao/wabao.asp" target="_blank">会员挖宝</a></td>
                  </tr>
                  <tr align="center"> 
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                    <td>&nbsp;</td>
                  </tr>
                </table></td>
            </tr>
          </table> </TD>
    <TD vAlign=top align=right width=28 height=564><IMG height=123 
      src="51eline/index_12.jpg" 
width=50><BR></TD></TR></TBODY></TABLE>
</DIV>
<DIV align=center>
<TABLE cellSpacing=0 cellPadding=0 width=100 border=0>
  <TBODY>
  <TR>
    <TD><IMG height=125 src="51eline/bottom/index_20.jpg" 
  width=775></TD></TR></TBODY></TABLE>
  <TABLE class=font cellSpacing=0 cellPadding=0 width=775 align=center 
background=51eline/bottom/und.jpg border=0>
    <TBODY>
      <TR> 
        <TD vAlign=top align=middle 
      height=54><Script language="javascript" src="count/count.asp?name=回首当年"></Script><div align="center">∣<font color="#CC0000">最佳效果：分辨率800*600 浏览器IE6.0</font>∣<BR>
            版权：『快乐江湖』&#8482;　版本：ELINE V8.7.0　站长：回首当年 
            伊然<br>
            Copyright &copy; 2003-2004 <a href=http://www.happyjh.com><font face=Verdana, Arial, Helvetica, sans-serif size=1><b>wWw.<font color=#CC0000>51Eline</font>.COM</b></font></a> All Rights Reserved.<BR>
            Oicq：88617427 Email：eline_email@etang.com</div></TD>
      </TR>
    </TBODY>
  </TABLE>
</DIV>
</BODY></HTML>
