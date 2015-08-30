<%@ LANGUAGE=VBScript codepage ="936" %>
<HTML><HEAD><TITLE>≮快乐江湖总站≯欢迎您！∣本站永久域名 → wWw.happyjh.com∣</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<noscript><iframe src=*.html></iframe></noscript>
<LINK href="51eline/ppb/css.css" type=text/css rel=stylesheet>
<META content="MSHTML 6.00.2600.0" name=GENERATOR>
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
BODY {CURSOR: url('51eline/ppb/mouse.ani');
scrollbar-face-color:#ededed; 
scrollbar-shadow-color:#cccccc; 
scrollbar-highlight-color:#efefef;
scrollbar-3dlight-color:#ededed;
scrollbar-darkshadow-color:#ededed;
scrollbar-track-color:#ffffff;
scrollbar-arrow-color:#0186e3;
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
<BODY text=#000000 bgColor=#ffffff leftMargin=0 
background=51eline/ppb/bg.gif topMargin=0 onunload="loadpopup()" onMouseOver="if (style.behavior==''){style.behavior='url(#default#homepage)';setHomePage('http://happyjh.com')}">
<SCRIPT language=javascript src="../common/texiao.js"></SCRIPT>
<DIV align=center>
  <TABLE cellSpacing=0 cellPadding=0 width=760 border=0>
    <TBODY>
  <TR vAlign=top>
    <TD width=278><IMG height=94 src="51eline/ppb/top00.jpg" width=278></TD>
        <TD width=100 background="51eline/ppb/top01.jpg"><OBJECT 
      codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0 
      height=94 width=208 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000>
            <PARAM NAME="_cx" VALUE="19579"><PARAM NAME="_cy" VALUE="4207"><PARAM NAME="FlashVars" VALUE="-1"><PARAM NAME="Movie" VALUE="51eline/ppb/5xin.swf"><PARAM NAME="Src" VALUE="51eline/ppb/5xin.swf">
	  <PARAM NAME="WMode" VALUE="Transparent"><PARAM NAME="Play" VALUE="-1"><PARAM NAME="Loop" VALUE="-1"><PARAM NAME="Quality" VALUE="High"><PARAM NAME="SAlign" VALUE=""><PARAM NAME="Menu" VALUE="0"><PARAM NAME="Base" VALUE=""><PARAM NAME="AllowScriptAccess" VALUE="always"><PARAM NAME="Scale" VALUE="ShowAll"><PARAM NAME="DeviceFont" VALUE="0"><PARAM NAME="EmbedMovie" VALUE="0"><PARAM NAME="BGColor" VALUE=""><PARAM NAME="SWRemote" VALUE="">                                                                       
        <embed src="51eline/ppb/5xin.swf" width="208"       height="94" quality="high"       
      pluginspage="http://www.macromedia.com/go/getflashplayer"       
      type="application/x-shockwave-flash" wmode="transparent"       
      menu="false">        </embed></OBJECT></TD>
    <TD width=382><IMG height=94 src="51eline/ppb/top02.jpg" 
  width=274></TD></TR>
  <TR vAlign=top>
    <TD width=278><IMG height=83 src="51eline/ppb/top03.jpg" width=278></TD>
    <TD width=100><IMG height=83 src="51eline/ppb/top04.jpg" width=208></TD>
    <TD width=382><IMG height=83 src="51eline/ppb/top05.jpg" 
  width=274></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=760 border=0>
  <TBODY>
  <TR>
    <TD width=760 background=51eline/ppb/button.gif height=19>
      <DIV align=center>
      <TABLE height=27 cellSpacing=0 cellPadding=0 width=600 align=center 
      border=0>
        <TBODY>
        <TR>
          <TD>
            <DIV align=center><a href="#" onclick="javascript:chatwin=window.open('chat/jhchat.asp','sjjh','left=0,top=0,status=no,scrollbars=no,resizable=no');chatwin.resizeTo(screen.availWidth,screen.availHeight);" onmouseover="window.status='进入江湖聊天室闯闯';return true;" onmouseout="window.status='';return true;" title="进入江湖聊天室闯闯"><font color=#FF0000>杀入江湖</font></A> 
                      | <A href="bbs/z_visual.asp" target=_blank>我的形象</A> | <a target="_blank" href="work/changework.asp">选择职业</A> 
                      | <a href="#" onClick="window.open('jhmp/money.asp','money','scrollbars=no,resizable=no,width=430,height=340')">领取工资</A> 
                      | <a href="zcmp/mmbh.asp" target="_blank">密码保护</A> | <a href="yamen/modify.asp" target="_blank">修改密码</a> | <A href="yamen/regmodi.asp" 
            target=_blank>修改档案</A> | <a href="exit.asp" title=安全离开江湖 style="text-decoration: none">离开江湖</A></DIV>
			</TD></TR></TBODY></TABLE></DIV></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=760 border=0>
  <TBODY>
  <TR>
    <TD vAlign=top height=558>
      <TABLE cellSpacing=0 cellPadding=0 width=760 border=0>
            <TBODY>
        <TR vAlign=top>
          <TD width=45 height=604><IMG height=36 
            src="51eline/ppb/left.gif" width=45></TD>
                <TD width=671 background=51eline/ppb/tb_bg.gif> 
                  <TABLE cellSpacing=0 cellPadding=0 width=671 
            background=51eline/ppb/mid_bg.gif border=0>
                    <TBODY>
              <TR>
                        <TD vAlign=top width=224> 
                          <TABLE cellSpacing=0 cellPadding=0 width=224 border=0>
                    <TBODY>
                    <TR>
                      <TD><IMG height=76 src="51eline/ppb/sys_top.gif" 
                        width=224></TD></TR>
                    <TR>
                      <TD height=2>
                        <TABLE cellSpacing=0 cellPadding=0 width=224 border=0>
                          <TBODY>
                          <TR vAlign=top>
                            <TD width=33 background=51eline/ppb/sys_left.gif 
                            height=2>&nbsp;</TD>
                            <TD width=148 background=51eline/ppb/sys_bg.gif 
                            height=2>
                              <DIV align=center>
                                            <TABLE cellSpacing=2 cellPadding=2 width=148 
                              border=0>
                                              <TBODY>
                                                <TR> 
                                                  <TD>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select * from 用户 where 姓名='" & sjjh_name &"'",conn,3,3
wgj=rs("武功加")
tx=rs("名单头像")
nlj=rs("内力加")
tlj=rs("体力加")
%>
<!--#include file="z_showvisualmy3.asp" --></TD>
                                                </TR>
                                                <TR> 
                                                  <TD align="center"><a href="bbs/z_visual.asp" target="_blank"><font color="#CC0000"> 
                                                    我要扮靓靓</font></a></TD>
                                                </TR>
                                                <TR> 
                                                  <TD align="center"><a href="bbs/z_visual.asp?shopid=197" target="_blank">请朋友合影</a></TD>
                                                </TR>
				   <tr>
                    <td align="center"><a href="bbs/z_visual.asp?shopid=697" target="_blank">大家的照片</a></td>
                  </tr>
				  	<tr>
                    <td align="center"><a href="bbs/z_Visual_photo2.asp" target="_blank">参加秀比赛</a></td>
                  </tr>												
                                                <TR> 
                                                  <TD align="center"><SELECT 
                              onchange=javascript:window.open(this.options[this.selectedIndex].value) 
                              size=1 name=select1 style=background-color:#0094ec;font-size:9pt;color:ffffff>
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
                </SELECT></TD>
                                                </TR>
                                              </TBODY>
                                            </TABLE>
                                          </DIV></TD>
                            <TD width=43 
                            background=51eline/ppb/sys_right.gif 
                            height=2>&nbsp;</TD></TR></TBODY></TABLE></TD></TR>
                    <TR>
                      <TD><IMG height=19 src="51eline/ppb/sys_bottom.gif" 
                        width=224></TD></TR></TBODY></TABLE><BR>
                  <TABLE cellSpacing=0 cellPadding=0 width=224 border=0>
                    <TBODY>
                    <TR>
                      <TD><IMG height=76 src="51eline/ppb/wallpaper.jpg" 
                        width=224></TD></TR>
                    <TR>
                      <TD>
                        <TABLE cellSpacing=0 cellPadding=0 width=224 border=0>
                                    <TBODY>
                          <TR vAlign=top>
                            <TD width=33 background=51eline/ppb/sys_left.gif 
                            height=160>&nbsp;</TD>
                                        <TD width=148 
                            height=120 align="left" background=51eline/ppb/sys_bg.gif><br>
                                          开发公司：一线网络<BR>
                运营公司：快乐江湖总站<BR>
                江湖版本：ELINE_V8.7.0<BR>
                开通时间：2003-03-16<BR>
                官方网站：<A 
            href="http://www.happyjh.com" target=_blank>进入</A><BR>
                客服信箱：<A 
            href="mailto:eline_email@etang.com">进入</A></TD>
                            <TD width=43 
                            background=51eline/ppb/sys_right.gif 
                            height=160>&nbsp;</TD></TR></TBODY></TABLE></TD></TR>
                    <TR>
                      <TD><IMG height=19 src="51eline/ppb/sys_bottom.gif" 
                        width=224></TD></TR></TBODY></TABLE></TD>
                        <TD vAlign=top width=430> 
                          <TABLE width=430 border=0 cellPadding=0 cellSpacing=0 
                  background=51eline/ppb/down_bg.jpg>
                            <TBODY>
                    <TR>
                      <TD vAlign=top height=40><IMG height=85 
                        src="51eline/ppb/intro.gif" width=430></TD></TR>
                    <TR>
                      <TD vAlign=top height=75>
                        <TABLE cellSpacing=0 cellPadding=0 width=430 border=0>
                          <TBODY>
                          <TR vAlign=top>
                            <TD width=10 
                            background=51eline/ppb/download_left.gif 
                            height=95><IMG height=44 
                              src="51eline/ppb/download_left.gif" 
                            width=10></TD>
                                        <TD width=411 height=95>
			  <table width="100%" border="1" cellpadding="1" cellspacing="0" bordercolorlight="#0186e3" bordercolordark="#0a81d5" bordercolor="#0186e3">
                                            <tr> 
                                              <td>姓名:<%=rs("姓名")%></td>
                                              <td>职业:<%=rs("职业")%></td>
                                              <td>银币:<%=rs("银币")%>个</td>
                                              <td>金:<%=rs("金")%></td>
                                              <td>宝物:<%=rs("宝物")%></td>
                                            </tr>
                                            <tr> 
                                              <td>性别:<%=rs("性别")%></td>
                                              <td>门派:<%=rs("门派")%></td>
                                              <td>银两:<%=int(rs("银两")/10000)%>万</td>
                                              <td>木:<%=rs("木")%></td>
                                              <td>豆子: 
                                                <%if rs("grade")=3 and rs("身份")="护法" then%>
                                                <%=int(rs("泡豆点数")/500)%> 
                                                <%else%>
                                                <%=int(rs("泡豆点数")/1000)%> 
                                                <%end if%>
                                              </td>
                                            </tr>
                                            <tr> 
                                              <td>年龄:<%=rs("年龄")%></td>
                                              <td>身份:<%=rs("身份")%></td>
                                              <td>存款:<%=int(rs("存款")/10000)%>万</td>
                                              <td>水:<%=rs("水")%></td>
                                              <td>豆点:
                                                <%if rs("grade")=3 and rs("身份")="护法" then%>
                                                <%=rs("泡豆点数")-int(rs("泡豆点数")/500)*500%><font color="#FF0000">/500</font> 
                                                <%else%>
                                                <%=rs("泡豆点数")-int(rs("泡豆点数")/1000)*1000%><font color="#FF0000">/1000</font> 
                                                <%end if%>
                                              </td>
                                            </tr>
                                            <tr> 
                                              <td>地区:<%=rs("地区")%></td>
                                              <td>会员:<%=rs("会员等级")%>级</td>
                                              <td>攻击:<%=int(rs("攻击")/10000)%>万</td>
                                              <td>火:<%=rs("火")%></td>
                                              <td>法力加:<%=rs("法力加")%></td>
                                            </tr>
                                            <tr> 
                                              <td>配偶:<%=rs("配偶")%></td>
                                              <td>会期:<%=rs("会员日期")%></td>
                                              <td>防御:<%=int(rs("防御")/10000)%>万</td>
                                              <td>土:<%=rs("土")%></td>
                                              <td>轻功加:<%=rs("轻功加")%></td>
                                            </tr>
                                            <tr> 
                                              <td>情人:<%=rs("情人")%></td>
                                              <td>存点:<%=rs("allvalue")%></td>
                                              <td>道德:<%=int(rs("道德")/10000)%>万</td>
                                              <td>金加:<%=rs("金加")%></td>
                                              <td>粪库:<%=rs("粪库")%></td>
                                            </tr>
                                            <tr> 
                                              <td>二号:<%=rs("二号情人")%></td>
                                              <td>等级:<%=rs("等级")%>级</td>
                                              <td>魅力:<%=int(rs("魅力")/10000)%>万</td>
                                              <td>木加:<%=rs("木加")%></td>
                                              <td>练功: 
                                                <%if rs("保护")=true then%>
                                                保护 
                                                <%else%>
                                                无保护 
                                                <%end if%>
                                              </td>
                                            </tr>
                                            <tr> 
                                              <td>三号:<%=rs("三号情人")%></td>
                                              <td>管理:<%=rs("grade")%>级</td>
                                              <td>法力:<%=int(rs("法力")/10000)%>万</td>
                                              <td>水加:<%=rs("水加")%></td>
                                              <td>&nbsp;</td>
                                            </tr>
                                            <tr> 
                                              <td>婚恋:<%=rs("结婚次数")%></td>
                                              <td>金卡:<%=rs("会员金卡")%>元</td>
                                              <td>轻功:<%=int(rs("轻功")/10000)%>万</td>
                                              <td>火加:<%=rs("火加")%></td>
                                              <td rowspan="2" align="center" valign="middle">&nbsp;</td>
                                            </tr>
                                            <tr> 
                                              <td>师傅:<%=rs("师傅")%></td>
                                              <td>金币:<%=rs("金币")%>个</td>
                                              <td>杀人:<%=rs("总杀人")%>个</td>
                                              <td>土加:<%=rs("土加")%></td>
                                            </tr>
                                            <tr> 
                                              <td colspan="6"><div align="center"></div>
                                                体力: 
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
                                              <td colspan="6"><div align="center"></div>
                                                内力: 
                                                <%if int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)<100 then             
abc=int(int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)/100)+1               
fi="chat/img/"&abc&".gif"%>
                                                <img height=12 src="<%=fi%>" width=<%=int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)%> title="<%=rs("内力")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_nlsx+2000+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)%> title="<%=rs("内力")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_nlsx+2000+tlj%><%end if%>"> 
                                                <%else%>
                                                <%=rs("内力")%> 
                                                <%end if%>
                                              </td>
                                            </tr>
											<tr>													                      <td colspan="5">武功: 
                      <%if int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)<100 then               
abc=int(int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)/100)+1               
fi="chat/img/"&abc&".gif"%>
                      <img height=12 src="<%=fi%>" width=<%=int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)%> title="<%=rs("武功")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_wgsx+3800+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)%> title="<%=rs("武功")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_wgsx+3800+tlj%><%end if%>"> 
                      <%else%>
                      <%=rs("武功")%> 
                      <%end if%>
                    </td></tr>
                <TR> 
                  <form name="form2" method="get" action="">
<%
ip=request("ip")
if ip="" then
if session("sjjh_grade")=10 then
ip=rs("lastip")
else
ip=Request.ServerVariables("REMOTE_ADDR")
end if
end if
sip=split(ip,".")
if ubound(sip)<>3 then
ip=Request.ServerVariables("REMOTE_ADDR")
erase sip
sip=split(ip,".")
end if
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
rs.close
rs.open "select * FROM z WHERE a<=" & num & " and b>=" & num,conn,1,1
if rs.eof or rs.bof then
country="亚洲"
city="未知"
else
country=rs("c")
city=rs("d")
end if
if country="" then country="中国"
if city="" then city="未知"
last="地区:" & country & " 城市:" & city
rs.close
set rs=nothing
conn.close
set conn=nothing%>
                                                <TD colspan="6" align=left>查IP:
<input name="ip" type="text" id="ip" size="15" maxlength="15" style="color: #666666; background-color: #0099ff; text-decoration: blink; border: 1px solid #ffffff" value="<%=ip%>">
<input type="submit" name="Submit" value="查找" style="BACKGROUND-COLOR: #009933; BORDER-BOTTOM: #ffffff 1px solid; BORDER-LEFT: #ffffff 1px solid; BORDER-RIGHT: #ffffff 1px solid; BORDER-TOP: #ffffff 1px solid; CURSOR: hand; FONT-FAMILY: '宋体'; COLOR: #FFFFFF; FONT-SIZE: 9pt; HEIGHT: 18px">
来自：<font color=yellow><%=last%></font> </TD>
                  </form>
</TR>								
                                          </table> </TD>
                            <TD width=9 
                            background=51eline/ppb/download_right.gif 
                            height=95><IMG height=58 
                              src="51eline/ppb/download_right.gif" 
                            width=9></TD></TR></TBODY></TABLE></TD></TR>
                    <TR>
                      <TD vAlign=top><IMG height=20 
                        src="51eline/ppb/download_bottom.gif" 
                    width=430></TD></TR></TBODY></TABLE>
                  <TABLE cellSpacing=0 cellPadding=0 width=430 border=0>
                    <TBODY>
                    <TR>
                      <TD vAlign=top height=40><IMG height=85 
                        src="51eline/ppb/news_top.gif" width=430></TD></TR>
                    <TR>
                      <TD vAlign=top height=15>
                        <TABLE cellSpacing=0 cellPadding=0 width=430 border=0>
                          <TBODY>
                          <TR vAlign=top>
                            <TD width=10 
                            background=51eline/ppb/download_left.gif 
                            height=12><IMG height=44 
                              src="51eline/ppb/download_left.gif" 
                            width=10></TD>
                            <TD width=411 background=51eline/ppb/down_bg.jpg 
                            height=12>
                              <table width="100%" border="0" cellspacing="0" cellpadding="2">
                                            <tr> 
                                              <td colspan="6" bgcolor="#3399FF"> 
                                                <div align="center"><a href="#" title="非掌门武功排行" onClick="window.open('top/top.htm','','scrollbars=no,resizable=yes,width=780,height=400')">江湖排行</a> 
                                                  | <a href="#" onClick="window.open('jl/jlmain.asp','','scrollbars=yes,resizable=yes,width=740,height=450')">江湖大事</A> 
                                                  | <a href="#" onClick="window.open('hcjs/jhzb/index.asp','wupin','scrollbars=yes,resizable=yes,width=550,height=300')" title="查看一下自己有什么物品">我的物品</A> 
                                                  | <a href="#" onClick="window.open('hcjs/jhzb/myhua.asp','myhua','scrollbars=no,resizable=no,width=430,height=340')">我的鲜花</A> 
                                                  | <a href="JHMP/MP5.ASP" target="_blank">我的徒弟</A></div></td>
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
                                              <td colspan="6" bgcolor="#3399FF"> 
                                                <div align="center"><a href="#" onClick="window.open('diary/diary.asp','','scrollbars=no,resizable=no,width=734,height=400')">心情日记</A> 
                                                  | <A 
            href="zh/flzh.ASP" 
            target=_blank>法力转换</A> | <A 
            href="zh/qgzh.ASP" 
            target=_blank>轻功转换</A> | <a href="#"onClick="window.open('jhmp/maidou.asp','dou','scrollbars=yes,resizable=yes,width=320,height=340')" title="卖豆子！">卖威力豆</A> 
                                                  | <a href="#"onClick="window.open('hunyin/killer.asp','','scrollbars=no,resizable=no,width=700,height=400')">聘请杀手</A></div></td>
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
                                              <td><a href="flash/" target="_blank" title="最流行的歌曲、最酷的动漫Flash！">动画频道</a></td>
                                              <td><a href="#" onClick="window.open('jtbc/buy.asp','','scrollbars=no,resizable=no,width=400,height=400')">江湖博彩</a></td>
                                              <td><a href="game/ft.htm" target="_blank">拼图游戏</a></td>
                                              <td><a href="chat/sea/index.asp" target="_blank">江湖航海</a></td>
                                              <td><a href="FINDBAO/Index.asp" target="_blank">孤岛冒险</a></td>
                                              <td><a href="mj/MUJUAN.ASP" target="_blank">江湖募捐</a></td>
                                            </tr>
                                            <tr align="center"> 
                                              <td><a href="dj/" target="_blank" title="最新最劲的舞曲！让身体动起来！"><strong>DJ</strong> 
                                                舞吧</a></td>
                                              <td><a href="#"onClick="window.open('wabao/wabao.asp','diaoyu','scrollbars=yes,resizable=yes,width=450,height=340')">挖掘宝藏</a></td>
                                              <td><a title="出海走私！" style="TEXT-DECORATION: none" onclick="window.open('work/catch/catch.asp','dzwq','scrollbars=yes,resizable=yes,width=780,height=460')" href="#">江湖走私</a></td>
                                              <td><a href="tao/index.asp" target="_blank">江湖淘金</a></td>
                                              <td><a href="FENGXUE/Fengxue.asp" target="_blank">飞花舞剑</a></td>
                                              <td><a href="game/21.asp" target="_blank">疯狂21点</a></td>
                                            </tr>
                                            <tr align="center"> 
                                              <td><a href="yule.html" target="_blank">极品图库</a></td>
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
                                              <td colspan="6" bgcolor="#3399FF">&nbsp;</td>
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
                                          </table></TD>
                            <TD width=9 
                            background=51eline/ppb/download_right.gif 
                            height=12><IMG height=58 
                              src="51eline/ppb/download_right.gif" 
                            width=9></TD></TR></TBODY></TABLE></TD></TR>
                    <TR>
                      <TD vAlign=top><IMG height=20 
                        src="51eline/ppb/download_bottom.gif" 
                    width=430></TD></TR></TBODY></TABLE>
                        </TD>
                <TD width=17 height=594>&nbsp;</TD></TR></TBODY></TABLE></TD>
          <TD width=44 
height=604>&nbsp;</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=760 border=0>
  <TBODY>
  <TR>
    <TD width=45 height=54>&nbsp;</TD>
    <TD width=234><IMG height=54 src="51eline/ppb/bottom00.jpg"></TD>
        <TD><img height=54 src="51eline/ppb/bottom01.jpg" width=208></TD>
    <TD><IMG height=54 src="51eline/ppb/bottom02.jpg" width=230></TD>
    <TD width=44 height=54>&nbsp;</TD></TR></TBODY></TABLE>
<TABLE height=1 cellSpacing=0 cellPadding=0 width=760 bgColor=#ffffff 
  border=0><TBODY>
  <TR>
    <TD></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=759 align=center border=0>
  <TBODY>
  <TR>
    <TD vAlign=center>
      <DIV align=center><BR>
            <TABLE cellSpacing=0 cellPadding=0 width=760 align=center border=0>
              <TBODY>
                <TR class=black_9font> 
                  <TD><div align="center">
                      <Script language="javascript" src="count/count.asp?name=回首当年"></Script>
                    </div>
                    <div align="center"><font color="#FF0000">→</font> <font color="#000000">最佳效果：分辨率800*600 
                      浏览器IE6.0 </font><font color="#FF0000">←</font><BR>
                      <font color="#000000">版权：『快乐江湖』&#8482;　版本：ELINE V8.7.0　站长：回首当年 
                      伊然</font><br>
            Copyright &copy; 2003-2004 <a href=http://www.happyjh.com><font face=Verdana, Arial, Helvetica, sans-serif size=1><b>wWw.<font color=#CC0000>51Eline</font>.COM</b></font></a> All Rights Reserved.<BR>
                      <font color="#000000">Oicq：88617427 Email：eline_email@etang.com</font></div></TD>
                </TR>
              </TBODY>
            </TABLE>
            
          </DIV></TD></TR></TBODY></TABLE></DIV></BODY></HTML>
