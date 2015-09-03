<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="config.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
%>
<HTML>
<HEAD>
<TITLE>我的状态-::一线网络www.happyjh.com::</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<%
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "error.asp?id=440"
%>
<style type=text/css>
<!--
body{
scrollbar-track-color:#ffffff;
SCROLLBAR-FACE-COLOR:#effaff;
SCROLLBAR-HIGHLIGHT-COLOR:#ffffff;
SCROLLBAR-SHADOW-COLOR:#eeeeee;
SCROLLBAR-3DLIGHT-COLOR:#eeeeee;
SCROLLBAR-ARROW-COLOR:#dddddd;
SCROLLBAR-DARKSHADOW-COLOR:#ffffff
}
td{font-size:9pt;}
a{font-size:9pt;text-decoration:none;}
a:hover{color:#ff0000;text-decoration:none;CURSOR:url('chat/1.cur');}
a:visited {color:0000ff;}
.hand{background-color:rgb(208,207,192);cursor:hand;}
-->
</style>
<script language="JavaScript">
<!--



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

//-->//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</HEAD>
<BODY BGCOLOR=#006699 LEFTMARGIN=1 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0>
<table border="1" width="140" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
<td width="100%" height="28">
<div align="center"><font color="#CCCCFF" size="2"><strong>我的状态</strong></font></div>
</td>
</tr>
</table>
<table width="141" border="0" cellpadding="0" cellspacing="0" background="bg2.gif">
  <tr> 
    <td height="1" colspan="3" align="right" valign="middle" bgcolor="#6666CC"> 
      <div align="right"><img src="IMAGES/blank.gif" width="1" height="1"></div></td>
    </tr>
    <tr> 
      <td width="1" bgcolor="#666699"><img src="IMAGES/blank.gif" width="1" height="1"></td>
      <td width="158" align="left" valign="top">
	  <table width="135" border="0" cellpadding="0" cellspacing="0">
        <TD COLSPAN=5 ROWSPAN=3 valign="top"> 
            <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select * from 用户 where 姓名='" & sjjh_name &"'",conn,3,3
wgj=rs("武功加")
tx=rs("名单头像")
nlj=rs("内力加")
tlj=rs("体力加")
%>
            <div align="left">
              <table width="138" border="0" cellpadding="2" cellspacing="0">
                <tr>
                  <td height="222" valign="top"> 
                    <!--#include file="z_showvisualmy.asp"-->
      <p>&nbsp;</p>
</td>
  </tr>
  <tr>
                  <td><div align="center"><img src="dot.gif"><a href="bbs/z_visual.asp" target="_blank"><font color="#CC0000"> 
                      个人形象设计</font></a></div></td>
                </tr>
    <tr>
                  <td><div align="center"><img src="dot.gif"> <a href="bbs/z_visual.asp?shopid=197" target="_blank"><font color="#009900">邀请朋友合影</font></a></div></td>
                </tr>
</table>
              &nbsp;姓名:<%=rs("姓名")%><br>
              &nbsp;性别:<%=rs("性别")%><br>
			  &nbsp;年龄:<%=rs("年龄")%><br>
			  &nbsp;地区:<%=rs("地区")%><br>
              &nbsp;配偶:<%=rs("配偶")%><br>
              &nbsp;情人:<%=rs("情人")%><br>
			  &nbsp;二号情人:<%=rs("二号情人")%><br>
			  &nbsp;三号情人:<%=rs("三号情人")%><br>
			  &nbsp;婚恋:<%=rs("结婚次数")%><br>
			  &nbsp;师傅:<%=rs("师傅")%><br>
              &nbsp;职业:<%=rs("职业")%> <br>
              &nbsp;门派:<%=rs("门派")%><br>
              &nbsp;身份:<%=rs("身份")%><br>
			  &nbsp;会员:<%=rs("会员等级")%>级<br>
			  &nbsp;会期:<%=rs("会员日期")%><br>
			  &nbsp;存点:<%=rs("allvalue")%><br>
              &nbsp;等级:<%=rs("等级")%>级<br>
              &nbsp;管理:<%=rs("grade")%>级 <br>
              &nbsp;金卡:<%=rs("会员金卡")%>元<br>
              &nbsp;金币:<%=rs("金币")%>个<br>
			  &nbsp;银币:<%=rs("银币")%>个<br>
              &nbsp;银两:<%=rs("银两")%>两<br>
              &nbsp;存款:<%=rs("存款")%>两<br>
              &nbsp;攻击:<%=rs("攻击")%><br>
              &nbsp;防御:<%=rs("防御")%><br>
              &nbsp;道德:<%=rs("道德")%><br>
              &nbsp;魅力:<%=rs("魅力")%><br>
			  &nbsp;法力:<%=rs("法力")%><br>
			  &nbsp;轻功:<%=rs("轻功")%><br>
              &nbsp;杀人:<%=rs("总杀人")%>个<br>
			  &nbsp;金:<%=rs("金")%><br>
			  &nbsp;木:<%=rs("木")%><br>
			  &nbsp;水:<%=rs("水")%><br>
			  &nbsp;火:<%=rs("火")%><br>
			  &nbsp;土:<%=rs("土")%><br>
			  &nbsp;金加:<%=rs("金加")%><br>
			  &nbsp;木加:<%=rs("木加")%><br>
			  &nbsp;水加:<%=rs("水加")%><br>
			  &nbsp;火加:<%=rs("火加")%><br>
			  &nbsp;土加:<%=rs("土加")%><br>
			  &nbsp;帖子数:<%=rs("帖子数")%><br>
              &nbsp;体力:<%if int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*100)<100 then              
abc=int(int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*100)/100)+1              
fi="chat/img/"&abc&".gif"%>
              <img height=12 src="<%=fi%>" width=<%=int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*100)%>  title="<%=rs("体力")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_tlsx+5260+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("体力")/(rs("等级")*sjjh_tlsx+5260+tlj))*100)%> title="<%=rs("体力")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_tlsx+5260+tlj%><%end if%>"> 
              <%else%>
              <%=rs("体力")%> 
              <%end if%>
              <br>
              &nbsp;内力:<%if int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)<100 then             
abc=int(int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)/100)+1               
fi="chat/img/"&abc&".gif"%>
              <img height=12 src="<%=fi%>" width=<%=int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)%> title="<%=rs("内力")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_nlsx+2000+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("内力")/(rs("等级")*sjjh_nlsx+2000+tlj))*100)%> title="<%=rs("内力")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_nlsx+2000+tlj%><%end if%>"> 
              <%else%>
              <%=rs("内力")%> 
              <%end if%>
              <br>
              &nbsp;武功:<%if int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)<100 then               
abc=int(int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)/100)+1               
fi="chat/img/"&abc&".gif"%>
              <img height=12 src="<%=fi%>" width=<%=int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)%> title="<%=rs("武功")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_wgsx+3800+tlj%><%end if%>"><img height=12 src="chat/img/5.gif" width=<%=100-int((rs("武功")/(rs("等级")*sjjh_wgsx+3800+tlj))*100)%> title="<%=rs("武功")%><%if sjjh_sx=1 then%>/<%=rs("等级")*sjjh_wgsx+3800+tlj%><%end if%>"> 
              <%else%>
              <%=rs("武功")%><br>
			  &nbsp;宝物:<%=rs("宝物")%><br>
	  &nbsp;豆子:<%if rs("grade")=3 and rs("身份")="护法" then%>
      <%=int(rs("泡豆点数")/500)%> 
      <%else%>
      <%=int(rs("泡豆点数")/1000)%> 
      <%end if%><br>
	  &nbsp;豆点:<%if rs("grade")=3 and rs("身份")="护法" then%>
      <%=rs("泡豆点数")-int(rs("泡豆点数")/500)*500%><font color="#FF0000">/500</font> 
      <%else%>                            
      <%=rs("泡豆点数")-int(rs("泡豆点数")/1000)*1000%><font color="#FF0000">/1000</font>                             
      <%end if%><br>
	  &nbsp;法力加:<%=rs("法力加")%><br>
	  &nbsp;轻功加:<%=rs("轻功加")%><br>
     &nbsp;粪库:<%=rs("粪库")%><br>
	 &nbsp;练功:<%if rs("保护")=true then%>
      练功保护                                         
      <%else%>
      没有保护                                         
      <%end if%><br></div>
    <%end if%><br></table>
	  </td>
      <td width="1" bgcolor="#666699"><img src="IMAGES/blank.gif" width="1" height="1"></td>
    </tr>
    <tr bgcolor="#666699"> 
      <td height="1" colspan="3"><img src="IMAGES/blank.gif" width="1" height="1"></td>
    </tr>
  </table>
</BODY>
</HTML>