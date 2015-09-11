<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
%>
<HTML>
<HEAD>
<TITLE>我的状态</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312">
<%
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
chatimage=Application("aqjh_chatimage")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
%>
<style type="text/css">BODY { COLOR: #000000; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt } TD { COLOR: #000000; 
FONT-FAMILY: "宋体"; FONT-SIZE: 9pt } A:link { COLOR: #000000;
FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none } A:visited { COLOR: 
#000000;FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: 
none } A:hover { COLOR: #FF0000; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; 
TEXT-DECORATION: underline blink } SELECT { BACKGROUND-COLOR: #ffcc00; BORDER-BOTTOM: 
#ffcc66 solid; BORDER-LEFT: #999900 solid; BORDER-RIGHT: #ffcc66 solid; BORDER-TOP: 
#999900 solid; FONT-SIZE: 9pt } INPUT { BACKGROUND-COLOR: RED; BORDER-BOTTOM: 
1px groove; BORDER-LEFT: 1px groove; BORDER-RIGHT: 1px groove; BORDER-TOP: 1px 
groove; COLOR: #ffffff; FONT-SIZE: 9pt } TD { }
BODY {
SCROLLBAR-FACE-COLOR: white;
 SCROLLBAR-HIGHLIGHT-COLOR: white;
 SCROLLBAR-SHADOW-COLOR: white;
 SCROLLBAR-3DLIGHT-COLOR: white;
 SCROLLBAR-ARROW-COLOR: white;
 SCROLLBAR-TRACK-COLOR: white;
 SCROLLBAR-DARKSHADOW-COLOR: white;
 SCROLLBAR-BASE-COLOR: white
} 
</style>
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
<BODY BGCOLOR=white LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 background='chat/<%=Application("aqjh_chatimage")%>' oncontextmenu=window.event.returnValue=false>
<table border="1" width="100%" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="##FF3366">
  <tr>
<td width="100%" height="28" background='<%=Application("aqjh_chatimage")%>'>
<div align="center"><font color="#FF6600" size="2"><strong>个人资料</strong></font></div>
</td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" background="../images/bg.jpg">
  <tr> 
    <td height="1" colspan="3" align="right" valign="middle" bgcolor="#6666CC"> </td>
    </tr>
    <tr> 
      <td width="1" bgcolor="#666699"> </td>
      <td width="158" align="left" valign="top">
	  <table width="135" border="0" cellpadding="0" cellspacing="0">
        <TD COLSPAN=5 ROWSPAN=3 valign="top"> 
            <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from 用户 where 姓名='" & aqjh_name &"'",conn,3,3
wgj=rs("武功加")
tx=rs("名单头像")
nlj=rs("内力加")
tlj=rs("体力加")
%>
            <div align="left">
              <table width="500" border="0" cellspacing="0" cellpadding="10">
                <tr>
    <td><div align="left"><img id=face src="<%=tx%>"></div></td>
  </tr>
</table>
              <font color=red>&nbsp;姓名:<%=rs("姓名")%>&nbsp;&nbsp;&nbsp;性别:<%=rs("性别")%>
			  &nbsp;&nbsp;&nbsp;年龄:<%=rs("年龄")%>
			  &nbsp;&nbsp;&nbsp;地区:<%=rs("地区")%>
              &nbsp;&nbsp;&nbsp;配偶:<%=rs("配偶")%>
              &nbsp;&nbsp;&nbsp;情人:<%=rs("情人")%><br><br>
			  &nbsp;婚恋:<%=rs("结婚次数")%>
			  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;师傅:<%=rs("师傅")%>
              &nbsp;&nbsp;&nbsp;职业:<%=rs("职业")%> 
              &nbsp;门派:<%=rs("门派")%>
              &nbsp;&nbsp;&nbsp;身份:<%=rs("身份")%><br><br>
			  &nbsp;会员:<%=rs("会员等级")%>级
			  &nbsp;&nbsp;&nbsp;会期:<%=rs("会员日期")%>
			  &nbsp;&nbsp;&nbsp;存点:<%=rs("allvalue")%>
              &nbsp;&nbsp;&nbsp;等级:<%=rs("等级")%>级
              &nbsp;&nbsp;&nbsp;管理:<%=rs("grade")%>级 <br><br>
              &nbsp;金卡:<%=rs("会员金卡")%>元
              &nbsp;&nbsp;&nbsp;金币:<%=rs("金币")%>个
              &nbsp;&nbsp;&nbsp;银两:<%=rs("银两")%>两
              &nbsp;&nbsp;&nbsp;存款:<%=rs("存款")%>两<br><br>
              &nbsp;攻击:<%=rs("攻击")%>
              &nbsp;&nbsp;&nbsp;攻加:<%=rs("攻击加")%>
              &nbsp;&nbsp;&nbsp;防御:<%=rs("防御")%>
              &nbsp;&nbsp;&nbsp;防加:<%=rs("防御加")%>
              &nbsp;&nbsp;&nbsp;道德:<%=rs("道德")%><br><br>
              &nbsp;魅力:<%=rs("魅力")%>
	      &nbsp;&nbsp;&nbsp;知质:<%=rs("知质")%>
              &nbsp;&nbsp;&nbsp;杀人:<%=rs("总杀人")%>个
              &nbsp;&nbsp;&nbsp;转生:<%=rs("转生")%>次
              &nbsp;花魁:<%=rs("花魁")%>
              &nbsp;状元:<%=rs("状元")%><br><br>
			  &nbsp;金:<%=rs("金")%>
			  &nbsp;&nbsp;&nbsp;木:<%=rs("木")%>
			  &nbsp;&nbsp;&nbsp;水:<%=rs("水")%>
			  &nbsp;&nbsp;&nbsp;火:<%=rs("火")%>
			  &nbsp;&nbsp;&nbsp;土:<%=rs("土")%><br><br>
			  &nbsp;金加:<%=rs("金加")%>
			  &nbsp;&nbsp;&nbsp;木加:<%=rs("木加")%>
			  &nbsp;&nbsp;&nbsp;水加:<%=rs("水加")%>
			  &nbsp;&nbsp;&nbsp;火加:<%=rs("火加")%>
			  &nbsp;&nbsp;&nbsp;土加:<%=rs("土加")%>
			  &nbsp;帖子数:<%=rs("帖子数")%><br><br>
&nbsp;豆点:<%if rs("grade")=3 and rs("身份")="长老" then%>
      <%=rs("泡豆点数")-int(rs("泡豆点数")/500)*500%>/500 
      <%else%>                            
      <%=rs("泡豆点数")-int(rs("泡豆点数")/1000)*1000%>/1000                            
      <%end if%>
	&nbsp;&nbsp;&nbsp;法力:<%=rs("法力")%>
	&nbsp;&nbsp;&nbsp;法力加:<%=rs("法力加")%>
	&nbsp;&nbsp;&nbsp;轻功:<%=rs("轻功")%><br><br>
	&nbsp;&nbsp;&nbsp;轻功加:<%=rs("轻功加")%>
     &nbsp;粪库:<%=rs("粪库")%>/5000000
	&nbsp;&nbsp;&nbsp;保护:<%if rs("保护")=true then%>闭关                                        
      <%else%>
无<%end if%><br></font></div>
<br></table>
	  </td>
      <td width="1" bgcolor="#666699"></td>
    </tr>
    <tr bgcolor="#666699"> 
      <td height="1" colspan="3"></td>
    </tr></table></BODY></HTML>
