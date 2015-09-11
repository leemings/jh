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
<style type=text/css>
<!--
body {color:ffffff;}
td{font-size:9pt;}
a:link {color:#ffCC00;}
a:visited {color:0000ff;}
a:hover{color:#FF9900;; cursor: hand}
.hand{background-color:rgb(208,207,192);cursor:hand;}
-->
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
<BODY BGCOLOR=#006699 LEFTMARGIN=0 TOPMARGIN=0 MARGINWIDTH=0 MARGINHEIGHT=0 background='chat/<%=Application("aqjh_chatimage")%>' oncontextmenu=window.event.returnValue=false>
<table border="1" width="100%" cellspacing="0" cellpadding="0" bgcolor="#006699" height="16" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr>
<td width="100%" height="28" background='<%=Application("aqjh_chatimage")%>'>
<div align="center"><font color="#CCCCFF" size="2"><strong>我的状态</strong></font></div>
</td>
</tr>
</table>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" background="BG34.GIF">
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
              <table width="130" border="0" cellspacing="0" cellpadding="10">
                <tr>
    <td><div align="left"><img id=face src="<%=tx%>"></div></td>
  </tr>
</table>
              &nbsp;姓名:<%=rs("姓名")%><br>
              &nbsp;性别:<%=rs("性别")%><br>
			  &nbsp;年龄:<%=rs("年龄")%><br>
			  &nbsp;地区:<%=rs("地区")%><br>
              &nbsp;配偶:<%=rs("配偶")%><br>
              &nbsp;情人:<%=rs("情人")%><br>
			  &nbsp;婚恋:<%=rs("结婚次数")%><br>
			  &nbsp;师傅:<%=rs("师傅")%><br>
              &nbsp;职业:<%=rs("职业")%> <br>
              &nbsp;种族:<%=rs("种族")%> <br>
              &nbsp;国家:<%=rs("国家")%> <br>
              &nbsp;职位:<%=rs("职位")%> <br>
              &nbsp;门派:<%=rs("门派")%><br>
              &nbsp;身份:<%=rs("身份")%><br>
			  &nbsp;会员:<%=rs("会员等级")%>级<br>
			  &nbsp;会期:<%=rs("会员日期")%><br>
			  &nbsp;存点:<%=rs("allvalue")%><br>
              &nbsp;等级:<%=rs("等级")%>级<br>
              &nbsp;管理:<%=rs("grade")%>级 <br><font color=red>
              &nbsp;金卡:<%=rs("会员金卡")%>元<br>
              &nbsp;金币:<%=rs("金币")%>个<br>
              &nbsp;银两:<%=rs("银两")%>两<br>
              &nbsp;存款:<%=rs("存款")%>两<br></font>
              &nbsp;攻击:<%=rs("攻击")%><br>
              &nbsp;攻加:<%=rs("攻击加")%><br>
              &nbsp;防御:<%=rs("防御")%><br>
              &nbsp;防加:<%=rs("防御加")%><br>
              &nbsp;道德:<%=rs("道德")%><br>
              &nbsp;魅力:<%=rs("魅力")%><br>
              &nbsp;威望:<%=rs("威望")%><br>
              &nbsp;精神:<%=rs("精神")%><br>
              &nbsp;魔力:<%=rs("魔力")%><br>
              &nbsp;仙术:<%=rs("仙术")%><br>
              &nbsp;黄金:<%=rs("黄金")%><br>
              &nbsp;智力:<%=rs("智力")%><br>
              &nbsp;佛法:<%=rs("佛法")%><br>
	          &nbsp;知质:<%=rs("知质")%><br>
	          &nbsp;元气:<%=rs("元气")%><br>
              &nbsp;杀人:<%=rs("总杀人")%>个<br>
              &nbsp;转生:<%=rs("转生")%>次<br>
              &nbsp;花魁:<%=rs("花魁")%><br>
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
&nbsp;豆点:<%if rs("grade")=3 and rs("身份")="护法" then%>
      <%=rs("泡豆点数")-int(rs("泡豆点数")/500)*500%><font color="#FF0000">/500</font> 
      <%else%>                            
      <%=rs("泡豆点数")-int(rs("泡豆点数")/1000)*1000%><font color="#FF0000">/1000</font>                             
      <%end if%><br>
	  &nbsp;法力:<%=rs("法力")%><br>
	  &nbsp;法力加:<%=rs("法力加")%><br>
	  &nbsp;轻功:<%=rs("轻功")%><br>
	  &nbsp;轻功加:<%=rs("轻功加")%><br>
     &nbsp;粪库:<%=rs("粪库")%><font color="#FF0000">/5000000</font><br>
	 &nbsp;保护:<%if rs("保护")=true then%>闭关                                        
      <%else%>
无<%end if%><br></div>
<br></table>
	  </td>
      <td width="1" bgcolor="#666699"></td>
    </tr>
    <tr bgcolor="#666699"> 
      <td height="1" colspan="3"></td>
    </tr></table></BODY></HTML>