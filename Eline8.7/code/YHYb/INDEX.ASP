<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>『E线江湖』-人妖馆</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../pic/css.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="../pic/bg.jpg">
<div align="center">
<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#000000" width="32%" height="31">
<tr align="center"> 
<td bgcolor="#669900" width="100%" height="25"><b><font size="4" color="#FFFFFF">人妖馆<img border="0" src="../chat/img/tu1.gif"></font></b></td>
</tr>
</table>
<br>
</div>
<table border=0 cellpadding=0 cellspacing=0 width="601" align="center">
<tbody> 
<tr> 
<td class=bg1 colspan=2 rowspan=2 
style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign=top> 
<table border="0" cellpadding="0" cellspacing="0" width="554">
<tr> 
          <td align="center" colspan="2" width="552"> <b> 
 <font size="2" color="#65251C"> 
            <font color="#FF0000"><%=sjjh_name%></font></font><font size="2" color="#008080">进入人妖馆，在这儿，你可以找到“美丽的鸭”聊天哦</font> 
            </b> 
          </td>
</tr>
<tr> 
<td align="center" rowspan="2" width="249"><img border="0" src="../chat/img/004.gif"><img border="0" src="../chat/img/17s.jpg"> 
</td>
          <td width="301"> <font size="3" color="#65251C"><b><font color="#FF0000">掌柜无花说道：</font></b></font><font size="2" color="#800000">大虾，快到我们人妖馆来，您可以和他们聊天或是发泄喔，还能增加内力哦，不过我们这儿的人妖都是卖艺不卖身哦，如果你是美男，没钱的话，你也可以到我们这来工作哦，不过如果你的美貌不够的话就不用来试了，嘻嘻，我们这只要美男哦。</font> 
          </td>
</tr>
<tr> 
<td width="301" valign="baseline" align="center"> 
<p><b><a href="xiaojie.asp"><font size="2" color="#669900">该乐则乐，没必要到泰国吧，进人妖馆碰碰运气</font> 
</a></b> 
</p>
<p><font size="2" color="#65251C"><b><a href="dengji.asp">哎，没钱了，我要在这上班。</a><br>
</b></font></p>
<p><b><font size="2"><a href="delgirl.asp"><font color="#FF0000">不干了，我要赎身！！！</font></a></font></b></p>
<p>&nbsp; </p>
</td>
</tr>
</table>
</td>
</tr>
<tr> 
<td colspan=2>&nbsp;</td>
</tr>
</tbody> 
</table>
</BODY></HTML> 
