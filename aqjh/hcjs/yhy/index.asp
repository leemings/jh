<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<HTML><HEAD><TITLE>怡红院</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../../css.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="../../bg.gif">
<div align="center">
<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#000000" width="32%" height="31">
<tr align="center"> 
<td bgcolor="#669900" width="100%" height="25"><b><font size="4" color="#FFFFFF">快乐怡红院</font></b></td>
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
            <font color="#FF0000"><%=aqjh_name%></font><font color="#800000">进入</font></font><font color="#800000" size="2">快乐</font><font size="2" color="#65251C">怡红院，在这儿，你可以找到美丽的姑娘聊天哦</font> 
            </b> 
          </td>
</tr>
<tr> 
<td align="center" rowspan="2" width="249"><img src="girl.jpg" width="150" height="174"> 
</td>
          <td width="301"> <font size="3" color="#65251C"><b><font color="#FF0000">韦春花说道：</font></b></font><font size="2" color="#8000FF">大虾，快到我们红园来，和外我们这漂亮的姑娘聊聊天，你就能增加内力哦，不过我们这儿的姑娘都是卖艺不卖身哦，如果你是姑娘，没钱的话，你也可以到我们这来工作哦，不过如果你的美貌不够的话就不来试了，嘻嘻，我们这只要美女哦。</font> 
          </td>
</tr>
<tr> 
<td width="301" valign="baseline" align="center"> 
<p><font size="2" color="#65251C"><b><a href="xiaojie.asp">人在江湖，没一个红粉知己怎么行呀，我要进红园碰碰运气。</a></b></font> 
</p>
<p><font size="2" color="#65251C"><b><a href="dengji.asp">哎，没钱了，我要在这上班。</a><br>
</b></font></p>
<p><b><font size="2"><a href="delgirl.asp"><font color="#000000">不干了，我要赎身！！！</font></a></font></b></p>
<p>&nbsp; </p>
</td>
</tr>
</table>
</td>
</tr>
<tr> 
<td colspan=2>&nbsp;</td>
</tr></tbody></table></BODY></HTML>