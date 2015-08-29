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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "Select  * from 用户 where 姓名='"&sjjh_name&"'",conn
ying=rs("银两")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<HTML>
<HEAD>
<title>赌骰子♀一线网络→wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type=text/css>
BODY {
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
</style>
</HEAD>

<body text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" leftmargin="0" topmargin="0" background="../jhimg/bk_hc3w.gif">
<div align="center"><font size="+3"><font class=c>赌骰子</font></font><br>
</div>
<div align=center>
<p>- 银子最少下注是 <b><font color="#0000FF">10</font> 两</b> -<br>
最大下注银子是 <font color="#CC0000"><b><font color="#0000FF">3000</font></b></font><b>
两 </b></p>
<form method="POST" action="dicedispose.asp">
<table border=1 cellspacing=0 cellpadding=0 align="center" width="307" bordercolordark="#FFFFFF" bgcolor="#CCCCCC" bordercolorlight="#000000">
<tr bgcolor="#336633">
<td width="337" height="25"><font size="2" class=c><b>&nbsp;&nbsp;请下注</b></font></td>
</tr>
<tr bgcolor=#CCCCCC>
<td align=center width="337"><font color="#008000" size="2">你现在一共有银子 <b><%=yin%>
两</b> 可以作为赌注</font></td>
</tr>
<tr bgcolor="#CCCCCC">
<td align=center width="337"><font color="#000000" size="2">我要下注：
<input type="text" name="splosh" size="10" value="0">
&nbsp;两</font></td>
</tr>
<tr bgcolor="#CCCCCC">
<td align=center width="337">
<input type="submit" value="下注啦！！！" name="B1" >
<input type="reset" value="我要考虑一下：）" name="B2" >
</td>
</tr>
</table>
</form>

<p align="center">警告：每次赌钱，你的魅力值会扣 <font color="#FF0000">1 </font>！<br>
<br>
</p>
<p align="center"><%=Application("sjjh_chatroomname")%>版权所有</p>
</div>

</BODY>
</HTML>
