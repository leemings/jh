<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
if Session("sjjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();}</script>"
	Response.End 
end if
name=sjjh_name
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM W WHERE b="& sjjh_id & " and i>0 and j=false and c<>'卡片' order by c ",conn
%>
<html>
<head>
<title>物品管理♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#3399CC" text="#000000" leftmargin="0" topmargin="0">
<div align="left">
<div align="center">时间：<font color="#000000" size="-1"><%=now()%></font>现有物品一览(装备的物品除外)<font face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font></div>
<div align="center"> <br>
<table border="1" align="center" width="558" cellpadding="1" cellspacing="0" height="31">
<tr align="center">
<td nowrap width="53" height="16">
<div align="center"><font color="#000000" size="-1">物品名</font></div>
</td>
<td nowrap width="30" height="16">
<div align="center"><font color="#000000">类型 </font></div>
<td nowrap width="40" height="16">
<div align="center"><font color="#000000" size="-1">数量 </font> </div>
<td nowrap width="41" height="16">
<div align="center"><font color="#000000">内力 </font></div>
<td nowrap width="42" height="16">
<div align="center"><font color="#000000">体力 </font></div>
<td nowrap width="39" height="16">
<div align="center"><font color="#000000">攻击 </font></div>
<td nowrap width="42" height="16">
<div align="center"><font color="#000000">防御 </font></div>
<td colspan="2" nowrap height="16">
<div align="center"><font color="#000000" size="-1">价钱</font></div>
</td>
<td nowrap width="56" height="16">
<div align="center"><font color="#000000">方式</font></div>
</td>
<td nowrap width="48" height="16">
<div align="center"><font color="#000000">数量</font></div>
</td>
<td nowrap width="54" height="16">
<div align="center"><font color="#000000">操作</font></div>
</td>
</tr>
<%
do while not rs.eof
%>
<tr bgcolor="#3399CC" onmouseout="this.bgColor='#3399CC';"onmouseover="this.bgColor='#DFEFFF';">
<form method=POST action='fs.asp?id=<%=rs("id")%>&name=<%=name%>'>
<td width="53" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("a")%>
</font> </div>
</td>
<td width="30" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("c")%></font>
</div>
<td width="40" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("i")%>
</font> </div>
<td width="41" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("e")%></font>
</div>
<td width="42" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("f")%></font>
</div>
<td width="39" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("g")%></font>
</div>
<td width="42" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("h")%></font>
</div>
<td colspan="2" height="3">
<div align="center"><font color="#000000" size="-1"><%=rs("l")%>
</font> </div>
</td>
<td height="3" width="56">
<div align="center"><font color="#000000">
<select name="wpfs">
<option value="1" selected>&nbsp;赠送</option>
<option value="2">&nbsp;交易</option>
<option value="3">&nbsp;二手</option>
<option value="4">保险柜</option>
</select>
</font></div>
</td>
<td height="3" width="48">
<div align="center"> <font color="#000000">
<select name="wpsl">
<option value="1" selected>1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
<option value="7">7</option>
<option value="8">8</option>
<option value="9">9</option>
</select>
</font></div>
</td>
<td height="3" width="54">
<div align="center"><font color="#000000">
<input type="SUBMIT" name="Submit"  value="进行">
</font></div>
</td>
</form>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<table width="64%" border="1" cellpadding="0" cellspacing="0">
<tr>
<td>说明：在这时你可以对自己的物品进行管理操作。<br>
1.赠送：把自己的东西赠送给她（他），此操作只是发生在异性之间。可以用此进行交流。<br>
2.交易：你可以与你的朋友进行物品交易，别人是买不到的。价钱由你定，亲兄弟明算账嘛。<br>
3.二手：这里你可以漫天要价，一个愿卖，一个愿买。不过让人“砸”了大头，我们可不管。<br>
4.保险柜：有什么东西想自己偷偷存起来，这里可是好地方呀。是不会丢的，不过收费哟。</td>
</tr>
</table>
</div>
</div>
</body>
</html>
