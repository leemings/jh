<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Write "<script>window.moveTo(200,30);</script>"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
function gjl(lb)
if isnull(rs(lb)) then
	gjl=0
	exit function
end if
data1=split(rs(lb),"|")
rs.close
rs.open "select * FROM b WHERE a='" & data1(0) &"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	gjl=0
	exit function
end if
gjl=int(rs("f")*(clng(data1(1))/rs(j)))
'myzb="<img src='../jhjs/images/"&rs("i")&"' alt='"&rs("c")&"'>["&rs("a")&"] 攻击:"&rs("f")&" 防御:"&rs("g")&" 耐久:<font color=red>"&data1(1)&"</font>/<font color=blue>"&rs("j")&"</font> 特性:"&rs("k")
rs.close
end function
function myzb(lb)
if isnull(rs(lb)) or trim(rs(lb))="" then
	rs.close
	myzb="未装备任何物品!"
	exit function
end if
data1=split(rs(lb),"|")
rs.close
rs.open "select * FROM b WHERE a='" & data1(0) &"'",conn,3,3
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	myzb="未装备任何物品!"
	exit function
end if
myzb="<img src='../jhjs/images/"&rs("i")&"' alt='"&rs("c")&"'>["&rs("a")&"] 攻击:"&rs("f")&" 防御:"&rs("g")&" 耐久:<font color=red>"&data1(1)&"</font>/<font color=blue>"&rs("j")&"</font> 特性:"&rs("k")
rs.close
end function
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT * FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
tlj=rs("体力加")
nlj=rs("内力加")
%>
<html>
<head>
<title>我的包袱♀wWw.happyjh.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#FFFFFF" background="back.gif">
<div align="center">
<table width="509" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="73" rowspan="3" valign="top">
<div align="center">
<input onClick="javascript:window.document.location.href='index.asp'" title=装备一览 type=button value="装备一览" name="button7">
<br>
<input onClick="javascript:window.document.location.href='daojun.asp'" title=装备头盔 type=button value="道具装备" name="button">
          <br>
        </div>
</td>
<td valign="top" width="436">
<div align="center"><img src="dao.gif" width="40" height="15">我的j一览<img src="dao1.gif" width="40" height="15">
<font color="#CC0000" face="幼圆"><a href="javascript:this.location.reload()">刷新</a></font>
</div>
</td>
</tr>
<tr>
<td valign="top" width="436" height="27">
<div align="center">用户<font color="#FF0000"><%=rs("姓名")%></font>状态，红色为上限！<font color="#FF0000"><font color="#000000">银子：<%=rs("银两")%>
经验：<%=rs("allvalue")%></font></font><font color="#FF0000"><font color="#000000">
等级：<%=rs("等级")%>级</font></font><br>
攻击力：<%=rs("攻击")%>
防御力：<%=rs("防御")%>
内力：<%=rs("内力")%><font color="#FF0000">/<%=rs("等级")*64+2000%></font>
体力：<%=rs("体力")%><font color="#FF0000">/<%=rs("等级")*240+5260%>
</font></div>
</td>
</tr>
<tr>
<td width="436" valign="top" height="172">
        <table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#000000" align="center">
          <tr>
            <td> 头盔：<%
            rs.close
            rs.open "SELECT z1 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3
            %>
              <%=myzb("z1")%>
            </td>
</tr>
<tr>
            <td> 装饰：<%rs.open "SELECT z2 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3 %>
              <%=myzb("z2")%>            </td>
</tr>
<tr>
            <td> 盔甲：<%rs.open "SELECT z3 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3 %>
              <%=myzb("z3")%>
            </td>
</tr>
<tr>
            <td height="2"> 左手：<%rs.open "SELECT z4 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3 %>
              <%=myzb("z4")%>
            </td>
</tr>
<tr>
            <td> 右手：<%rs.open "SELECT z5 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3 %>
              <%=myzb("z5")%>
            </td>
</tr>
<tr>
            <td> 双脚：<%rs.open "SELECT z6 FROM 用户 WHERE 姓名='"&sjjh_name&"'",conn,3,3 %>
              <%=myzb("z6")%>
<%
set rs=nothing
conn.close
set conn=nothing
%>
            </td>
</tr>
</table>
</td>
</tr>
</table>
<table width="475" border="0" cellspacing="0" cellpadding="0">
<tr>
<td valign="top" align="center">
<input onClick="JavaScript:window.close()" title=关闭窗口 type=button value="关闭窗口" name="button2">
</td>
</tr>
</table>
</div>
</body>
</html>
