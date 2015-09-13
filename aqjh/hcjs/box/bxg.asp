<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../const.asp"-->
<!--#include file="../../mywp.asp"-->
<%
Response.Write "<SCRIPT LANGUAGE=javascript>if(window.name!='bxg'){this.location.href='index.asp'}</SCRIPT>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs1.open "select id,会员等级,w1,w2,w3,w4,w5,w6,w7,w8 from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
if rs1.eof or rs1.bof then
	rs1.close
	set rs1=nothing
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>{alert('你不是江湖中人！');location.href = 'javascript:history.go(-1)';}</script>" 
	Response.End 
end if
myid=rs1("id")
hydj=rs1("会员等级")

wpzs=10
wpzs=wpzs+hydj*5	'得出最多可存放几种物品
rs.open "select * from w where a="&myid&" order by d",conn,2,2
%>
<html>
<head>
<title>我的储物箱</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
a:active {  font: 12px "宋体"; color: #FFFFFF; background: #000000}
a:link {  font: 12px "宋体"; color: #FFFFFF}
a:visited {  font: 12px "宋体"; color: #FFFFFF}
body {  font: 12px "宋体"; color: #FFFFFF}
input {  font: 12px "宋体"; color: #FFFFFF; text-decoration: none; background: #000000; border: none}
td {  font: 12px "宋体"; color: #FFFFFF; text-align: center}
-->
</style>
<link rel="stylesheet" href="STYLE.CSS" type="text/css">
</head>
<body bgcolor="#000000" text="#FFFFFF" topmargin="0" leftmargin="0" link="#00CCFF" vlink="#00CCFF" alink="#00CCFF" oncontextmenu=self.event.returnValue=false>
<br>
<div align="center"><font size="5" color="#FF0000"><b class="3d"><font size="6">我的储物箱</font></b></font><br>
<br>
<font color="#FFFF00">点击右边的物品名，再按“放入储物箱”按钮即可放入储物箱。</font><br><br>
<a href="#" onclick="javascript:parent.bxg.document.location.reload();">点击刷新储物箱</a><br>
<table width="82%" border="1" cellpadding="0" cellspacing="0">
<tr>
<form name="bxgform" method="post" action="inthebox.asp">
<td colspan="5" height="29">
<input type="hidden" name="filedname" size="5" maxlength="5" value="">
<input type="text" readonly name="wpname" size="10" maxlength="15" class="input">
<input type="text" name="sl" size="5" maxlength="5" class="input">
<input type="submit" name="Submit2" value="放入储物箱" style="border: 1 solid #FFFFFF">
</td>
</form>
</tr>
<tr>
<td width="199" height="24"><font color="#FFCC66"><b>物品名称</b></font></td>
<td width="117" height="24"><font color="#FFCC66"><b>所属类别</b></font></td>
<td width="93" height="24"><font color="#FFCC66"><b>数量</b></font></td>
<td width="68" height="24"><font color="#FFCC66"><b>取出数量</b></font></td>
<td width="150" height="24"><font color="#FFCC66"><b>操作</b></font></td>
</tr>
<%jsq=0
if rs.eof or rs.bof then%>
<tr><td colspan="5" height="29">你的储物箱中什么也没有</td></tr>
<%else
do while not rs.bof and not rs.eof
	tempp=""
if jsq>=wpzs then
	yhfiled=rs("d")
	yid=rs("a")
	wid=rs("id")
	wm=rs("b")
	ws=rs("c")
	conn.execute "delete * from w where id="&wid
	tempp=add(rs1(yhfiled),wm,ws)
	conn.execute "update 用户 set "&yhfiled&"='"&tempp&"' where id="&yid
else
%>
<tr>
<td width="199"><%=rs("b")%></td>
<%sslb="无"
select case rs("d")
	case "w1"
		sslb="补药"
	case "w2"
		sslb="毒药"
	case "w3"
		sslb="装备"
	case "w4"
		sslb="暗器"
	case "w5"
		sslb="卡片"
	case "w6"
		sslb="物品"
	case "w7"
		sslb="鲜花"
	case "w8"
		sslb="配药"
	case else
		sslb="类别出错"
end select
%>
<td width="117"><%=sslb%></td>
<td width="93"><%=rs("c")%></td>
<form name="form2" method="post" action="qwupin.asp?id=<%=rs("id")%>">
<td width="68" align="center">
<input type="text" name="sl" size="5" maxlength="5" style="background-color: #800000; color: #FFFFFF; border: 1 solid #52436D">
</td>
<td width="150" align="center">
<input type="submit" name="Submit3" value="取出物品"  style="border: 1 solid #FFFFFF">
</td>
</form>
</tr>
<%jsq=jsq+1
end if
rs.movenext
loop%>
<tr>
<td colspan="5" height="29"><br>
        你的储物箱共存放了<font color="#FFFF00"><%=jsq%></font>种物品<br>
<br>
<div align="left"><font color="#00FFFF">非会员可以存放10种物品，一级会员可存放15种物品<br>
          二级会员可存放20种物品，三级会员可存放25种物品，四级会员可存放30种物品<br>
          当会员到期后或物品总数超出应有的数量时则自动将储物箱中放不下的物品自动加到自已身上。</font><br>
</div>
</td>
</tr>
<%end if
rs1.close
set rs1=nothing
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<br>
<FONT color=#0000ff>&copy; 版权所有 2005-2006 </FONT><A href="http://www.7758530.com/" target=_blank><FONT color=#0000ff>快乐江湖网</FONT></A></div></body></html>