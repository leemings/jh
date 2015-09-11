<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
wupinid=clng(Request("wupinid"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sqlstr="select * from b where id="&wupinid
rs.open sqlstr,conn
if rs.EOF or rs.BOF then
Response.Redirect "modifywupin.htm"
Response.End
else
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
<title>兵器修改</title></head>

<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<p align="center"><font color="#000000" size="2">兵器修改<br>
  <br>
  注：类型修改请填药品，暗器，兵器，头盔，盔甲，双脚，装饰，毒药,卡片九种类型，否则添加无效。</font></p>
<form method="post" action="modifynow.asp">
<table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="305">
<tr>
<td width="105"><font color="#FFFFFF" size="2">ID</font></td>
<td width="189"><font color="#FFFFFF" size="2">
<input type="text" name="wupinid" readonly
value="<%=rs("ID")%>" size="20">
</font></td>
</tr>
<tr>
<td width="105"><font color="#FFFFFF" size="2">物品名</font></td>
<td width="189"><font color="#FFFFFF" size="2">
<input type="text" name="wupinname"
value="<%=rs("a")%>" size="20">
</font></td>
</tr>
<tr>
<td width="105"><font color="#FFFFFF" size="2">类型</font></td>
<td width="189"><font color="#FFFFFF" size="2">
<input type="text" name="wupinlx"
value="<%=rs("b")%>" size="20">
</font></td>
</tr>
<tr>
<td width="105"><font color="#FFFFFF" size="2">内力</font></td>
      <td width="189"><font color="#FFFFFF" size="2"> 
        <input type="text" name="wupinnl"
value="<%=rs("d")%>" size="20">
</font></td>
</tr>
<tr>
<td width="105"><font color="#FFFFFF" size="2">体力</font></td>
<td width="189"><font color="#FFFFFF" size="2">
<input type="text" name="wupintl"
value="<%=rs("e")%>" size="20">
</font></td>
</tr>
<tr>
<td width="105"><font color="#FFFFFF" size="2">攻击</font></td>
<td width="189"><font color="#FFFFFF" size="2">
<input type="text" name="wupingj"
value="<%=rs("f")%>" size="20">
</font></td>
</tr>
<tr>
<td width="105"><font color="#FFFFFF" size="2">防御</font></td>
<td width="189"><font color="#FFFFFF" size="2">
<input type="text" name="wupinfy"
value="<%=rs("g")%>" size="20">
</font></td>
</tr>
<tr>
<td width="105"><font size="2" color="#FFFFFF">说明</font></td>
<td width="189"><font color="#FFFFFF" size="2">
<input type="text" name="wupinsm"
value="<%=rs("c")%>" size="20">
</font></td>
</tr>
<tr>
<td width="105"><font color="#FFFFFF" size="2">银两</font></td>
<td width="189"><font color="#FFFFFF" size="2">
<input type="text" name="wupinyl"
value="<%=rs("h")%>" size="20">
</font></td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<input type="submit" value="确 定">
<font color="#CCCCCC">------- </font>
<input  onClick="javascript:window.document.location.href='binqi.asp'" value="返 回" type=button name="back">
</div>
</td>
</tr>
</table>
</form>
<%
end if
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<div align="center"> <font color="#FFFFFF" size="2">数据库更新程序因为时间有限，没有加入为空值时的检测，请大家更改时注意没有值的地方填0</font></div>