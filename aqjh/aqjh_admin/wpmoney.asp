<%@ LANGUAGE=VBScript codepage ="936" %><%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
%>
<html>
<head>
<title>物品买价格设整器</title>
<link rel="stylesheet" href="../setup.css">
<script LANGUAGE="JavaScript">
<!--

function FrmAddLink_onsubmit() {
if(document.FrmAddLink.moneybei.value=="")
{
alert("银两倍数没有写，无法完成程序！")
document.FrmAddLink.moneybei.focus()
return false
}
}

//-->
</script>
</head><LINK href="css/css.css" type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<form method=post action="wpmoney1.asp" LANGUAGE="javascript"
onsubmit="return FrmAddLink_onsubmit()" name="FrmAddLink">
<table  border=1 cellspacing="1" align="center" cellpadding="0" bordercolor="#000000" bgcolor="#006699" width="229">
<tr>
<td colspan="2" height="29">
<div align="center"><font color="#FFFFFF">物品买价格设整器</font></div>
</td>
</tr>
<tr>
<td width="66">
<div align="center"><font color="#FFFFFF">类型</font></div>
</td>
<td width="154"><font color="#FFFFFF">
<select name="lx">
<option value="右手" selected>手持刀剑</option>
<option value="左手">手持护具</option>
<option value="暗器">暗器</option>
<option value="头盔">头盔</option>
<option value="双脚">双脚</option>
<option value="盔甲">盔甲</option>
<option value="装饰">装饰</option>
<option value="药品">吃用药</option>
<option value="毒药">毒药</option>
</select>
</font></td>
</tr>
<tr>
<td width="66">
<div align="center"><font color="#FFFFFF">银两倍数</font></div>
</td>
<td width="154"><font color="#FFFFFF" size="2">
<input type="text" name=moneybei value="0">
</font></td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<input type=submit value="确 定" name="submit">
<font color="#CCCCCC">------- </font>
<input  onClick="javascript:window.document.location.href='binqi.asp'" value="返 回" type=button name="back">
</div>
</td>
</tr>
</table>
</form>
<div align="center"><br>
注：选择装备类型，选择银两倍数，确下就可以自动更新物品价钱！<br>
物品价钱= (内力+体力+攻击+防御)*银两倍数<br>
如果内、体、攻、防等出现负数，将自动转换成正数！</div>
</body></html>