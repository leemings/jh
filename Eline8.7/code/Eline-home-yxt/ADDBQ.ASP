<%Response.Expires=0
Response.Buffer=true
sjjh_id=Session("sjjh_id")
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if session("sjjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if sjjh_grade<>10 or instr(Application("sjjh_admin"),sjjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
%>
<html>
<head>
<title>装备添加♀wWw.happyjh.com♀</title>
<link rel="stylesheet" href="../setup.css">
<script LANGUAGE="JavaScript">
<!--

function FrmAddLink_onsubmit() {
if (document.FrmAddLink.wupinname.value=="")
{
alert("名称没有填！")
document.FrmAddLink.wupinname.focus()
return false
}
else if(document.FrmAddLink.wupinll.value=="")
{
alert("增加攻击没有填！")
document.FrmAddLink.wupinll.focus()
return false
}
else if(document.FrmAddLink.wupintl.value=="")
{
alert("增加防御没有填！")
document.FrmAddLink.wupintl.focus()
return false
}
}

//-->
</script>
</head>
<body text="#000000" background="../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">

<form method=post action="bqaddok.asp" LANGUAGE="javascript"
onsubmit="return FrmAddLink_onsubmit()" name="FrmAddLink">
<table  border=1 cellspacing="1" align="center" cellpadding="0" bordercolor="#000000" bgcolor="#006699" width="208">
<tr>
<td width="57" height="29">
<div align="center"><font color="#FFFFFF" size="2">兵器名</font></div>
</td>
<td width="154" height="29"><font color="#FFFFFF" size="2">
<INPUT type="text" name=wupinname>
</font></td>
</tr>
<tr>
<td width="57">
<div align="center"><font color="#FFFFFF" size="2">类型</font></div>
</td>
<td width="154"><font color="#FFFFFF" size="2">
<select name="lx">
<option value="右手" selected>手持刀剑</option>
<option value="左手">手持护具</option>
<option value="暗器">暗器</option>
</select>
</font></td>
</tr>
<tr>
<td width="57">
<div align="center"><font color="#FFFFFF" size="2">攻击</font></div>
</td>
<td width="154"><font color="#FFFFFF" size="2">
<input type="text" name=wupinll value="0">
</font></td>
</tr>
<tr>
<td width="57">
<div align="center"><font color="#FFFFFF" size="2">防御</font></div>
</td>
<td width="154"><font color="#FFFFFF" size="2">
<input type="text" name=wupintl value="0">
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
注：如果类型选择了暗器的话，程序会把攻击或防御的参数自动变为负数处理。</div>
</body>
</html>