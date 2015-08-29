<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>

<html>
<head>
<title>提交状纸♀wWw.51eline.com♀</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css.css">
<script LANGUAGE="javascript">
<!--

function FrmAddLink_onsubmit() {
if (document.FrmAddLink.topic.value=="")
{
alert("状纸题目没有填！")
document.FrmAddLink.topic.focus()
return false
}
else if(document.FrmAddLink.name.value=="")
{
alert("状告何人没有填！")
document.FrmAddLink.name.focus()
return false
}
else if(document.FrmAddLink.play.value=="")
{
alert("对被告的要求没有填！")
document.FrmAddLink.play.focus()
return false
}
else if(document.FrmAddLink.text.value=="")
{
alert("状词没有填！")
document.FrmAddLink.text.focus()
return false
}
}

//-->
</script>
</head>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF">
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="569" valign="bottom"><img src="../images/sy5.gif" width="58" height="23"></td>
<td width="209"><img src="../images/sy1.jpg" width="227" height="83"></td>
</tr>
<tr>
<td background="../images/sy4.gif" width="569" align="right"><img src="../images/sy3.gif" width="94" height="42"></td>
<td background="../images/sy4.gif" width="209"><img src="../images/sy2.jpg" width="227" height="42"></td>
</tr>
</table>
<form action=newplan.asp method=post  LANGUAGE="javascript"
onsubmit="return FrmAddLink_onsubmit()" name="FrmAddLink">
&nbsp;&nbsp;&nbsp;<b><font color="#FF6633">提交状纸</font></b>
<table border=0 cellspacing=0 cellpadding=0 align=center width="590">
<tr>
<td height="26" colspan="2">状纸题目:
<input name=topic size=50 maxlength="30">
最长30字</td>
</tr>
<tr>
<td height="17" colspan="2">状告何人:
<input name=name size=10 maxlength="10">
请写完整的姓名，否则不受理</td>
</tr>
<tr>
<td colspan="2" height="14">
<div align="left">如何处理:
<select name="play" size="1">
<option value="罚款一万" selected>罚款一万</option>
<option value="罚款十万">罚款十万</option>
<option value="坐牢">坐 牢</option>
</select>
罚款全数归受害人所有</div>
</td>
</tr>
<tr align="center">
<td colspan="2">
<div align="center"><b><img src="../jhimg/bbs/edit.gif" width="18" height="13">
下面写你的状词：<br>
</b>
<textarea name=text cols=50 rows=10></textarea>
</div>

<input type=submit value=发表 name="submit" style="background-color:FF0000;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'">
&nbsp;&nbsp;&nbsp;&nbsp;
<input type=reset value=清除 name="reset" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'">
</td>
</tr>
</table>
</form>
</body>
</html>
