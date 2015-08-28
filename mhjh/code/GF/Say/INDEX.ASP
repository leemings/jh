<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username=""  then Response.Redirect "../../error.asp?id=016"
dim conn,rs
%>

<html>
<head>
<title>提交状纸</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet><script LANGUAGE="javascript">
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

<body background='../../chatroom/bg1.gif' oncontextmenu=self.event.returnValue=false leftmargin="5" marginwidth="5">
<table width=506 border=0 cellspacing=0 cellpadding=0 align=center height="19">
<tr>
<td height="26" width="508"><b>提交状纸</b></td>
<td align=right height="26" width="7">
<div align="center">　</div>
</td>
</tr>
</table>
<form action=newplan.asp method=post  LANGUAGE="javascript"
onsubmit="return FrmAddLink_onsubmit()" name="FrmAddLink">
<table border=0 cellspacing=0 cellpadding=0 align=center width="506">
<tr>
<td height="26" colspan="2" width="504"> <b>状纸题目：</b>
<input name=topic size=50 maxlength="30">   
最长30字</td>   
</tr>   
<tr>   
<td height="17" colspan="2" width="504"><b>状告何人：   
<input name=name size=20 maxlength="20">   
请写完整的姓名，否则不受理</b></td>   
</tr>   
<tr>   
<td colspan="2" height="14" width="504">   
<div align="left"><b>要求对被告的处理：   
<select name="play" size="1">   
<option value="罚款一万" selected>罚款一万</option>   
<option value="罚款十万">罚款十万</option>   
<option value="坐牢">坐 牢</option>   
</select>   
罚款全数归受害人所有</b></div>   
</td>   
</tr>   
<tr>   
<td colspan="2" width="504">   
<div align="center"><b>下面书写你的状词：<br>   
</b>   
<textarea name=text cols=50 rows=15></textarea>   
</div>   
<p align="center">   
<input type=submit value=发表 name="submit" style="background-color:FF0000;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'"> 
&nbsp;&nbsp;&nbsp;&nbsp; <input type=reset value=清除 name="reset" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'">  
</p>  
</td>  
</tr>  
</table>  
</form>  
<p align="center">爱火神话&nbsp; Copyright 2001-2003 [ <a href="manage.asp">官府入口</a>   
]</p>   
</body>   
</html>   
  
