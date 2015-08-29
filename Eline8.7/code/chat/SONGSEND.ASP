<%@ LANGUAGE=VBScript codepage ="936" %>
<%

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
Response.Expires=0
useronlinename=Application("sjjh_useronlinename"&session("nowinroom"))
if sjjh_name="" or Session("sjjh_inthechat")<>"1" or Instr(useronlinename," "&sjjh_name&" ")=0 then Response.Redirect "chaterr.asp?id=001"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj1=n & "-" & y & "-" & r
sj2=s & ":" & f & ":" & m
sj3=sj1 & " " & sj2
songname=Request.Form("songname")
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
chatbgcolor=Application("sjjh_chatbgcolor")
chatimage=Application("sjjh_chatimage")
if chatbgcolor="" then chatbgcolor="008888"%>
<!--#INCLUDE FILE="midsound.asp" -->
<html>
<head>
<title>好歌赠好友</title>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link rel="stylesheet" href="readonly/style.css">
<script language="JavaScript">
function gozj(){url="songsend.asp?zj="+document.forms[0].songzj.options[document.forms[0].songzj.selectedIndex].value;window.location.href=url;}
function check(){document.forms[0].send.disabled=1;return true;}
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="<%=chatimage%>" bgproperties="fixed" text="#FFFFFF">
<div align="center"><font color="#FFFF00" style="font-size:12pt">好歌赠好友</font></div>
<hr size=1 color=FFFF00><br>
<table border="0" width="100%">
<form method="post" action="songsendok.asp" onsubmit='return(check());'>
<tr>
<td>
<p><font color="#FFFF00">选择曲目：</font><br>
<select name="song" style="font-size:9pt">
<%For t=1 to 10%>
<option value="<%=t%>"><%=mid(t)%></option>
<%next%>
erase mid
</select>
</p>
<p><font color="#FFFF00">选择好友：</font><br>
<select name="towho" style="font-size:9pt"><%for i=0 to x%>
<option value="<%=online(i)%>"<%if CStr(online(i))=CStr(sjjh_name) then response.write " selected"%>><%=online(i)%></option>
<%next%></select>
</p>
<p> <font color="#FFFF00">祝福话语：</font><br>
<input type="text" name="infofo" size="16" style="font-size:9pt" maxlength="100" value="谨以《%%》表达我的心意……"></p>
<p>注：%%代表歌名
</p>
<table border="0">
<tr>
<td>
<input type="submit" name="send" value="发送">
<input type="button" name="abort" value="放弃" onClick="javascript:history.go(-1)">
</td>
</tr>
</table>
</td>
</tr>
</form>
</table>
</body>
</html>