<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
nowtime=now()
nowtimetype="#"&month(nowtime)&"/"&day(nowtime)&"/"&year(nowtime)&" "&hour(nowtime)&":"&minute(nowtime)&":"&second(nowtime)&"#"
acpage=request("acpage")
if not isnumeric(acpage) then acpage=1
acpage=clng(acpage)
if acpage<1 then acpage=1
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from pet where username='"&username&"' order by option_t desc",conn,1,2
if not( rst.EOF or rst.BOF) then
rst.PageSize=20
if acpage>rst.PageCount then acpage=rst.PageCount
rst.AbsolutePage=acpage
for i=1 to rst.PageSize
if rst.EOF or rst.BOF then exit for
if rst("option_T")>nowtime then
opt=rst("option_M")&"到"&rst("option_T")
elseif rst("exist")=True then
opt="<a href='attend.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='我饿了,我困了,我倦了!';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">照顾</a>"
else
opt="<a href='visitpet.asp?id="&rst("id")&"' onmouseover="&chr(34)&"window.status='好久不见了,真想念你呀!';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">探望</a>"
end if
msg=msg&"<tr><td>"&rst("biological")&"</td><td>"&rst("sinew")&"/"&rst("maxsinew")&"</td><td>"&rst("health")&"</td><td>"&rst("cleanliness")&"</td><td>"&rst("happy")&"</td><td>"&opt&"</td></tr>"
rst.MoveNext
next
msg=msg&"</table><table width=100% border=3 bgcolor=005b00><form action='pet.asp' method=post id=form1 name=form1>"
if acpage>1 then
msg=msg&"<td><a href='pet.asp?acpage=1' onmouseover="&chr(34)&"window.status='第一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">第一页</a></td><td><a href='pet.asp?acpage="&acpage-1&"' onmouseover="&chr(34)&"window.status='前一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">前一页</a></td>"
else
msg=msg&"<td>第一页</td><td>前一页</td>"
end if
if acpage<rst.PageCount then
msg=msg&"<td><a href='pet.asp?acpage="&acpage+1&"' onmouseover="&chr(34)&"window.status='下一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">下一页</a></td><td><a href='pet.asp?acpage="&rst.PageCount&"' onmouseover="&chr(34)&"window.status='最后一页';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&">最后一页</a></td>"
else
msg=msg&"<td>后一页</td><td>最后一页</td>"
end if
msg=msg&"<td>第<input type=text name=acpage size=4 value='"&acpage&"'>页/共"&rst.PageCount&"页<input type=submit value='GO' id=submit1 name=submit1></td></tr></form>"
else
msg="<tr><td colspan=5>你没有宠物可供照顾</td></tr>"
end if
rst.close
set rst=nothing
%>
<html>
<head>
<title>宠物之家</title>
<link rel=stylesheet href='../style.css'>
</head>
<body background='../chatroom/bg1.gif'  oncontextmenu=self.event.returnValue=false>
<p align=center><br></p>
<div align="center">
<center>
<table border="0" cellspacing="1" cellpadding="0" width="64%">
<tr>
<td width="14%" align="center"><img border="0" src="../image/POKE04.GIF"></td>
<td width="14%" align="center"><img border="0" src="../image/POKE15.GIF"></td>
<td width="15%" align="center"><img border="0" src="../image/POKE13.GIF"></td>
</tr>
</table>
</center>
</div>
<hr>
<p align="center">所有属性皆为你上次看望之前的最后属性,想知道它们现在怎么样吗?去看看它们吧!
<table width=80% align=center border=3>
<tr align=center bgcolor=ffffff><td>宠物</td><td>体力</td><td>健康</td><td>清洁</td><td>快乐</td><td>操作</td></tr>
<%=msg%>
</table>
<table border="0" cellspacing="1" width="100%">
  <tr>
    <td width="100%">
    <p align="center"><a target="_self" href="../chatroom/zengmyanimal.asp">
    <font color="#FF0000">赠送给好友</font></a></td>
  </tr>
</table>

</html>