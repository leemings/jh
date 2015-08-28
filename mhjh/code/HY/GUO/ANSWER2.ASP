<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "../../error.asp?id=016"
if Session("usepro")<>"1" then
Session("usepro")=""
response.redirect "index.asp"
end if
%>
<html>
<head>
<title>问题回答</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href="../../style.css" rel=stylesheet>
</head>

<body background="../../chatroom/bg1.gif">
<table width="595" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td><%dim r,w 
r=0
w=0
for i=1 to 10 
ann=request.form("ann"&I&"")
an=request.form("an"&I&"")
response.write "第"&I&"题"
if ann=an then 
response.write "正确<br>"
r=r+1
else
response.write "错！答案是"&ann&"<br>"
w=w+1
end if
next
if r>=7 then
Session("usepro")="2"
response.redirect "go4.asp"
else
Session("usepro")=""
response.write "唉，绣花枕头呀!你没过关，点<a href='index.asp'>这里</a>回去吧~"
end if
%></td>
  </tr>
</table>
</body>
</html>
