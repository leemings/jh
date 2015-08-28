<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=session("yx8_mhjh_username")
if username="" then
Response.Write "<script language=javascript>top.location.replace('../error.asp?id=016');</script>"
Response.End
end if
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject ("adodb.recordset")
rst.Open "select * from biological where biological='"&session("yx8_mhjh_userfight")&"'",conn
if rst.EOF or rst.BOF then
msg="<script language=javascript>location.href='action.asp'</script>"
else
biological=rst("biological")
hp=rst("hp")
picture=rst("picture")
attack=rst("attack")
defence=rst("defence")
rst.Close
today=date()
todaytype="#"&month(today)&"/"&day(today)&"/"&year(today)&"#"
rst.Open "select * from fight where username='"&username&"'",conn
if rst.EOF or rst.BOF then
conn.Execute "insert into fight(username,biological,hp,fnum,ldate) values('"&username&"','"&biological&"',"&hp&",0,"&todaytype&")"
else
if rst("ldate")<today then
conn.Execute "update fight set biological='"&biological&"',hp="&hp&",fnum=0,ldate="&todaytype&" where username='"&username&"'"
else
conn.Execute "update fight set biological='"&biological&"',hp="&hp&" where username='"&username&"'"
end if
end if
msg="<head><link rel=stylesheet href='../style.css'></head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif' topmargin=0 left margin=0><form name=form1><table border=3 align=center><tr><td align=center bgcolor=FFFF00>"&biological&"</td></tr><tr><td align=center><img src='"&picture&"' width=54 height=72></td></tr><tr><td><input type=text size=12 name=hp value='ÉúÃü£º"&hp&"' disabled></td></tr><tr><td><input type=text size=12 name=attack value='¹¥»÷£º"&attack&"' disabled></td></tr><tr><td><input type=text size=12 name=defence value='·ÀÓù£º"&defence&"' disabled></td></tr><tr><td></td></tr></table></form></body>"
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write msg
%>