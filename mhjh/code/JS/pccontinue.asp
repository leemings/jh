<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
username=Session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from pc21 where username='"&username&"'",conn,1,2
zhspc=rst("z0")&";"&rst("z1")&";"&rst("z2")&";"&rst("z3")&";"&rst("z4")&";"&rst("z5")
uhspc=rst("u0")&";"&rst("u1")&";"&rst("u2")&";"&rst("u3")&";"&rst("u4")&";"&rst("u5")
upoint=rst("upoint")
randomize(timer())
do while instr(";"&zhspc&";"&uhspc&";;",";"&rndu&";")<>0
rndu=clng(rnd()*51+1)
loop
upcarr=split(uhspc,";")
upointadd=rndu mod 13
if upointadd=0 or upointadd>10 then upointadd=1
pcend=true
for i=0 to 5
if upcarr(i)=53 then
rst("u"&i)=rndu
rst("upoint")=upoint+upointadd
rst.Update
if upoint+upointadd>=21 then
Response.Redirect "pcend.asp"
else
upcarr(i)=rndu
pcend=false
end if
exit for
end if
next
if pcend=true then 	Response.Redirect "pcend.asp"
zpcarr=split(zhspc,";")
for j=0 to 5
if j<=i and zpcarr(j)<>53 then
zpcarr(j)=0
else
for i=j to 5
zpcarr(i)=53
next
exit for
end if
next
rst.Close
set rst=nothing
conn.Close
set conn=nothing
Response.Write "<html><head><link rel=stylesheet href='../style.css'><head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'><p align=center><script language=javascript>parent.imgfrm.showcard("&zpcarr(0)&","&zpcarr(1)&","&zpcarr(2)&","&zpcarr(3)&","&zpcarr(4)&","&zpcarr(5)&","&upcarr(0)&","&upcarr(1)&","&upcarr(2)&","&upcarr(3)&","&upcarr(4)&","&upcarr(5)&");</script>当前点数："&upoint+upointadd&"<input type=button value=' 继续发牌 ' onclick=javascript:location.replace('pccontinue.asp');> <input type=button value=' 开 牌 ' onclick=javascript:location.replace('pcend.asp');></p></body></html>"
%>
