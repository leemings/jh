<%
pcbet=Request.Form("pcbet")
if not isnumeric(pcbet) then Response.Redirect "../error.asp?id=024"
pcbet=clng(pcbet)
if pcbet>9999 or pcbet<1 then Response.Redirect "../error.asp?id=024"
username=Session("yx8_mhjh_username")
if username="" then Response.Redirect "../error.asp?id=016"
dim zdice(2)
randomize()
do
zdice(0)=clng(rnd()*5+1)
zdice(1)=clng(rnd()*5+1)
zdice(2)=clng(rnd()*5+1)
if zdice(0)=zdice(1) and zdice(0)=zdice(2) then
zgrade=3
ratio=10
zpoint=zdice(0)*3
elseif zdice(0)=zdice(1) then
zgrade=2
ratio=2
zpoint=zdice(2)
elseif zdice(1)=zdice(2) then
zgrade=2
ratio=2
zpoint=zdice(0)
elseif zdice(0)=zdice(2) then
zgrade=2
ratio=2
zpoint=zdice(1)
end if
loop while not(zdice(0)=zdice(1) or zdice(1)=zdice(2) or zdice(0)=zdice(2))
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
rst.Open "select * from 用户 where 姓名='"&username&"' and 银两>="&pcbet,conn
if rst.EOF or rst.BOF then Response.Redirect "../error.asp?id=030"
rst.Close
rst.Open "select * from dice where username='"&username&"'",conn
if rst.EOF or rst.BOF then
conn.Execute "insert into dice values('"&username&"',"&zdice(0)&","&zdice(1)&","&zdice(2)&","&zgrade&","&zpoint&",0,"&ratio&","&pcbet&",0)"
else
conn.Execute "update dice set z0="&zdice(0)&",z1="&zdice(1)&",z2="&zdice(2)&",zgrade="&zgrade&",zpoint="&zpoint&",frequency=0,ratio="&ratio&",bet="&pcbet&" where username='"&username&"'"
end if
rst.Close
set rst=nothing
conn.Execute "update 用户 set 银两=银两-"&pcbet&" where 姓名='"&username&"'"
conn.Close
set conn=nothing
response.write "<html><head><link rel=stylesheet href='../style.css'></head><body oncontextmenu=self.event.returnValue=false background='../chatroom/bg1.gif'><p align=center><input type=button value=' 开始 ' onclick=javascript:location.replace('pdcontinue.asp');></p><script language=javascript>parent.imgfrm.showdice(0,"&zdice(0)&","&zdice(1)&","&zdice(2)&");parent.imgfrm.document.images[3].src=parent.imgfrm.imagedb[0].src;parent.imgfrm.document.images[4].src=parent.imgfrm.imagedb[0].src;parent.imgfrm.document.images[5].src=parent.imgfrm.imagedb[0].src;</script></body></html>"
%>
