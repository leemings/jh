<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/z_school_char.asp" -->
<%
'=========================================================
' File: z_school_classuser.asp
' Version:5.0
' Date: 2003-1-13
' Script Written by wxzlxl http://www.zmcn.com QQ:628122
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'========================================================= 
const zdes=20
dim page,pga,pgb
if request("page")="" or not isnumeric(request("page")) then
	page=1
else
	page=clng(request("page"))
end if
stats="班级成员"
call nav()
call head_var(0,1,boardtype,"z_school_class.asp?boardid="&boardid)

if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>您没有在本社区同学录的权限，请<a href=login.asp>登录</a>或者同管理员联系。"
	founderr=true
elseif BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
	Errmsg=Errmsg+"<br><li> 错误的同学录参数！请确认您是从有效的连接进入。"
	founderr=true
elseif cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>本同学录为认证同学录，请<a href=login.asp>登录</a>并确认您是本班同学。"
		founderr=true
	else
		if chkschoollogin(boardid,membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>你不是本班同学，请先<a href=z_school_inclass.asp?boardid="&boardid&">加入本班</a>。"
			founderr=true
		end if
	end if
end if
if founderr then
	call dvbbs_error()
else
	call class1()
end if
call activeonline()
call footer()

sub class1()
dim boarduser,userzz
set rs=conn.execute("select boarduser,txlpd from board where boardid="&boardid)
if rs(0)<>"" then
	boarduser=split(replace(rs(0)," ",""),",")
else
	Errmsg=Errmsg+"<br>"+"<li>该同学录还没有同学！"
	call dvbbs_error()
	exit sub
end if
boarduser=split(replace(rs(0)," ",""),",")
dim userinfo,a,pgc,pgd
dim txlpd
txlpd=split(rs("txlpd"),"$")
set rs=nothing
pgd=ubound(boarduser)+1
pgc="本班人数"&pgd
a=1
response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"
response.write "<tr><td height=20 class=TopLighNav1 align=center colspan=5><a href=show.asp?filetype=1&boardid="&boardid&">班级相册</a> | <a href=z_school_classuser.asp?boardid="&boardid&">班级成员</a> | <a href=JavaScript:openScript('z_school_sms.asp?boardid="&boardid&"',500,400)>群体短信</a> | <a href=list.asp?boardid="&boardid&">班级论坛</a> | "
if txlpd(2)<>"" then
	response.write "<a href=http://"&txlpd(2)&" target=_blank>班级主页</a> | "
end if
response.write "<a href=z_school_inclass.asp?action=xg&boardid="&boardid&">我的资料</a> | <a href=announce.asp?boardid="&boardid&">我要留言</a> | <a href=z_school_inclass.asp?boardid="&boardid&">加入班级</a> | <a href=z_school_inclass.asp?action=out&boardid="&boardid&">退出班级</a> | <a href=z_school_admin.asp?boardid="&boardid&">班级管理</a></td></tr>"
response.write "<tr><td height=25 class=tablebody1 align=center colspan=5>本班共有 <b>"&pgd&"</b> 位同学。</td></tr><tr><th height=25>姓　　名</th><th>E-MAIL</th><th>电　　话</th><th>地　　　　　　址</th><th>城　　市</th></tr>"
if pgd > zdes then
	pga=zdes*(page-1)
	pgb=zdes*page-1
	pgc=pgc&"，分页："
	for i = 1 to (pgd+zdes-1)/zdes 
		if i = page then
			pgc=pgc&" [<font color=#FF0000>"&i&"</font>] "
		else
			pgc=pgc&" [<a href=z_school_classuser.asp?boardid="&boardid&"&page="&i&">"&i&"</a>] "
		end if
	next
else
	pga=0
	pgb=pgd-1
end if
if pgb > ubound(boarduser) then
	pgb=ubound(boarduser)
end if
dim ss,dd,ff,gg
for i = pga to pgb
	if boarduser(i)<>"" then
		set rs=conn.execute("select userid,useremail,userinfo from [User] where username='"&boarduser(i)&"'")
		if rs(2)<>"" then
			userinfo=split(rs(2),"|||")
			ss=userinfo(0)
			dd=userinfo(5)
			ff=userinfo(13)
			gg=userinfo(14)
		else
			ss=" "
			dd=" "
			ff=" "
			gg=" "
		end if
		response.write "<tr height=25><td class=tablebody"&a&">&nbsp;<a href=z_school_user.asp?id="&rs(0)&"&boardid="&boardid&" target=_blank>"&ss&"</a>(<a href=dispuser.asp?id="&rs(0)&"&boardid="&boardid&" target=_blank>"&boarduser(i)&"</a>)</td><td class=tablebody"&a&">&nbsp;"&rs(1)&"</td><td class=tablebody"&a&">&nbsp;"&ff&"</td><td class=tablebody"&a&">&nbsp;"&gg&"</td><td class=tablebody"&a&">&nbsp;"&dd&"</td></tr>"
		set rs=nothing
	end if
	if a=1 then a=2 else a=1
next
response.write "<tr height=25><td height=25 class=tablebody1 colspan=5>&nbsp;"&pgc&"</td></tr>"
response.write "</TABLE>"
'response.write userzz
end sub
%>