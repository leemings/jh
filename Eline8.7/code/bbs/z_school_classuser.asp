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
stats="�༶��Ա"
call nav()
call head_var(0,1,boardtype,"z_school_class.asp?boardid="&boardid)

if Cint(GroupSetting(14))=0 then
	Errmsg=Errmsg+"<br>"+"<li>��û���ڱ�����ͬѧ¼��Ȩ�ޣ���<a href=login.asp>��¼</a>����ͬ����Ա��ϵ��"
	founderr=true
elseif BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
	Errmsg=Errmsg+"<br><li> �����ͬѧ¼��������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
elseif cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>��ͬѧ¼Ϊ��֤ͬѧ¼����<a href=login.asp>��¼</a>��ȷ�����Ǳ���ͬѧ��"
		founderr=true
	else
		if chkschoollogin(boardid,membername)=false then
			Errmsg=Errmsg+"<br>"+"<li>�㲻�Ǳ���ͬѧ������<a href=z_school_inclass.asp?boardid="&boardid&">���뱾��</a>��"
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
	Errmsg=Errmsg+"<br>"+"<li>��ͬѧ¼��û��ͬѧ��"
	call dvbbs_error()
	exit sub
end if
boarduser=split(replace(rs(0)," ",""),",")
dim userinfo,a,pgc,pgd
dim txlpd
txlpd=split(rs("txlpd"),"$")
set rs=nothing
pgd=ubound(boarduser)+1
pgc="��������"&pgd
a=1
response.write "<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center>"
response.write "<tr><td height=20 class=TopLighNav1 align=center colspan=5><a href=show.asp?filetype=1&boardid="&boardid&">�༶���</a> | <a href=z_school_classuser.asp?boardid="&boardid&">�༶��Ա</a> | <a href=JavaScript:openScript('z_school_sms.asp?boardid="&boardid&"',500,400)>Ⱥ�����</a> | <a href=list.asp?boardid="&boardid&">�༶��̳</a> | "
if txlpd(2)<>"" then
	response.write "<a href=http://"&txlpd(2)&" target=_blank>�༶��ҳ</a> | "
end if
response.write "<a href=z_school_inclass.asp?action=xg&boardid="&boardid&">�ҵ�����</a> | <a href=announce.asp?boardid="&boardid&">��Ҫ����</a> | <a href=z_school_inclass.asp?boardid="&boardid&">����༶</a> | <a href=z_school_inclass.asp?action=out&boardid="&boardid&">�˳��༶</a> | <a href=z_school_admin.asp?boardid="&boardid&">�༶����</a></td></tr>"
response.write "<tr><td height=25 class=tablebody1 align=center colspan=5>���๲�� <b>"&pgd&"</b> λͬѧ��</td></tr><tr><th height=25>�ա�����</th><th>E-MAIL</th><th>�硡����</th><th>�ء�����������ַ</th><th>�ǡ�����</th></tr>"
if pgd > zdes then
	pga=zdes*(page-1)
	pgb=zdes*page-1
	pgc=pgc&"����ҳ��"
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