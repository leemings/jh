<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="chatconfig.asp"-->
<%'�ᱦʤ�������wWw.51eline.com��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("sjjh_useronlinename"&nowinroom)),LCase(" "&sjjh_name&" "))=0 then 
	Session("sjjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
sjjh_name=session("sjjh_name")
sjjh_grade=session("sjjh_grade")
inroom=session("nowinroom")
call roompd("ʤ������")
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
if sjjh_grade<9 then
	says=bdsays(says)
end if
act=0
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says=slgg(mid(says,i+1),inroom)
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function slgg(fn1,inroom)
	sjjh_roominfo=split(Application("sjjh_room"),";")
	chatinfo=split(sjjh_roominfo(inroom),"|")
	fjname=chatinfo(0)
	erase sjjh_roominfo
	erase chatinfo
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("sjjh_usermdb")
	rs.open "select * from �ᱦ",conn,2,2
	xb=rs("����")
	slid=rs("�ھ�")
	rs.close
	if xb then
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ��ʤ���߻�ϵͳ�Ѿ������������ᱦʤ�������ˣ�');}</script>"
		response.end
	end if
	zxrs=application("sjjh_onlinelist"&inroom)
	 '��ʼͳ��������Ա��Ŀ
	rjsq=ubound(zxrs)
	jsq=0
	sname=""
	for i=1 to rjsq
		sfgf=split(zxrs(i),"|")
		if sfgf(0)<>session("sjjh_name") and sfgf(2)<>"�ٸ�" then
			sname=sfgf(0)
			jsq=jsq+1
		end if
		if jsq>1 then
			erase zxrs
			erase sfgf
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('��ʾ��������δ���������г�������ķǹٸ���Ա�ڶᱦ�������䣡');}</script>"
			response.end
		end if
		erase sfgf
	next
	erase zxrs
	if sname="" then sname=sjjh_name
	rs.open "select id from �û� where ����='"&sname&"'",conn
	if not (rs.eof or rs.bof) then
		slid=rs("id")
		rs.close
		conn.execute "update �ᱦ set �ھ�="& slid &",ʱ��=now(),��ȡ=false,��������=0,����=true"
	else
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ���������ʤ���ߣ�"&sname&"���˺��ڽ����в����ڣ�����վ����ϵ��');}</script>"
		response.end
	end if
	set rs=nothing
	conn.close
	set conn=nothing
	gg="<font color=blue>"& sname &"</font>��������ս�����ڶ����<font color=red><b>�ᱦ�����ھ�</b></font>��վ��"& application("sjjh_user") &"����������<font color=red><b>�Ͻ��«</b></font>������<font color=blue>"& sname &"</font><img src=sjfunc/guanjun.gif>���˱���Ҫ�������������ѵĸ������޼��õ���ҡ�"
	slgg="<bgsound src=wav/tsbj.mid loop=1><table bgcolor=CC6699 width=85% align=center border=1 cellspacing=0 cellpadding='2' bordercolorlight=000000 bordercolordark=FFFFFF><tr><td align=center><font color=00FFFF><b>�ᱦ����ʤ������</b></font></td><tr><td bgcolor='#6699CC'><div align='center'><font size=-1>"&gg&"</font></div></td></tr></table>"
	call dbxx(slgg)
end function
%>