<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<%'���һ��
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(session("nowinroom")),"|")
fjname=chatinfo(0)
erase aqjh_roominfo
erase chatinfo
if aqjh_jhdj<jhdj_fz or aqjh_grade<2 then
	Response.Write "<script language=JavaScript>{alert('��ʾ�����һ����Ҫ["&jhdj_fz&"]��������ȼ�2�����ϲſ��Բ�����');}</script>"
	Response.End
end if
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","&")
if aqjh_grade<9 then
	says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=red><b>�����һ����</b><font color=" & saycolor & ">"+zbyh()+"</font>"
call chatsay(act,towhoway,towho,saycolor,addwordcolor,addsays,says)
'����
function zbyh()
nowinroom=session("nowinroom")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select grade,�ȼ�,����,���,����1 from �û� where ����='"&aqjh_name&"'",conn,2,2
bliu=rs("����1")
grade=rs("grade")
mp=rs("����")
sf=rs("���")
rs.close
select case sf
	case "����"
		wgxs=0.02
		gjxs=0.02
		fyxs=0.02
	case "����"
		wgxs=0.015
		gjxs=0.015
		fyxs=0.015
	case "����"
		wgxs=0.01
		gjxs=0.01
		fyxs=0.01
	case "����"
		wgxs=0.005
		gjxs=0.005
		fyxs=0.005
end select
if mp="�ٸ�" or mp="����" or mp="����" or mp="����" or grade<2 or (sf<>"����" and sf<>"����" and sf<>"����" and sf<>"����") then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�������ٸ��������������˲�����ʹ�����һ�����ܣ�����Ҫ�����Ϊ���š����ϡ��������������ſ���ʹ�ã�');}</script>"
	Response.End
end if
zb3=0	'�ӹ���
zb4=0	'�ӷ���
zb5=0	'���书
if bliu<>"����1" then
	zbdata=split(bliu,"|")
	zs=ubound(zbdata)
	if zs=4 then
		zb1=clng(zbdata(0))	'ʹ�ô���
		zb2=zbdata(1)	'���һ��ʹ��ʱ��
		erase zbdata
		sj=DateDisj("d",zb2,now())
		if zb1>=3 and sj<=0 then
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=JavaScript>{alert('������Ѿ�ʹ�ù��������һ�������ˣ�ÿ��������ʹ�����Σ�');}</script>"
			Response.End
		end if
	else
		zb1=0	'ʹ�ô���
		zb2=now()	'���һ��ʹ��ʱ��
	end if
end if
'�������ߵ�����
rs.open "select ����,�书,����,���� from �û� where ����='"&mp&"' and grade<"& grade
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	zbyh="<font color=red>["&mp&"]</font><font color=green>["&sf&"]</font>##��ʹ�����һ����������ѵĹ���������ϧ"&mp&"��û��һ����##����ȼ��͵ĵ��ӣ����һ������ʹ��ʧ�ܣ�"
	exit function
end if

mpzxrs=0
zbjwg=0
zbjgj=0
zbjfy=0 
'��ʼͳ��������Ա��Ŀ
do while not(rs.eof or rs.bof)
	name=rs("����")
	if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom"))),LCase(" "&name&" "))<>0 then
		mpzxrs=mpzxrs+1
		zbjwg=int(zbjwg+rs("�书")*wgxs)
		zbjgj=int(zbjgj+rs("����")*gjxs)
		zbjfy=int(zbjfy+rs("����")*fyxs)
	end if
	rs.movenext
loop
rs.close
if mpzxrs=0 then
	set rs=nothing
	conn.close
	set conn=nothing
	zbyh="<font color=red>["&mp&"]</font><font color=green>["&sf&"]</font>##��ʹ�����һ����������ѵĹ���������ϧ"&mp&"��û��һ����##����ȼ��͵ĵ������ߣ����һ������ʹ��ʧ�ܣ�"
	exit function
end if
zb1=zb1+1
zb2=now()
zb3=zbjgj
zb4=zbjfy
zb5=zbjwg
bliu=zb1&"|"&zb2&"|"&zb3&"|"&zb4&"|"&zb5
conn.execute "update �û� set ����1='"& bliu &"' where ����='"&aqjh_name&"'"
set rs=nothing
conn.close
set conn=nothing
zbyh="##���������ҵĵ����Ǵ��һ������<font color=red><b>"&mp&"</b></font>���ֵ��ǣ����ŵ���ͬ�����������ˣ����Ҫͬ��Э����Ϊ����������������"&mp&"������������"&mpzxrs&"������һ����Ӧ��##�书��ʱ���ӣ�<font color=red><b>"&zbjwg&"</b></font>���������ӣ�<font color=red><b>"&zbjgj&"</b></font>���������ӣ�<font color=red><b>"&zbjfy&"</b></font>����Чʱ��Ϊ<font color=red><b>60����</b></font>��"
end function
%>