<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../const3.asp"-->
<%'��ԯ
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
onlinekill=aqjh_onlinekill
if onlinenow<onlinekill and chatinfo(0)<>aqjh_chat1 and chatinfo(0)<>aqjh_chat2 then
	Response.Write "<script language=JavaScript>{alert('��ʾ:�Ƕ�Թ�����ߵ���"&onlinekill&"�˲��ö��䣡');}</script>"
	Response.End
end if
next
%>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
%>
<!--#include file="pkif.asp"-->
<%
if aqjh_jhdj<jhdj_fz then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ҫ["&jhdj_fz&"]�����ϲſ��Բ�����');}</script>"
	Response.End
end if
if Instr(LCase(Application("aqjh_useronlinename"&nowinroom)),LCase(" "&aqjh_name&" "))=0 then 
	Session("aqjh_inthechat")="0" 
	Response.Redirect "chaterr.asp?id=001" 
end if 
towho=Trim(Request.Form("towho"))
towhoway=Request.Form("towhoway")
saycolor=Request.Form("saycolor")
addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
says=Trim(Request.Form("sy"))
'�����뿪������Ѩ�ж�
call dianzan(towho)
if len(says)>150 then says=Left(says,150)
says=lcase(says)
says=Replace(says,"&amp;","")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
iii=int(rnd()*8)+1
says="<font color=green><img src='xx/image/p"&iii&".gif' border='0'>����ԯ��<font color=" & saycolor & ">"+attack(mid(says,i+1),towho)+"</font>"
call chatsay("��ԯ",towhoway,towho,saycolor,addwordcolor,addsays,says)

'����(��ƴ����)
function attack(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,����,grade,����,����ʱ��,��¼ from �û� where ����='" & to1 &"'",conn,2,2
dlsj=DateDiff("n",rs("��¼"),now())
if dlsj<2 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�[%%]���й�����%%���߻�û��2���Ӱ���"
        exit function
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�[%%]���й�����%%�Ѿ������ˣ����ʽ����Ƿǣ�"
        exit function
end if
if rs("����")="�ٸ�" and aqjh_grade<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶԹٸ���[%%]���й�����"
        exit function
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<3 and rs("����")="��" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�[%%]���й�����%%�ոձ�����ɱ����"
        exit function
end if
if rs("�ȼ�")<30 and rs("����")="��" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�[%%]������ԯ������%%�ȼ�̫���ˣ�"
        exit function
end if
if rs("����")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�[%%]���й�����%%�������������벻Ҫ͵Ϯ��"
        exit function
end if
rs.close
rs.open "select ��Ա�ȼ�,����,����,����,����,�书,����,����,ɱ����,����ʱ��,grade,����ʱ��,��¼,z1,z2,z3,z4,z5,z6 from �û� where ����='" & aqjh_name &"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<3 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�[%%]���й�������ոձ�����ɱ������������3�����ٱ���ɣ�"
        exit function
end if
dlsj=DateDiff("n",rs("��¼"),now())
if dlsj<2 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�[%%]���й����������߻�����2���ӾͲ�Ҫ��ɱ���ˣ�"
        exit function
end if
if rs("��Ա�ȼ�")<1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�[%%]���й������㲻�ǵȼ���Ա,û��Ȩ��ʹ����ԯ��ʽ��"
        exit function
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�[%%]���й��������ǳ����˲����Բ�����"
        exit function
end if
if rs("����")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�[%%]���й��������������������벻Ҫ���У�"
        exit function
end if
if rs("ɱ����")>=int(Application("aqjh_killman"))  and aqjh_grade<10 then
	conn.execute "update �û� set ����=false where ����='" & aqjh_name &"'"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�[%%]���й�����������̫�࣬��������ɱ������"&Application("aqjh_killman")&"�����ٷ����ˣ�"
        exit function
end if
sj=DateDiff("n",rs("����ʱ��"),now())
xinxi=""
baodoudj=1
if sj<=60 then
	xinxi="<font color=red><b>��������</b></font>"
	baodoudj=3
end if
yxjfyfrom=rs("����")
yxjgjfrom=rs("����")
yxjwgfrom=rs("�书")
nlfrom=rs("����")
menpai=rs("����")
z1=null
z2=null
z3=null
z4=null
z5=null
z6=null
'ȡ���Լ�װ�������Ĺ���������(�������������;�-1)
if not(isnull(rs("z1"))) and instr(rs("z1"),"|")<>0 then
	data=split(rs("z1"),"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z1=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z1=null
			wpsm="##[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
if not(isnull(rs("z2"))) and instr(rs("z2"),"|")<>0 then
	data=split(rs("z2"),"|")
	if ubound(data)=5 then	
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z2=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z2=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
if not(isnull(rs("z3"))) and instr(rs("z3"),"|")<>0 then
	data=split(rs("z3"),"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z3=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z3=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
if not(isnull(rs("z4"))) and instr(rs("z4"),"|")<>0 then
	data=split(rs("z4"),"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z4=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z4=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
if not(isnull(rs("z5"))) and instr(rs("z5"),"|")<>0 then
	data=split(rs("z5"),"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		yxjtx=data(4)
		gjtx="<font color=red><b>["&data(0)&"]</b></font>"
		if (clng(data(1))-1)>0 then
			z5=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z5=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
if not(isnull(rs("z6"))) and instr(rs("z6"),"|")<>0 then
	data=split(rs("z6"),"|")
	if ubound(data)=5 then
		yxjfyfrom=yxjfyfrom+clng(data(2))
		yxjgjfrom=yxjgjfrom+clng(data(3))
		if (clng(data(1))-1)>0 then
			z6=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z6=null
			wpsm=wpsm&"##[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
conn.execute "update �û� set z1='"&z1&"',z2='"&z2&"',z3='"&z3&"',z4='"&z4&"',z5='"&z5&"',z6='"&z6&"' where  ����='"&aqjh_name&"'"
rs.close
'��ʼ��ԯ
rs.open "select id,c,d,e,f,g FROM n WHERE a='" & trim(fn1) & "' and b='" & aqjh_name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������ԯ��û���������书["& fn1 &"]ѽ��');}</script>"
	Response.End
end if
wgid=int(rs("id"))
wgwg=int(rs("c"))
wgnl=int(rs("d"))
wgdj=int(sqr(rs("e")/20))+1
if wgdj>=0 and wgdj<2 then
xxtx="����"
end if
if wgdj>=2 and wgdj<4 then
xxtx="�Ṧ"
end if
if wgdj>=4 and wgdj<6 then
xxtx="����"
end if
if wgdj>=6 and wgdj<8 then
xxtx="����"
end if
if wgdj>=8 and wgdj<10 then
xxtx="����"
end if
if wgdj>=10 then
xxtx="���"
end if
wgsm=rs("f")
if rs("g")="���" then
	randomize timer
	gif="<img src=../jhmp/gif/wg"& int(rnd*130)+1 &".gif height=70>"
else
	gif="<img src=../jhmp/gif/"& rs("g")&" height=70>"
end if
if rs("c")>nlfrom or rs("d")>yxjwgfrom then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�������������书����!');}</script>"
	Response.End
end if
rs.close
'���Լ��ж�
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn
mytodel=rs(xxtx)
rs.close
'�Է����ж�
rs.open "select ����,����,����,�书,��Ա�ȼ�,����,�Ṧ,����,���,z1,z2,z3,z4,z5,z6 from �û� where ����='" & to1 &"'",conn
to1=rs("����")
yxjgjto=rs("����")
yxjfyto=rs("����")
yxjwgto=rs("�书")
jhhy=rs("��Ա�ȼ�")
todel=rs(xxtx)
z1=null
z2=null
z3=null
z4=null
z5=null
z6=null
'ȡ���Է�װ�������Ĺ���������(�������������;�-1)
if not(isnull(rs("z1"))) and instr(rs("z1"),"|")<>0 then
	data=split(rs("z1"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z1=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z1=null
			wpsmto="%%[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
if not(isnull(rs("z2"))) and instr(rs("z2"),"|")<>0 then
	data=split(rs("z2"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z2=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z2=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
if not(isnull(rs("z3"))) and instr(rs("z3"),"|")<>0 then
	data=split(rs("z3"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z3=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z3=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
if not(isnull(rs("z4"))) and instr(rs("z4"),"|")<>0 then
	data=split(rs("z4"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z4=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z4=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
if not(isnull(rs("z5"))) and instr(rs("z5"),"|")<>0 then
	data=split(rs("z5"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		yxjtotx=data(4)
		if (clng(data(1))-1)>0 then
			z5=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z5=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
if not(isnull(rs("z6"))) and instr(rs("z6"),"|")<>0 then
	data=split(rs("z6"),"|")
	if ubound(data)=5 then
		yxjfyto=yxjfyto+clng(data(2))
		yxjgjto=yxjgjto+clng(data(3))
		if (clng(data(1))-1)>0 then
			z6=data(0)&"|"&(clng(data(1))-1)&"|"&data(2)&"|"&data(3)&"|"&data(4)&"|"&data(5)
		else
			z6=null
			wpsmto=wpsmto&"%%[<font color=red><b>"&data(0)&"</b></font>]������"
		end if
	end if
end if
conn.execute "update �û� set z1='"&z1&"',z2='"&z2&"',z3='"&z3&"',z4='"&z4&"',z5='"&z5&"',z6='"&z6&"' where  ����='"&to1&"'"
rs.close
randomize timer
ki=int(rnd*6)+4
killer=int(((yxjgjfrom+yxjwgfrom+yxjfyfrom)-(yxjfyto+yxjgjto+yxjwgto))/ki)*baodoudj+wgwg+wgnl+wgdj*100
xxkill=int(killer/10000)
txkill=int(killer/10000)
if killer>0 then
 if todel<xxkill then
	xxkill=todel
	txkill=todel
 end if
 xxtt="��ȡ�Է�"
else
 if mytodel<abs(xxkill) then
	xxkill=mytodel
	txkill=mytodel
 else
	xxkill=abs(xxkill)
	txkill=abs(txkill)
 end if
 xxtt="������"
end if
randomize timer
'���书��������
if aqjh_name=Application("aqjh_user") then
	conn.execute "update n set e=e+1 where id="&wgid
end if
randomize timer
ki=int(rnd*500)
if yxjtx<>"��" and yxjtx<>"" then
	if yxjtx<>yxjtotx then
		conn.execute "update �û� set "&yxjtx&"="&yxjtx&"+"&ki&" where ����='" & aqjh_name &"'"
		conn.execute "update �û� set  "&yxjtx&"="&yxjtx&"-"&ki&" where ����='" & to1 &"'"
		gjtx=gjtx&"�������ƶԷ�����ȡ�Է�[<font color=blue><b>"&yxjtx&"</b></font>]����:+<font color=red>"&ki&"</font>��!"
	else
		conn.execute "update �û� set "&yxjtx&"="&yxjtx&"-"&ki&" where ����='" & aqjh_name &"'"
		conn.execute "update �û� set  "&yxjtx&"="&yxjtx&"+"&ki&" where ����='" & to1 &"'"
		gjtx=gjtx&"�������Է������ƣ�����ȡ[<font color=blue><b>"&yxjtx&"</b></font>]����:-<font color=red>"&ki&"</font>��"
	end if
end if
gjtx=gjtx&wpsm&wpsmto
if killer>0 then
	conn.execute "update �û� set ����=����-20,����=����-" & abs(wgnl) & ",�书=�书-"& abs(wgwg) &",����=����-"& int(killer/10) & " where ����='" & aqjh_name &"'"
	conn.execute "update �û� set ����=����-"  & killer & " where ����='" & to1 &"'"
	conn.execute "update �û� set "&xxtx&"="&xxtx&"+"  & txkill & " where ����='" & aqjh_name &"'"
	conn.execute "update �û� set "&xxtx&"="&xxtx&"-"  & xxkill & " where ����='" & to1 &"'"
	djinfo="##�书��ǿ"
	mekill=abs(int(killer/10))
else
	if abs(killer)>5000 then killer=5000
	conn.execute "update �û� set ����=����-20,����=����-" & abs(wgnl) & ",�书=�书-"& abs(wgwg) &",����=����-"& abs(killer) & " where ����='" & aqjh_name &"'"
	conn.execute "update �û� set "&xxtx&"="&xxtx&"-"  & xxkill & " where ����='" & aqjh_name &"'"
	conn.execute "update �û� set "&xxtx&"="&xxtx&"+"  & xxkill & " where ����='" & to1 &"'"
	mekill=abs(int(killer/10))
	randomize timer
	killer=int(rnd*400)+1
	djinfo="##�书��������"
end if
rs.open "select ���� from �û� where ����='" & aqjh_name &"'",conn,2,2
randomize timer
wav=int(rnd*23)+1
mywav="wav"&wav&".wav"
ii=int(rnd()*8)+1
if rs("����")<-100 then
	attack=xinxi & "##<bgsound src=wav/gj/"&mywav&" loop=1>����ԯ��ʽ<img src='xx/image/img"&ii&".gif' border='0'>("& fn1 &")����["&wgdj&"��]��Ч[<font color=red>������"&xxtx&":"& txkill &"</font>]{"&wgsm&"}����"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]ʹ��%%�����½�:<font color=red>-"& killer &"</font>��,�Լ���%%���书�������:<font color=red>-"& mekill &"</font>��,<br>����%%���������##ֻ�о���ǰһЩ��˫��һ�ţ�����������̨����ȥ�ˡ���"&gjtx
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call boot(aqjh_name,"���У������ߣ�"&aqjh_name&",["&menpai&"]������������"&fn1)
	exit function
end if
rs.close
rs.open "select ����,����,����,ͨ��,�ȼ� from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")>-100 then
	attack=xinxi & "##<bgsound src=wav/gj/"&mywav&" loop=3>����ԯ��ʽ<img src='xx/image/img"&ii&".gif' border='0'>("& fn1 &")����["&wgdj&"��]��Ч[<font color=red>"&xxtt&xxtx&":"& txkill &"</font>]{"&wgsm&"}����"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]ʹ��%%�����½�:<font color=red>-"& killer &"</font>��,�Լ���%%���书�������:<font color=red>-"& mekill &"</font>��."&gjtx
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
	conn.execute "update �û� set ɱ����=ɱ����+1,��ɱ��=��ɱ��+1 where ����='" & aqjh_name &"'"
	if rs("����")=Application("aqjh_baowuname") then
		conn.execute "update �û� set ��������=0,����='"& Application("aqjh_baowuname") &"' where ����='" & aqjh_name &"'"
		conn.execute "update �û� set ��������=0,����='��',����=true,����=100,����=2000 where ����='" & to1 &"'"
		attack="##<img src='xx/image/img"&ii&".gif' border='0'>��%%�ı���:"& Application("aqjh_baowuname") &"���ߣ���õ��˱�������������Ҫ���������ſ��Եõ�����Ķ�����"&gjtx
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		exit function
	end if
if rs("�ȼ�")>=20 then
	all=rs("�ȼ�")*2
else
	all=2
end if
if rs("ͨ��")=False then
		conn.execute "update �û� set ״̬='��',����ʱ��=now(),allvalue=allvalue-"&all&",�¼�ԭ��='"&aqjh_name&"|���У�"&fn1&"' where ����='" & to1 &"'"
		conn.execute "update �û� set allvalue=allvalue+"&int(all/2)&" where ����='" & aqjh_name &"'"
		attack=xinxi & "##<bgsound src=wav/daipu.wav loop=1>����ԯ��ʽ<img src='xx/image/img"&ii&".gif' border='0'>("& fn1 &")����["&wgdj&"��]��Ч[<font color=red>��ȡ�Է�"&xxtx&":"& txkill &"</font>]{"&wgsm&"}����<img src=xx/gif/WG14.GIF>%%,%%��������..�����Ӵ�������һλ��Ϻ!"&gjtx&"%%�ݵ����<font color=red>"&all&"</font>��,��һ�뱻##���ߡ���"
else
	conn.execute "update �û� set ͨ��=False,״̬='��',����ʱ��=now(),����=0,����=0,����=0,�书=0,����=0,����=0,���=int(���/2),�¼�ԭ��='"&aqjh_name&"|���У�"&fn1&"' where ����='" & to1 &"'"
	conn.execute "update �û� set ����=����+3000000 where ����='" & aqjh_name &"'"
	attack="����ͨ������"&xinxi & "##<bgsound src=wav/daipu.wav loop=1>����ԯ��ʽ<img src='xx/image/img"&ii&".gif' border='0'>("& fn1 &")����["&wgdj&"��]��Ч[<font color=red>��ȡ�Է�"&xxtx&":"& txkill &"</font>]{"&wgsm&"}����<img src=xx/gif/WG14.GIF>%%,%%��������..Ϊ�ˣ�##�õ�300��Ԫ�ٸ�����!"&gjtx
end if
call boot(to1,"���У������ߣ�"&aqjh_name&",["&menpai&"]"&fn1)
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','��ԯ:"&fn1&"','����')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>