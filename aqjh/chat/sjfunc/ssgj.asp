<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="chatconfig.asp"-->
<%'��������ϵͳ�����ļ���ֻ��Ҫָ�򱻹����ߣ�����ʽ����
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0

for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
onlinekill=onlinenow/40+1
if onlinenow<0 and chatinfo(0)<>"��Թ���" then
	Response.Write "<script language=JavaScript>{alert('����������������0�˲��ܶ��䣡���ȥ��Թ���.����û������');}</script>"
	Response.End
end if
useronlinename=Application("aqjh_useronlinename"&nowinroom)

next
%>
<%'����(��ƴ����)
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")


if chatinfo(5)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻���Է��У�');}</script>"
	Response.End
end if
f=Minute(time())
if  f>60 and chatinfo(0)<>"��Թ���" then
Response.Write "<script language=JavaScript>{alert('���޶���ʱ��ΪÿСʱ��ǰ30���ӣ�');window.close();}</script>"
	Response.End 
end if
call dbdz(1,"����")
if aqjh_jhdj<30 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������Ҫ30�����ϲſ��Բ�����');}</script>"
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
says=Replace(says,"&amp;","��")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
says="<font color=green>�����޹�����<font color=" & saycolor & ">"+attack(mid(says,i+1),towho)+"</font>"
call chatsay("����",towhoway,towho,saycolor,addwordcolor,addsays,says)

'����(��ƴ����)
function attack(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����,�ȼ�,����,grade,����,����ʱ��,���,����1,lasttime,ͨ��,����,���� from �û� where ����='" & to1 &"'",conn,2,2
tomp=rs("����")
tobliu=rs("����1")
tosf=rs("���")
npc=rs("����")
if to1="���"  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����޴��,�����������??��');}</script>"
	Response.End
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<10 and rs("����")="��" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ոո������10���ӣ������ȷŹ����ɣ�');}</script>"
	Response.End
end if
if rs("grade")>=6  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ�Թٸ����й�������');}</script>"
	Response.End
end if
if rs("�ȼ�")<=30 and rs("����")="��" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���벻Ҫ��30�����µĽ������ֲ�����');}</script>"
	Response.End
end if
if rs("����")="true"  and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Է��������������벻Ҫ͵Ϯ!');}</script>"
	Response.End
end if

sj=DateDiff("n",rs("lasttime"),now())
jfj=rs("lasttime")
if sj<1 and rs("ͨ��")=false then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	attack="<font color=red><b>##��%%����ʧ�ܣ�ԭ��##�ոյ�¼�����һ�����1���ӣ���Ҫ͵Ϯ��</b></font>"
	exit function
end if
rs.close
rs.open "select �ȼ�,����,���,����,����,����,������,������,�书,����,����,����1,ɱ����,����ʱ��,grade,����ʱ��,z1,z2,z3,z4,z5,z6,lasttime,ת��,����,�ٻ���1 from �û� where ����='" & aqjh_name &"'",conn
zaohuan1=rs("�ٻ���1")
if InStr(LCase(useronlinename)," " & LCase(zaohuan1) & "")=0 then 
Response.Write "<script language=javascript>alert('������˯����,�����ٻ����ޣ�');</script>"
Response.End 
end if
if to1=zaohuan1  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����޴��,�������ɱ??��');}</script>"
	Response.End
end if
if to1=Application("figo") then 

rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	attack="<font color=red><b>##֧ʹ���޶�%%����ʧ�ܣ�ԭ��"&Application("figo")&"���ʺ����޹���</b></font>"
	exit function
end if

if to1=aqjh_name  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����޴��,�ҿ�,i����you??��');}</script>"
	Response.End
end if

if rs("�ȼ�")<=30 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���㻹�Ǹ����֣�������30����ʱ���ٴ�ܰɣ�');}</script>"
	Response.End
end if
sj=datediff("n",rs("lasttime"),now())
if sj<1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	attack="<font color=red><b>##��%%����ʧ�ܣ�ԭ��##�ոյ�¼�����һ�����1���ӣ���׼��һ�°ɡ�</b></font>"
	exit function
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<10 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����ոո������10���ӣ������������ɣ�');}</script>"
	Response.End
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǳ����˲����Բ�����');}</script>"
	Response.End
end if
if rs("grade")>=6 and aqjh_grade<9 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����ǹٸ������벻Ҫ���й���������');}</script>"
	Response.End
end if

if rs("����")="true" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������������벻Ҫ����!');}</script>"
	Response.End
end if
if rs("ɱ����")>=int(Application("aqjh_killman"))  and aqjh_grade<10 then
	conn.execute "update �û� set ����='false' where ����='" & aqjh_name &"'"
	Response.Write "<script language=JavaScript>{alert('��ʾ��������̫�࣬��������ɱ������"& Application("aqjh_killman") &"�����ٷ����ˣ�');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

sj=DateDiff("n",rs("����ʱ��"),now())
xinxi=""
baodoudj=1
if sj<=60 then
	xinxi="<font color=red><b>��������</b></font>"
	baodoudj=3
end if
yxjfyfrom=(rs("����")+rs("������"))*rs("ת��")
yxjgjfrom=rs("����")+rs("������")
zhuan=rs("ת��")
daode=rs("����")
if daode>50000 then 
yxjwgfrom=rs("�书")*(1+zhuan/50)
end if
if daode<=-5000 then 
yxjwgfrom=rs("�书")*(1-zhuan/50)
end if
if daode>-5000 and daode<=50000 then 
yxjwgfrom=rs("�书")
end if
bliu=rs("����1")
nlfrom=rs("����")
menpai=rs("����")
mysf=rs("���")
if bliu<>"����" and menpai<>"����" and menpai<>"����" and menpai<>"����" and menpai<>"�ٸ�" and (mysf="����" or mysf="����" or mysf="����" or mysf="����") then
	zbyhdata=split(bliu,"|")
	zbsj=ubound(zbyhdata)
	if zbsj=4 then
		zbsj=DateDiff("n",zbyhdata(1),now())
		if zbsj<60 then
			xinxi=xinxi&"<font color=red><b>��##���һ����Ч����������Ϊ����߹�����</b></font>"
			yxjfyfrom=yxjfyfrom+clng(zbyhdata(3))
			yxjgjfrom=yxjgjfrom+clng(zbyhdata(2))
			yxjwgfrom=yxjgjfrom+clng(zbyhdata(4))
		end if
	end if
	erase zbyhdata
end if
'ȡ���Լ�װ�������Ĺ���������(�������������;�-1

rs.close
rs.open "select * from �û� where ����='" & zaohuan1 &"'",conn
sstl=rs("����")
'��ʼ�书
'rs.open "select id,c,d,e,f,g FROM y WHERE a='" & trim(fn1) & "' and b='" & menpai & "'",conn
'if rs.eof or rs.bof then
	'rs.close
	'set rs=nothing	
	'conn.close
	'set conn=nothing
	'Response.Write "<script language=JavaScript>{alert('��ʾ���������:"& menpai &" ��û���������书["& fn1 &"]ѽ��');}</script>"
	'Response.End
'end if
wgid=3
wgwg=10000
wgnl=10000
wgdj=20
wgsm="'"&zaohuan1&"'ʹ�����޷�������%%,%%���ŵ�һ���亹"
'if rs("g")="���" then
	randomize timer
	gif="<img src=../jhmp/gif/wg"& int(rnd*129)+1 &".gif width=80 height=80  >"
'else
	'gif="<img src=../jhmp/gif/"& rs("g")&" width=80 height=80>"
'end if

if rs("����")<10000 or rs("�书")<10000 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������������书����!');}</script>"
	Response.End
end if
rs.close
'�Է����ж�
rs.open "select ����,����,������,����,���,����,������,�书,��Ա�ȼ�,z1,z2,z3,z4,z5,z6,����,����,ת�� from �û� where ����='" & to1 &"'",conn
to1=rs("����")
daode=rs("����")
zhuan=rs("ת��")
yxjgjto=rs("����")+rs("������")
yxjfyto=rs("����")+rs("������")
if daode>50000 then 
yxjwgto=rs("�书")*(1+zhuan/50)
end if
if daode<=-5000 then 
yxjwgto=rs("�书")*(1-zhuan/50)
end if
if daode>-5000 and daode<=50000 then 
yxjwgto=rs("�书")
end if
jhhy=rs("��Ա�ȼ�")
zbyh=""
if tobliu<>"����" and tomp<>"����" and tomp<>"����" and tomp<>"����" and (tosf="����" or tosf="����" or tosf="����" or tosf="����") then
	zbyhdata=split(tobliu,"|")
	zbsj=ubound(zbyhdata)
	if zbsj=4 then
		zbsj=DateDiff("n",zbyhdata(1),now())
		if zbsj<60 then
			zbyh="<font color=red><b>��%%���һ����Ч�����ߵ���Ϊ����߹�����</b></font>"
			yxjfyto=yxjfyto+clng(zbyhdata(3))
			yxjgjto=yxjgjto+clng(zbyhdata(2))
			yxjwgto=yxjgjto+clng(zbyhdata(4))
		end if
	end if
	erase zbyhdata
end if
z1=null
z2=null
z3=null
z4=null
z5=null
z6=null
wpsmto=zbyh
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
'if killer<=100 then
'	randomize timer
'	killer=int(rnd*99)+1
'end if
randomize timer
'���书��������

randomize timer
ki=int(rnd*500)
if yxjtx<>"��" and yxjtx<>"" then
	if yxjtx<>yxjtotx then
		conn.execute "update �û� set "&yxjtx&"="&yxjtx&"+"&ki&" where ����='"&zaohuan1&"'"
		conn.execute "update �û� set  "&yxjtx&"="&yxjtx&"-"&ki&" where ����='" & to1 &"'"
		gjtx=gjtx&"�������ƶԷ�����ȡ%%[<font color=blue><b>"&yxjtx&"</b></font>]����:+<font color=red>"&ki&"</font>��!"
	else
		conn.execute "update �û� set "&yxjtx&"="&yxjtx&"-"&ki&" where ����='"&zaohuan1&"'"
		conn.execute "update �û� set  "&yxjtx&"="&yxjtx&"+"&ki&" where ����='" & to1 &"'"
		gjtx=gjtx&"�������Է������ƣ�'"&zaohuan1&"'����ȡ[<font color=blue><b>"&yxjtx&"</b></font>]����:-<font color=red>"&ki&"</font>��"
	end if
end if
gjtx=gjtx&wpsm&wpsmto
if killer>0 then
	conn.execute "update �û� set ����=����-20,����=����-" & abs(wgnl) & ",�书=�书-"& abs(wgwg) &",����=����-"& int(killer/10) & " where ����='"&zaohuan1&"'"
	conn.execute "update �û� set ����=����-"  & killer & " where ����='" & to1 &"'"
	djinfo="'"&zaohuan1&"'�书��ǿ"
	mekill=abs(int(killer/10))
else
	if abs(killer)>5000 then killer=5000
	conn.execute "update �û� set ����=����-20,����=����-" & abs(wgnl) & ",�书=�书-"& abs(wgwg) &",����=����-"& abs(killer) & " where ����='"&zaohuan1&"'"
	mekill=abs(int(killer/10))
	randomize timer
	killer=int(rnd*400)+1
'	killer=0
	djinfo="'"&zaohuan1&"'�书��������"
end if
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn,2,2
randomize timer
wav=int(rnd*13)+1
mywav="wav"&wav&".wav"
if sstl<-100 then
	attack=xinxi & "'"&zaohuan1&"'<bgsound src=wav/gj/"&mywav&" loop=1>��������ʽ����["&wgdj&"��]{"&wgsm&"}����"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]ʹ��%%�����½�:<font color=red>-"& killer &"</font>��,�Լ���%%���书�������:<font color=red>-"& mekill &"</font>��,<br>����%%���������<font color=red>"&zaohuan1&"</font>ֻ�о���ǰһ�ڣ�˫��һ�ţ�����������̨����ȥ�ˡ���"&gjtx
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call boot(zaohuan1,"���У������ߣ�"&zaohuan1&",["&menpai&"]������������")
	exit function
end if
'--------------------------------------------
randomize()
	myxy = Int(Rnd*1000)	'�˴���50000��Ϊ��������õ�װ���ĸ��ʣ���Խ�󣬻���ԽС
	ddwp=""
	tp=""
	lb1=""
	zdsj=""
	otherwp="��ͷ|̫������|��ʬ��|����|���|��ʬѪ|�̱�ʯ|����ʯ|ľ̿|ˮ��ʯ|�챦ʯ|�����"
	select case myxy
		case 10,22,323,350,411,534,978	'�˴�Ϊ�õ���Ʒװ��
			dim tmprs
			tmprs=conn.execute("Select count(*) As ��Ʒװ�� from b where lb='��Ʒ'")
			jpzs=tmprs("��Ʒװ��")-1
			set tmprs=nothing
			if isnull(jpzs) then jpzs=0
			sql="select * from b where lb='��Ʒ'"
			set rs1=conn.execute(sql)
			randomize()
			xy=int(rnd*jpzs)
			i=0
			mywpsj=""
			do while i<=xy and not(rs1.eof or rs1.bof)
				if i=xy then
					ddwp=rs1("a")
					tp=rs1("i")
					lx=rs1("b")
					select case lx
						case "����","����","ͷ��","����","˫��","װ��"
							lb1="w3"
							zdsj=rs("w3")
						case "ҩƷ"
							lb1="w1"
							zdsj=rs("w1")
						case "��Ƭ"
							lb1="w5"
							zdsj=rs("w5")
						case "��ҩ"
							lb1="w2"
							zdsj=rs("w2")
						case "����"
							lb1="w4"
							zdsj=rs("w4")
						case "�ʻ�"
							lb1="w7"
							zdsj=rs("w7")
						case "��ҩ"
							lb1="w8"
							zdsj=rs("w8")
						case else
							lb1="w6"
							zdsj=rs("w6")
					end select
				end if
				i=i+1
				rs1.movenext
			loop
			rs1.close
			set rs1=nothing
		case else	'��Ϊ�õ������otherwp�ж������Ʒ
			js=split(otherwp,"|")
			n=ubound(js)
			randomize()
			xy=int(rnd*n)
			ddwp=js(xy)
			erase js
			lb1="w6"
			zdsj=rs("w6")
			
	end select
'------------------------------------------



'----------------------------------------------
rs.close
rs.open "select ����,����,����,ͨ��,�ȼ�,���� from �û� where ����='" & to1 &"'",conn,2,2

randomize timer
npcr=int(rnd*8)+1
npcr1=int(rnd*41)+1
select case npcr
case 1
conn.execute "update �û� set ����=����-"& mekill &" where ����='" &  zaohuan1 &"'"
npc="<br><font color=red>��NPC������</font>%%��[NPC����]��ʽ{�����ջ�}����"&zaohuan1&",ɱ��"&zaohuan1&"����:<font color=red>-"& mekill *2&"</font>�㣬"&zaohuan1&"����ֱ���Ρ���"
case 2
conn.execute "update �û� set ����=����+"&npcr1&"*3583,����=����+"&npcr1&"*3228 where ����='" & to1 &"'"
npc="<br><font color=red>������������</font>%%�о�������֧����æʹ����<font color=red>[NPCר����Ч��ҩ]100��</font>,�������ϱ�����<font color=red>"&npcr1*35833&"</font>��,��������<font color=red>"&npcr1*32282&"</font>�㡭��"
case 3
conn.execute "update �û� set ����=����-"& mekill &"*3 where ����='" &  zaohuan1&"'"

npc="<br><font color=red>��NPC��ɱ����</font>%%��[NPC����]��ʽ{һ����ʽ}����"&zaohuan1&",ɱ��"&zaohuan1&"����:<font color=red>-"& mekill*6 &"</font>�㣬���"&zaohuan1&"��ͷת�򡭡�"
case 4
conn.execute "update �û� set ����=����-"& mekill &"*2 where ����='" &  zaohuan1e &"'"

npc="<br><font color=red>��NPC���С�</font>%%��[NPC����]��ʽ{��������}����"&zaohuan1&",ɱ��"&zaohuan1&"����:<font color=red>-"& mekill*4 &"</font>�㣬���"&zaohuan1&"ֱת����"

case 5

npc="<br><font size=4 color=red>%%����һ��ͷ�Σ�������Ϣ</font>"

case 6
conn.execute "update �û� set ����=����-"& mekill &"*3 where ����='" &  zaohuan1 &"'"

npc="<br><font color=red>��NPC��ɱ����</font>%%��[NPC����]��ʽ{һ����ʽ}����"&zaohuan1&",ɱ��"&zaohuan1&"����:<font color=red>-"& mekill*6 &"</font>�㣬���"&zaohuan1&"��ͷת�򡭡�"
case 7
conn.execute "update �û� set ����=����+"&npcr1&"*3583,����=����+"&npcr1&"*3228 where ����='" & to1 &"'"
npc="<br><font color=red>������������</font>%%�о�������֧����æʹ����<font color=red>[NPCר����Ч��ҩ]100��</font>,�������ϱ�����<font color=red>"&npcr1*35833&"</font>��,��������<font color=red>"&npcr1*32282&"</font>�㡭��"

case 8

npc="<br><font size=4 color=red>%%���ú�����������Ϣ</font>"

end select



'npc="<br><FONT color=green>��npc������</FONT>npc %% ������һ��,���ѹ���,��æ����"&zaohuan1&","&zaohuan1&"��Ұ��ҧ��һ��,��������"& mekill &""
if  rs("����")<>"npc" then
    aa=xinxi &"'"&zaohuan1&"'<bgsound src=wav/gj/"&mywav&" loop=1>����"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]ʹ��%%�����½�:<font color=red>-"& killer &"</font>��,�Լ���%%���书�������:<font color=red>-"& mekill &"</font>��."&gjtx  
	else 
	    aa=xinxi &"'"&zaohuan1&"'<bgsound src=wav/gj/"&mywav&" loop=1>����"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]ʹ��%%�����½�:<font color=red>-"& killer &"</font>��,�Լ���%%���书�������:<font color=red>-"& mekill &"</font>��."&gjtx&npc

end if 
if rs("����")>-100 then
    attack=aa
	rs.close 
	set rs=nothing 
	conn.close 
	set conn=nothing 
	exit function 
	else 
end if 
if rs("����")<>"npc" then
	conn.execute "update �û� set ɱ����=ɱ����+1,��ɱ��=��ɱ��+1 where ����='" & aqjh_name &"'" 
	else
		conn.execute "update �û� set ��ɱ��=��ɱ��+1 where ����='" & aqjh_name &"'" 
	
		end  if

	if rs("����")=Application("aqjh_baowuname") then 
		conn.execute "update �û� set ��������=0,����='"& Application("aqjh_baowuname") &"' where ����='" & aqjh_name &"'" 
		conn.execute "update �û� set ��������=0,����='��',����='�չ�',����=100,����=2000 where ����='" & to1 &"'" 
		attack="'"&zaohuan1&"'��%%��<img src=img/z1.gif>����:"& Application("aqjh_baowuname") &"���ߣ���������##��õ��˱�������������Ҫ���������ſ��Եõ�����Ķ�����"&gjtx
		rs.close 
		set rs=nothing	 
		conn.close 
		set conn=nothing 
		exit function 
	end if 
wpname=""&ddwp&""
'randomize timer
	ii=int(rnd()*4)+1
iii=int(rnd()*20)+1
wusl=""&ii&""

		Application("aqjh_diuqi1")=lb1&"|"&wpname&"|"&wusl&"|"&now()&"|"&npc
	Application.UnLock
'------------------------------------------------------------
randomize
s=int(rnd*500000)+1
sjy=int(rnd*200)+1
snjy=int(rnd*200)+1
yn=0
  jstl=int(rnd*200)+1
  jstl2=int(rnd*200)+1
  jstl3=int(rnd*200)+1
  jstl4=int(rnd*200)+1
  jstl5=int(rnd*200)+1
  jstl6=int(rnd*200)+1
  jstl7=int(rnd*200)+1
q4=int(rnd*8898)+1000
q1=int(rnd*8898)+100
q2=int(rnd*9898)+10
q3=int(rnd*9898)+1
	Application.Lock
Application("aqjh_zlp1")=q4
Application("aqjh_zlp2")=q1
Application("aqjh_zlp3")=q2

  Application("aqjh_ww1")=jstl
  Application("aqjh_ww2")=jstl2
  Application("aqjh_ww3")=jstl3
  Application("aqjh_ww4")=jstl4
  Application("aqjh_ww5")=jstl5
  Application("aqjh_ww6")=jstl6
  Application("aqjh_ww7")=jstl7

	Application("aqjh_by1")=iii+1
Application("aqjh_yb8")=sjy+100
	Application("aqjh_yb9")=s+100
Application("aqjh_yb7")=srjy+100
	Application.UnLock
'--------------------------------------------------------------

dlyp="    <font color=red><input type=button value='ҩƷ???��' onClick=javascript:disabled=1;window.open('by1.asp?tl1="&Application("aqjh_by1")&"','d') name=tongyi"&regjm+1&"></font>"
dlyl="     <input type=button value='���� "&Application("aqjh_yb9")&"' onClick=javascript:disabled=1;window.open('yb9.asp?tl9="&Application("aqjh_yb9")&"','d') name=tongyi"&regjm&">"
dlwp="     <input type=button value='"&wpname&" "&wusl&"��' onClick=javascript:disabled=1;window.open('dq1.asp?tl9="&Application("aqjh_yb9")&"','d') name=tongyi"&regjm&">"
dlwq="     <input type=button value='����' onClick=javascript:disabled=1;window.open('ww1.asp?tl="&Application("aqjh_ww1")&"','d') name=tongyi"&regjm&">"
dlaq="<input type=button value='����' onClick=javascript:disabled=1;window.open('ww3.asp?tl="&Application("aqjh_ww3")&"','d') name=tongyi"&regjm&">"
dlkp="<input type=button value='��Ƭ' onClick=javascript:disabled=1;window.open('ww4.asp?tl="&Application("aqjh_ww4")&"','d') name=tongyi"&regjm&">"
dlxh="<input type=button value='�ʻ�' onClick=javascript:disabled=1;window.open('ww5.asp?tl="&Application("aqjh_ww5")&"','d') name=tongyi"&regjm&">"
dldgxh="<input type=button value='�����ʻ�' onClick=javascript:disabled=1;window.open('ww6.asp?tl="&Application("aqjh_ww6")&"','d') name=tongyi"&regjm&">"
dldy="<input type=button value='��ҩ' onClick=javascript:disabled=1;window.open('ww2.asp?tl="&Application("aqjh_ww2")&"','d') name=tongyi"&regjm&">"
dljy="<input type=button value='���� "&Application("aqjh_yb8")&"' onClick=javascript:disabled=1;window.open('yb8.asp?tl8="&Application("aqjh_yb8")&"','d') name=tongyi"&regjm&">"
rdljy="<input type=button value='��� "&Application("aqjh_yb7")&"' onClick=javascript:disabled=1;window.open('yb7.asp?tl7="&Application("aqjh_yb7")&"','d') name=tongyi"&regjm&">"


randomize
dlxz=int(rnd*500000)+1
if dlxz>499900 or dlxz<=90 then 
dlzs=dlyp&dlyl&dlwp&dlaq&dlkp&dlxh&dldgxh&dldy&dljy
else if dlxz<=490900 and dlxz>=300000 then
dlzs=dlyp&dldy&dlwp&dlaq

else if dlxz<300000 and dlxz>=186000 then

dlzs=dlyp&dlaq&dldy&dldgxh

else dlzs=dlaq&dldgxh&dlyl&dldy

end if
end if
end if
if  rs("����")<>"npc" then
  ' bb=xinxi & ""&zaohuan1&"<bgsound src=wav/daipu.wav loop=1>����<img src=sjfunc/003.gif>%%,%%��������..�����Ӵ�������һλ��Ϻ!"&gjtx&"%%�ݵ����һ�أ�������~"&rdljy
	else 
	    bb=xinxi & ""&zaohuan1&"<bgsound src=wav/daipu.wav loop=1>����<img src=sjfunc/003.gif>%%,%%��������..�����Ӵ�������һλ��Ϻ!"&gjtx&"<br><FONT color=red>������Ʒ��</FONT>%%����һ˿����,������Ʒһ����·��,˭�쵽��˭�ġ���"&dlzs
	    
end if 


if rs("ͨ��")=False then 
		conn.execute "update �û� set ״̬='��',����ʱ��=now(),�¼�ԭ��='"&aqjh_name&"|���У����޹���' where ����='" & to1 &"'"
		attack=bb
end if 
if rs("ͨ��")=True then 
 conn.execute "update �û� set ͨ��=False,״̬='��',��ɱ��=��ɱ��-1,����=0,����=0,����=0,�书=0,����=0,����=0,���=int(���/2),�¼�ԭ��='"&aqjh_name&"|���У����޹���' where ����='" & to1 &"'" 
 conn.execute "update �û� set ����=����+10000000 where ����='" & aqjh_name &"'" 
 attack="����ͨ������"&xinxi & ""&zaohuan1&"����"&gif&"%%,%%��������..Ϊ�ˣ�����##�õ�һǧ��Ԫ�ٸ�����!"&gjtx 
end if 

call boot(to1,"���У������ߣ�"&zaohuan1&",["&menpai&"]") 
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & zaohuan1 & "','���޹���','����')" 
rs.close 
set rs=nothing	 
conn.close 
set conn=nothing 
end function 
%>
