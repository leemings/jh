<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc.asp"-->
<!--#include file="func.asp"-->
<!--#include file="../../mywp.asp"-->
<!--#include file="../../const3.asp"-->
<%
aqjh_roominfo=split(Application("aqjh_room"),";")
nowinroom=session("nowinroom")
chatroomnum=ubound(aqjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("aqjh_useronlinename"&i)),"  ")
	onlinenum=ubound(online)+1
	onlinenow=onlinenow+onlinenum
chatinfo=split(aqjh_roominfo(nowinroom),"|")
useronlinename=Application("aqjh_useronlinename"&nowinroom)
onlinekill=aqjh_onlinekill
if onlinenow<onlinekill and chatinfo(0)<>aqjh_chat1 and chatinfo(0)<>aqjh_chat2 then
	Response.Write "<script language=JavaScript>{alert('��ʾ:�Ƕ�Թ�����ߵ���"&onlinekill&"�˲��ö��䣡');}</script>"
	Response.End
end if
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
nowinroom=session("nowinroom")
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
says=Replace(says,"&amp;","&")
'��ϵͳ��ֹ�ַ�����
if aqjh_grade<9 then
says=bdsays(says)
end if
act=1
towhoway=0
i=instr(says,"$")
fnn1=mid(says,i+1)
'�ж϶�npc����
if InStr(";" & Application("aqjh_npc"), ";" & towho & "|")<>0 then
   pk="��NPC������"
else
   pk="�����䡿"
end if
'�ж϶�npc����
says="<font color=red>"&pk&"<font color=" & saycolor & ">"+attack(mid(says,i+1),towho)+"</font>"
call chatsay("����",towhoway,towho,saycolor,addwordcolor,addsays,says)
function attack(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='" & to1 &"'",conn,2,2
if rs.eof then
   rs.close
   rs.open "select * from [npc] where n����='" & to1 &"'",conn,2,2
   if rs.eof then
      wt=""
   else
      wt="npc"
   end if
else
   wt="ren"
end if
rs.close
select case wt
case "npc"
'��ʼ�ж�npc
rs.open "select * from npc where n����='" & to1 &"'",conn,2,2
if rs.eof or rs.bof then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��NPC["&to1&"]�������ڻ���������NPC�����˰ɣ�');}</script>"
	Response.End	
end if
if rs("n����")<0 then
        rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
'ɾ������NPC����
nowinroom=session("nowinroom")
Application.Lock
onlinelist=Application("aqjh_onlinelist"&session("nowinroom"))
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),to1 & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application("aqjh_useronlinename"&nowinroom)=Replace(Application("aqjh_useronlinename"&nowinroom)," " & to1 & " ","")
Application("aqjh_npc")=Replace(Application("aqjh_npc"),";"&to1&"|","")
Application.UnLock
response.write "<script Language=Javascript>parent.m.location.reload();</script>"
attack="##:NPC�Ѿ����������ˣ��ٸ��Ѿ��ջأ�"
exit function
end if
rs.close
rs.open "select * from �û� where ����='" & aqjh_name &"'",conn,2,2
if rs("����")=True and aqjh_grade<10 then
    attack="##:�����ڱ����У��벻Ҫ��NPC���й�����"
    exit function
end if
if rs("����")="����" then
    attack="##:���ǳ����ˣ����ܲ���NPC��ܣ�"
    exit function
end if
yxjfyfrom=rs("����")
yxjgjfrom=rs("����")
yxjwgfrom=rs("�书")
yxjdjfrom=rs("�ȼ�")
nlfrom=rs("����")
menpai=rs("����")
rs.close
'��ʼ�书
rs.open "select id,c,d,e,f,g FROM y WHERE a='" & trim(fn1) & "' and b='" & menpai & "'",conn
if rs.eof or rs.bof then
    attack="##:�������:"&menpai&"��û���������书["&fn1&"]ѽ��"
    exit function
end if
wgid=int(rs("id"))
wgwg=int(rs("c"))
wgnl=int(rs("d"))
wgdj=int(sqr(rs("e")/100))+1
wgsm=rs("f")
if rs("g")="���" then
	randomize timer
	gif="<img src=../jhmp/gif/wg"& int(rnd*130)+1 &".gif height=70>"
else
	gif="<img src=../jhmp/gif/"& rs("g")&" height=70>"
end if
if rs("c")>nlfrom or rs("d")>yxjwgfrom then
    attack="##:�����������书���㣬���ܶ�npc���й������Ȳ�һ���ɣ�"
    exit function
end if
rs.close
'��NPC���ж�
rs.open "select * from npc where n����='" & to1 &"'",conn
to1=rs("n����")
yxjgjto=rs("n����")
yxjfyto=rs("n����")
yxjwgto=rs("n�书")
yxjdjto=int(sqr(int(rs("n����")/50)))
nlto=rs("n����")
nzhuren=rs("n����")
ndiren=rs("n����")
sj=DateDiff("n",rs("n����ʱ��"),now())
rs.close
'�ж�����
if nzhuren=aqjh_name then
    attack="##:NPC%%����Ĳ��£��㲻�ܶ�����������������������"
    exit function
end if
'�жϵ���
if sj>5 then
if ndiren<>aqjh_name and ndiren<>"��" and Instr(LCase(useronlinename),LCase(" "&ndiren&" "))<>0 then
    attack="##:NPC%%����"&ndiren&"����У��벻Ҫ���֣���Ȼ%%ҧ���㣡"
    exit function
end if
end if
randomize timer
ki=int(rnd*6)+4
killer=int(((yxjgjfrom+yxjwgfrom+yxjfyfrom)-(yxjfyto+yxjgjto+yxjwgto)*(yxjdjfrom/yxjdjto))/ki)+wgwg+wgnl+wgdj*100
if killer>0 then
	conn.execute "update �û� set ����=����-20,����=����-" & abs(wgnl) & ",�书=�书-"& abs(wgwg) &",����=����-"& int(killer/30) & " where ����='" & aqjh_name &"'"
	conn.execute "update npc set n����=n����-"&int(killer/15)&",n����='"&aqjh_name&"',n����ʱ��=now(),n����ʱ��=now() where n����='" & to1 &"'"
	djinfo="##�书��ǿ"
        mekill=abs(int(killer/30))
        killer=int(killer/15)
else
	if abs(killer)>5000 then killer=5000
        killer=int(abs(killer)/10)
	mekill=killer*10
        conn.execute "update �û� set ����=����-20,����=����-" & abs(wgnl) & ",�书=�书-"& abs(wgwg) &",����=����-"&mekill&" where ����='" & aqjh_name &"'"
	conn.execute "update npc set n����=n����-"&killer&",n����='"&aqjh_name&"',n����ʱ��=now(),n����ʱ��=now() where n����='" & to1 &"'"
	djinfo="##�书��������"
end if
rs.open "select ���� from �û� where ����='" & aqjh_name &"'",conn,2,2
randomize timer
wav=int(rnd*23)+1
mywav="wav"&wav&".wav"
randomize timer
dz=int(rnd*9)+1
select case dz
case 1
 dz_ok="�����޹�"
case 2
 dz_ok="��������"
case 3
 dz_ok="����˫Ȯ"
case 4
 dz_ok="Ǭ���跨"
case 5
 dz_ok="��ɽ����"
case 6
 dz_ok="��������"
case 7
 dz_ok="�������"
case 8
 dz_ok="����ȭ"
case 9
 dz_ok="�񽣷�ħ"
end select
if rs("����")<-100 then
	attack="##<bgsound src=wav/gj/"&mywav&" loop=1>��["& menpai &"]��ʽ("& fn1 &")����["&wgdj&"��]{"&wgsm&"}����"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]ʹ��NPC[%%]�����½�:<font color=red>-"& killer &"</font>��,�Լ�������[%%]����һ��<fonr color=red>["&dz_ok&"]</font>,ʹ�������½�:<font color=red>-"& mekill &"</font>��,##�似�����ˣ���%%���ҧ������"
        conn.execute "update npc set n����='��' where n����='" & to1 &"'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        call boot(aqjh_name,"��NPCҧ���������ߣ�"&aqjh_name&",["&menpai&"]������������"&fn1)
	exit function
end if
rs.close
rs.open "select * from npc where n����='" & to1 &"'",conn,2,2
zlwp=rs("n��Ʒ")
zlyl=rs("n����")
n_jy=rs("n����")
n_dj=int(sqr(int(rs("n����")/50)))
n_baochu=(n_dj/aqjh_jhdj)*1000
n_jjy=int((n_dj/aqjh_jhdj)*500)
add_type=rs("n��Ʒ����")
nzhuren=rs("n����")
if zlyl>0 then
   njy="<font color=red>��NPC��Ѫ��<font color=" & saycolor & ">NPC[%%]����"&zlyl&"������������������<font color=red>"&zlyl*0.1&"</font>�㡣<br>"
   conn.execute "update npc set n����=0,n����=n����+"&zlyl*0.1&" where n����='" & to1 &"'"
end if
if rs("n����")>-100 then
	attack="##<bgsound src=wav/gj/"&mywav&" loop=1>��["& menpai &"]��ʽ("& fn1 &")����["&wgdj&"��]{"&wgsm&"}����"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]ʹ��NPC[%%]�����½�:<font color=red>-"& killer &"</font>�㡣<br>"
        attack=attack&njy&"<font color=red>��NPC������<font color=" & saycolor & ">����##�ǵ���NPC[%%]��NPC[%%]��һ��������ʹ����һ��<font color=red>["&dz_ok&"]</font>��ʹ��##�����½�:<font color=red>-"& mekill &"</font>��."
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
'�Ƿ��б�����Ʒ
if n_baochu>=Application("aqjh_npcwp") then
'������Ʒ
data1=split(zlwp,";")
data2=ubound(data1)
for kk=0 to data2-1
    data3=split(data1(kk),"|")
    wp_sl=data3(1)
    wp_name=trim(data3(0))
    Application.Lock
    Application("npc_wp"&kk)=1
    Application.UnLock
    randomize()
    regjm=int(rnd*3348998)
    wp_list=wp_list&"<input type=button value='"&wp_name&"|"&wp_sl&"��' name=npc"&regjm&" style='background-color:#86A231;color:FFFFFF;border: 1 double' onClick=javascript:npc"&regjm&".disabled=1;window.open('npc_wp.asp?dj="&aqjh_jhdj&"&wp_name="&wp_name&"&wp_sl="&wp_sl&"&kk="&kk&"&wp_type="&add_type&"','d','scrollbars=0,resizable=0,width=450,height=500')> "
next
    '�Ƿ�������װ��
     if aqjh_jhdj<n_dj then
       rs.close
       rs.open "SELECT * FROM x where f='��Ʒ' order by g",conn,1,2
       wp_count=rs.recordcount+2
       randomize timer
       wp_rnd=int(rnd*wp_count)+1
       if wp_rnd<rs.recordcount then
          rs.moveFirst
          rs.move wp_rnd
          wp_jian=rs("a")
       else
          wp_jian="��"
       end if
       if wp_jian<>"��" then
        rs.close
        rs.open "select * from �û� where ����='" & aqjh_name &"'",conn,2,2
        zstemp=add(rs("w3"),wp_jian,1)
        conn.execute "update �û� set w3='"&zstemp&"' where  ����='"&aqjh_name&"'"
        jian="<font color=red><b>##�õ�һ��"&wp_jian&"��</b></font>"
       else
        jian="<font color=red><b>##ֻ����һ�Ѳ�Ѱ���ı���˲����ʧ�ˣ�����ֻ���Լ��������ã�</b></font>"
       end if
     end if 
conn.execute "update npc set n����='"&aqjh_name&"',n������=n������+1,n����ʱ��='"&now()&"' where n����='" & to1 &"'"
attack="##<bgsound src=wav/daipu.wav loop=1>��["& menpai&"]��ʽ("& fn1 &")����["&wgdj&"��]{"&wgsm&"}����"&gif&"NPC[%%],%%����##�Ķ��֣��Ѿ�����һϢ.."&jian&"[%%]����������Ʒ�ķ����ѣ�������һ�أ���ҿ���ѽ��"&wp_list
'ս��Ʒ
else
  attack="##<bgsound src=wav/daipu.wav loop=1>��["& menpai &"]��ʽ("& fn1 &")����["&wgdj&"��]{"&wgsm&"}����"&gif&"NPC[%%],%%����##�Ķ��֣��Ѿ�����һϢ...���ǣ�û�õ��κζ�����"
end if
conn.execute "update �û� set allvalue=allvalue+"&n_jjy&" where ����='" &aqjh_name&"'"
attack=attack&"<br><font color=ff0099>����ϲ��</font>##������NPC����%%������ֵ����<font color=red>"&n_jjy&"</font>�㡣"
'ɾ������NPC����
nowinroom=session("nowinroom")
Application.Lock
onlinelist=Application("aqjh_onlinelist"&session("nowinroom"))
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),to1 & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application("aqjh_useronlinename"&nowinroom)=Replace(Application("aqjh_useronlinename"&nowinroom)," " & to1 & " ","")
Application("aqjh_npc")=Replace(Application("aqjh_npc"),";"&to1&"|","")
Application.UnLock
response.write "<script Language=Javascript>parent.m.location.reload();</script>"
case "ren"
rs.open "select ת��,״̬,�ȼ� from �û� where ����='" & aqjh_name &"'",conn
zstt=rs("ת��")
todj=rs("�ȼ�")
rs.close
rs.open "select ����,�ȼ�,����,grade,����,ɱ��ʱ��,��¼,ת�� from �û� where ����='" & to1 &"'",conn,2,2
dlsj=DateDiff("n",rs("��¼"),now())
if dlsj<3 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�%%���й�����%%���߻�û��3���Ӱ���"
        exit function
end if
dlsj=DateDiff("n",rs("ɱ��ʱ��"),now())
if dlsj<30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##��<font color=red><b>�㲻�ܶ�%%���й��������ոձ���ɱ�������30���Ӻ�ɣ�</b></font>"
        exit function
end if
if rs("�ȼ�")<=90 and rs("ת��")=0 and todj>=90 then
if rs("�ȼ�")+100<=todj  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	attack="<font color=red><b>##��%%����ʧ�ܣ�ԭ��%%����90��,��##�����%%��100����,���ܶ��䰡��</b></font>"
	exit function
end if
end if
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�%%���й��������ǳ����ˣ�"
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
if rs("�ȼ�")<=30 and rs("����")="��" and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�%%���й�������Ҫ�Գ��뽭�����ֲ�����"
        exit function
end if
if rs("����")=True and aqjh_grade<10 and zstt<5 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�%%���й������Է��������������벻Ҫ͵Ϯ!"
        exit function
end if
rs.close
rs.open "select ����,����,����,ת��,�ȼ�,����,�书,����,����,ɱ����,����ʱ��,grade,����ʱ��,��¼,z1,z2,z3,z4,z5,z6 from �û� where ����='" & aqjh_name &"'",conn
if rs("����")="����" and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�%%���й��������ǳ����˲����Բ���!"
        exit function
end if
if rs("�ȼ�")>90 and zstt=0 and todj<90 then
if rs("�ȼ�")-100>todj then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	attack="<font color=red><b>##��%%����ʧ�ܣ�ԭ��%%����90��,��##�����%%��100����,���ܶ���Ŷ��</b></font>"
	exit function
end if
end if
dlsj=DateDiff("n",rs("��¼"),now())
if dlsj<3 and aqjh_grade<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�%%���й����������߻�����3���ӾͲ�Ҫ��ɱ���ˣ�"
        exit function
end if
if rs("����")=True and aqjh_grade<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�%%���й��������������������벻Ҫ���У�"
        exit function
end if
if rs("ɱ����")>=int(Application("aqjh_killman"))  and aqjh_grade<10 then
	conn.execute "update �û� set ����=false where ����='" & aqjh_name &"'"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
        attack="##���㲻�ܶ�%%���й�����������̫�࣬��������ɱ������"& Application("aqjh_killman") &"�����ٷ����ˣ�"
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
'ȡ���Լ�װ�������Ĺ���������(�������������;�-1
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
'��ʼ�书
rs.open "select id,c,d,e,f,g FROM y WHERE a='" & trim(fn1) & "' and b='" & menpai & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���������:"& menpai &" ��û���������书["& fn1 &"]ѽ��');}</script>"
	Response.End
end if
wgid=int(rs("id"))
wgwg=int(rs("c"))
wgnl=int(rs("d"))
wgdj=int(sqr(rs("e")/100))+1
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
'�ж϶Է���û��npc
rs.close
rs.open "select top 1 * from NPC where n����='"&to1&"'",conn
if not rs.eof then
  n_name=rs("n����")
  n_tl=rs("n����")
  n_nl=rs("n����")
  n_wg=rs("n�书")
  n_gj=rs("n����")
  n_fy=rs("n����")
  n_diren=rs("n����")
  if InStr(";" &Application("aqjh_npc"), ";" & n_name & "|")<>0 then
  isnpc=true
  end if
end if
'�ж϶Է���û��npc
'�ж��Լ���û��npc
rs.close
rs.open "select top 1 * from NPC where n����='"&aqjh_name&"'",conn
if not rs.eof then
  my_n_name=rs("n����")
  my_n_dj=int(sqr(int(rs("n����")/50)))
  my_n_diren=rs("n����")
  if InStr(";" &Application("aqjh_npc"), ";" & my_n_name & "|")<>0 then
  ismynpc=true
  end if
end if
'�ж��Լ���û��npc
'�Է����ж�
rs.close
rs.open "select ����,����,����,�书,�ȼ�,��Ա�ȼ�,z1,z2,z3,z4,z5,z6 from �û� where ����='" & to1 &"'",conn
to1=rs("����")
to1_dj=rs("�ȼ�")
yxjgjto=rs("����")
yxjfyto=rs("����")
yxjwgto=rs("�书")
jhhy=rs("��Ա�ȼ�")
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
if isnpc=true then
npc_kill=int((yxjgjfrom+yxjwgfrom+yxjfyfrom)/2-(n_gj+n_fy+n_wg))
end if
randomize timer
'���书��������
if menpai<>"����" then
	conn.execute "update y set e=e+1 where id="&wgid
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
'�ж�npc
        if isnpc=true then
	conn.execute "update �û� set ����=����-"  & int(killer/2) & " where ����='" & to1 &"'"
        conn.execute "update �û� set ����=int(����-����*0.07),�书=int(�书-�书*0.07),����=int(����-����*0.1) where ����='" &aqjh_name &"'"
        npc_attack="<font color=red>��NPC�������˹�����</font>����##��%%���й���,NPC����"&n_name&":��������!������##�书��ǿ,ֻ���˵����ˣ��������书���½���1/14�������½���1/10..."
        else
	conn.execute "update �û� set ����=����-"  & killer & " where ����='" & to1 &"'"
        end if
	djinfo="##�书��ǿ"
	mekill=abs(int(killer/10))
else
	if abs(killer)>5000 then killer=5000
'�ж�npc
        if isnpc=true then
        npc_kill=abs(npc_kill)
        conn.execute "update �û� set ����=����-"&int(npc_kill/20)& ",�书=�书-"&int(npc_kill/20)&",����=����-"&int(npc_kill/10)&" where ����='" & aqjh_name &"'"
        npc_attack="<font color=red>��NPC�������˹�����</font>����##��%%���й���,NPC����"&n_name&"��������!������##�书��������,�������书���½�<font color=red>-"&int(npc_kill/20)&"</font>�㣬�����½�<font color=red>-"&int(npc_kill/10)&"</font>�㡣"
        end if
       	conn.execute "update �û� set ����=����-20,����=����-" & abs(wgnl) & ",�书=�书-"& abs(wgwg) &",����=����-"& abs(killer) & " where ����='" & aqjh_name &"'"
	mekill=abs(int(killer/10))
	randomize timer
	killer=int(rnd*400)+1
	djinfo="##�书��������"
end if
'npc�ж�
if isnpc=true then
 if Instr(LCase(useronlinename),LCase(" "&n_diren&" "))=0 or n_diren="��" then
       conn.execute "update npc set n����='"&aqjh_name&"',n����ʱ��=now() where n����='"&n_name&"'"
 end if
end if
if ismynpc=true then
 if Instr(LCase(useronlinename),LCase(" "&my_n_diren&" "))=0 or my_n_diren="��" then
     conn.execute "update npc set n����='"&to1&"',n����ʱ��=now() where n����='"&my_n_name&"'"
 end if
 my_abate=int(my_n_dj*500/to1_dj)
 my_npc_attack="<font color=red>��NPC��ս������</font>NPC����["&my_n_name&"]�������˹���%%,��ȥ��ҧ��%%һ��,%%��ʧ<font color=red>����-"&my_abate&"</font>"
 conn.execute "update �û� set ����=����-"&my_abate&" where ����='"&to1&"'"
end if
rs.open "select ���� from �û� where ����='" & aqjh_name &"'",conn,2,2
randomize timer
wav=int(rnd*13)+1
mywav="wav"&wav&".wav"
if rs("����")<-100 then
	attack=xinxi & "##<bgsound src=wav/gj/"&mywav&" loop=1>��["& menpai &"]��ʽ("& fn1 &")����["&wgdj&"��]{"&wgsm&"}����"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]ʹ��%%�����½�:<font color=red>-"& killer &"</font>��,�Լ���%%���书�������:<font color=red>-"& mekill &"</font>��,����%%���������##ֻ�о���ǰһЩ��˫��һ�ţ�����������̨����ȥ����"&gjtx
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	call boot(aqjh_name,"���У������ߣ�"&aqjh_name&",["&menpai&"]������������"&fn1)
	exit function
end if
rs.close
rs.open "select ����,����,����,ͨ�� from �û� where ����='" & to1 &"'",conn,2,2
if rs("����")>-100 then
	attack=xinxi & "##<bgsound src=wav/gj/"&mywav&" loop=1>��["& menpai &"]��ʽ("& fn1 &")����["&wgdj&"��]{"&wgsm&"}����"&gif&"%%,[<font color=red><b>"&djinfo&"</b></font>]ʹ��%%�����½�:<font color=red>-"& killer &"</font>��,�Լ���%%���书�������:<font color=red>-"& mekill &"</font>��."&gjtx&"<br>"&my_npc_attack&"<br>"&npc_attack
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
		attack="##��%%�ı���:"& Application("aqjh_baowuname") &"���ߣ���õ��˱�������������Ҫ���������ſ��Եõ�����Ķ�����"&gjtx
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		exit function
	end if
if rs("ͨ��")=False then
	conn.execute "update �û� set ɱ��ʱ��=now(),״̬='��',�¼�ԭ��='"&aqjh_name&"|���У�"&fn1&"' where ����='" & to1 &"'"
	attack=xinxi & "##<bgsound src=wav/daipu.wav loop=1>��["& menpai &"]��ʽ("& fn1 &")����["&wgdj&"��]{"&wgsm&"}����"&gif&"%%,%%��������..�����Ӵ�������һλ��Ϻ!"&gjtx&"<br>"&npc_attack
else
	conn.execute "update �û� set ͨ��=False,״̬='��',����=0,����=0,����=0,�书=0,����=0,����=0,���=int(���/2),�¼�ԭ��='"&aqjh_name&"|���У�"&fn1&"' where ����='" & to1 &"'"
	conn.execute "update �û� set ����=����+5000000 where ����='" & aqjh_name &"'"
	attack="����ͨ������"&xinxi & "##<bgsound src=wav/daipu.wav loop=1>��["& menpai &"]��ʽ("& fn1 &")����["&wgdj&"��]{"&wgsm&"}����"&gif&"%%,%%��������..Ϊ�ˣ�##�õ�500��Ԫ�ٸ�����!"&gjtx&"<br>"&npc_attack
end if
call boot(to1,"���У������ߣ�"&aqjh_name&",["&menpai&"]"&fn1)
conn.execute "insert into l(b,a,c,e,d) values ('" & to1 & "',now(),'" & aqjh_name & "','"&fn1&"','����')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
case else
 Response.Write "<script language=JavaScript>{alert('��ʾ���Ҳ�����������');}</script>"
 Response.End
end select
end function
%>