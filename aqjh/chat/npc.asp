<html>

<head>
<meta http-equiv="refresh" content="120;url=npc.asp">
<style type="text/css">
A {COLOR: #ffffff; TEXT-DECORATION: none}
td{font-size:9pt;color:ffffff}
</style>
</head>

<body background=bg.gif oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false onselectstart=event.returnValue=false>
<!--#include file="npc_chat.asp"-->
<!--#include file="../const3.asp"-->
<!--#include file="../showchat.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatinfo=split(aqjh_roominfo(nowinroom),"|")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
allnpc=Application("aqjh_npc")
response.write now()
'�ж�ˢ��ʱ���Ƿ񱻿�
sxsj=DateDiff("n",application("npc_sxsj"),now())
if sxsj>10 then
	application("npc_sxsj")=now()
	application("npc_sxr")=aqjh_name
end if
if application("npc_sxr")=aqjh_name then
application("npc_sxsj")=now()	'npcˢ��ʱ��
'NPC�Զ�����
zdnpc_ok="no"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
if chatinfo(0)=aqjh_chat3 and allnpc="" and Minute(time())<15 then
	Application.Lock
	rs.open "select * from [npc] order by id desc",conn,3,3
        sl=rs.recordcount
        randomize()
        npcid=int(rnd*(sl))
        rs.move npcid
        zdnpc_name=rs("n����")
	if Instr(Application("aqjh_npc"),";" & zdnpc_name & "|")=0 then
        zdnpc_id=rs("id")
        zdnpc_sex=rs("n�Ա�")
        zdnpc_jhtx=rs("nͼƬ")
        zdnpc_nccc=rs("n������")
        zdnpc_j=int(sqr(int(rs("n����")/50)))
        conn.execute  "update npc set n������=n������+1,n����='��',n����='��',n����="&zdnpc_j*100&",n����ʱ��='"&now()&"',n����="&zdnpc_j*150&",n�书="&zdnpc_j*800&",n����="&zdnpc_j*5000&",n����="&zdnpc_j*150&",n����="&zdnpc_j*150000&",n����ʱ��='"&now()&"' where n����='"&zdnpc_name&"'"
        myonline =zdnpc_name& "|" &zdnpc_sex& "|NPC|"&rs("n����")&"|" &zdnpc_jhtx& "|" &zdnpc_j& "|"&zdnpc_id&"|0|1|"&"����"
        rs.close
        call addnpc(zdnpc_name)
        mess="<bgsound src=readonly/cd.mid loop=1><font color=black>��NPC�Զ����롿" & zdnpc_name &"</a></font>"&zdnpc_nccc
	call showchat(mess)
	zdnpc_ok="yes"
	else
		rs.close
	end if
	Application.UnLock
end if
'NPC�Զ�����
if chatinfo(0)=aqjh_chat3 and zdnpc_ok="no" then
 npc_room=split(Application("aqjh_npc"),"|")
 npc_inroom=ubound(npc_room)
 if npc_inroom>=1 then
    randomize timer
    npcr=int(rnd*20)+1
    mekill=50*npcr
    nl=int(rnd*npc_inroom)
    nl1=npc_room(nl)
    nl2=split(nl1,";")
    nl_name=trim(nl2(1))
    select case npcr
    case 1,5,8,18
	npcr2=npcr
	sql="select n���� from npc where n����='"&nl_name&"'"
	rs.open sql,conn,2,2
	if not rs.eof then
		n_zhuren=rs("n����")
	end if
	rs.close
	if n_zhuren<>aqjh_name then
	 sql="select * from �û� where ����='"&aqjh_name&"'"
	 rs.open sql,conn,2,2
	 if rs("����")=false and rs("����")<>"����" and rs("�ȼ�")>18 then
		select case npcr2
		case 1
		mess="<font color=red>��NPC������</font>["&nl_name&"]��[NPC����]��ʽ{�����ջ�}����"&aqjh_name&",ɱ��" &  aqjh_name &"����:<font color=red>-"& mekill&"</font>�㣬" &  aqjh_name &"����ֱ���Ρ���"
		conn.execute "update �û� set ����=����-"& mekill &" where ����='" &  aqjh_name &"'"
		conn.execute  "update npc set n������=n������+1 where n����='"&nl_name&"'"
		case 5
		mess="<font color=red>��NPC��ɱ����</font>["&nl_name&"]��[NPC����]��ʽ{һ����ʽ}����"&aqjh_name &",ɱ��" &  aqjh_name &"����:<font color=red>-"& mekill&"</font>�㣬���" &  aqjh_name &"��ͷת�򡭡�"
		conn.execute "update �û� set ����=����-"& mekill &" where ����='" &  aqjh_name &"'"
		conn.execute  "update npc set n������=n������+1 where n����='"&nl_name&"'"
		case 8
		mess="<font color=red>������������</font>["&nl_name&"]�о�������֧����æʹ����<font color=red>[NPCר����Ч��ҩ]100��</font>,�������ϱ�����<font color=red>1000</font>��,��������<font color=red>800</font>�㡭��"
		conn.execute  "update npc set n����=n����+1000,n����=n����+800,n������=n������+1 where n����='"&nl_name&"'"
		case 18
		mess="<font color=red>��ֹͣ������</font><font color=green>["&nl_name&"]</font>���ú�����������Ϣ��</font>"
		end select
		call showchat(mess)
	 end if
	 rs.close
	end if
    case 3,6,7,9,10,12,14,20
	npcr2=npcr
	select case npcr2
	case 3
		mess="<font color=red>��NPC������</font>NPC����<font color=blue>["&nl_name&"]</font>���������������,û��һ���˹�����,Խ��Խ��!~~"
	case 6
		mess="<font color=red>��NPC�п���</font>NPC����<font color=blue>["&nl_name&"]</font>���������˵����������һ���ݺὭ��,ȴ�в���["&aqjh_name&"]���С��!��"
	case 7
		mess="<font color=red>��NPC���¡�</font>NPC����<font color=blue>["&nl_name&"]</font>������Ҷ��޾���ɵ�,��е�:������,���·���!~~"
	case 9
		mess="<font color=red>��NPC�й䡿</font>NPC����<font color=blue>["&nl_name&"]</font>�ڽ�������һȦ,û���ҵ�һ�����ĵ�����"
	case 10
		mess="<font color=red>��NPC���衿</font>NPC����<font color=blue>["&nl_name&"]</font>ʵ��̫������,����:����������ͷ....."
	case 12
		mess="<font color=red>��NPC������</font><font color=blue>["&nl_name&"]</font>ָ��["&aqjh_name&"]������������С��,�㿪�˱������Ҵ�ѽ,��Ȼ���һ���ӵ����г���,Ҫ���ʹ��������ױ����ã���"
	case 14
		mess="<font color=red>��NPC������</font>NPC����<font color=blue>["&nl_name&"]</font>�о������е���,������,��,����Ҫ������....</font>"
	case 20
		mess="<font color=red>��NPC�����硿</font>NPC����<font color=blue>["&nl_name&"]</font>������["&aqjh_name&"]���Դ����������죬��һ����ȫ�������Ƶĵ��ߣ�����ҡ��ҡͷ���������ֻ�����Ҳ̫ԭʼ�ˣ���"
	end select
	call showchat(mess)
    end select
 end if
end if
end if
'npc�Զ�����
randomize()
rnd1=int(rnd*400)+1
says=""
kill=int(rnd*1000)+400
mekill=kill*10
npc_room=split(Application("aqjh_npc"),"|")
npc_inroom=ubound(npc_room)
if npc_inroom>=1 then
Set conn=Server.CreateObject("ADODB.CONNECTION") 
Set rs=Server.CreateObject("ADODB.RecordSet") 
conn.open Application("aqjh_usermdb")
randomize timer
nl=int(rnd*npc_inroom)
nl1=npc_room(nl)
nl2=split(nl1,";")
nl_name=trim(nl2(1))
sql="select * from npc where n����='"&nl_name&"'"
rs.open sql,conn,2,2
if not rs.eof then
    n_diren=rs("n����")
end if
rs.close
if n_diren=aqjh_name then
sql="select * from �û� where ����='"&aqjh_name&"'"
rs.open sql,conn,2,2
if rs("����")=true and rs("����")<>"����" and rs("�ȼ�")>80 then
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>��NPC�Զ��ÿ���<font color=balck>"&aqjh_name&"��Ϊ�Լ����������˾�û��,<img src=sjfunc/1.gif>������NPC["&nl_name&"]ʹ��һ�Ž���������ɵĽ����"&aqjh_name&"�ı���״̬,����ǳ�ΪNPC["&nl_name&"]���˵��³�,������Ҳ����!</font>"
call showchat(says)
conn.execute  "update �û� set ����=false where ����='"&aqjh_name &"'"
end if
if rs("����")=false and rs("����")<>"����" and rs("�ȼ�")>80 then
rs.close
dz=int(rnd*80)+1
sql="select * from y where id="&dz&""
rs.open sql,conn,2,2
 dz_ok=rs("a")
rs.close
npcdj=int(rnd*5)+1
Select Case npcdj
case 1
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>��NPC�Զ�������<font color=balck>"&aqjh_name&"������������<img src=sjfunc/1.gif>ͻȻ��NPC["&nl_name&"]ʹ��һ��["&dz_ok&"]���˵ò��ᣬ�����½�"&mekill&"�㣬�����½�"&kill&"�㣬����ǳ�ΪNPC["&nl_name&"]���˵��³�!</font>"
case 2
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>��NPC�Զ�������<font color=balck>"&aqjh_name&"������Ϊ��<img src=sjfunc/1.gif>�����ϵ�NPC["&nl_name&"]ͻȻʹ��һ��["&dz_ok&"]���˵ò��ᣬ�����½�"&mekill&"�㣬�����½�"&kill&"�㣬����ǳ�ΪNPC["&nl_name&"]���˵��³�!</font>"
case 3
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>��NPC�Զ�������<font color=balck>"&aqjh_name&"���ڼ�����Ϣ<img src=sjfunc/1.gif>ȴ��NPC["&nl_name&"]ʹ��һ��["&dz_ok&"]���˵ò��ᣬ�����½�"&mekill&"�㣬�����½�"&kill&"�㣬����ǳ�ΪNPC["&nl_name&"]���˵��³�!</font>"
case 4
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>��NPC�Զ�������<font color=balck>"&aqjh_name&"���ڻ�԰�ͻ�<img src=sjfunc/1.gif>ͻȻNPC["&nl_name&"]�ӻ����бĳ�,ʹ��һ��["&dz_ok&"]���˵ò��ᣬ�����½�"&mekill&"�㣬�����½�"&kill&"�㣬����ǳ�ΪNPC["&nl_name&"]���˵��³�!</font>"
case 5
says="<bgsound src=wav/dgjj0.WAV loop=1><font color=red>��NPC�Զ�������<font color=balck>"&aqjh_name&"����ͬ���˷绨ѩ��<img src=sjfunc/1.gif>�����뱻NPC["&nl_name&"]ʹ��һ��["&dz_ok&"]���˵ò��ᣬ�����½�"&mekill&"�㣬�����½�"&kill&"�㣬����ǳ�ΪNPC["&nl_name&"]���˵��³�!</font>"	
end select
call showchat(says)
conn.execute  "update �û� set ����=����-"&mekill&",����=����-"&kill&" where ����='"&aqjh_name &"'"
conn.execute  "update npc set n������=n������+1 where n����='"&nl_name&"'" 
end if
end if
end if
'npc�Զ�����
'NPC�Զ��˳�
if Application("aqjh_npc")<>"" and chatinfo(0)=aqjh_chat3 then
  npc_room=split(Application("aqjh_npc"),"|")
  npc_inroom=ubound(npc_room)
  Set conn=Server.CreateObject("ADODB.CONNECTION")
  Set rs=Server.CreateObject("ADODB.RecordSet")
  conn.open Application("aqjh_usermdb")
  for i=0 to npc_inroom-1
	nl1=npc_room(nl)
	nl2=split(nl1,";")
	nl_name=trim(nl2(1))
        sql="select * from [npc] where n����='"&nl_name&"'"
        rs.open sql,conn,1,1
        sj=DateDiff("n",rs("n����ʱ��"),now())
        if sj>15 then
	call delnpc(nl_name)
	mess= "<bgsound src=wav/gf.wav loop=1><font color=black>��NPC�Զ��˳���</font>NPC����["&nl_name&"]������ȥ��û�ҵ��óԵĶ��������˼�Ȧ��������������ʧ��..."
	call showchat(mess)
        end if
        rs.close
  next
end if
'NPC�Զ��˳�
%>