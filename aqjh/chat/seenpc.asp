<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const3.asp"-->
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(aqjh_name)&" ")=0 then ydl=0
	if ydl=1 and clng(nowinroom)<>i then 
		Session.Abandon
		Response.Redirect "../error.asp?id=140"
		Response.End 
	end if
        aqjh_roominfo=split(Application("aqjh_room"),";")
        chatinfo=split(aqjh_roominfo(nowinroom),"|")
        if chatinfo(0)<>aqjh_chat3 then
	Response.Write "<script language=JavaScript>{alert('��ʾ:NPCֻ����["&aqjh_chat3&"]����Ч��');window.close();}</script>"
	Response.End
end if
if aqjh_ifnpc=0 then
        Response.Write "<script language=JavaScript>{alert('��ʾ:NPC��ʱ��ϵͳ�رգ�');window.close();}</script>"
	Response.End
end if
next 
'����
chatroomname=trim(Application("aqjh_chatroomname"&session("nowinroom")))
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
'ȡ��ǰ����npc����
npcsl=ubound(split(Application("aqjh_npc"),";"))
npcsl_inroom=ubound(split(Application("aqjh_npc"&nowinroom),";"))
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
id=clng(Request.QueryString("id"))
act=Request.QueryString("act")
Response.Write "<script>window.moveTo(100,30);document.onmousedown=click;}</script>"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
jbmoney=1
if act<>"" then
	'��ս�ж�
	act=clng(act)
	if act<>1 and act<>2 and act<>3 then
		act=1
	end if
	if session("aqjh_grade")<application("aqjh_npcff") and act>1 then
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('��ʾ��" & aqjh_name & "��û�з���NPC��Ȩ��!');location.href = 'seenpc.asp?id=" & id & "';</script>"
		Response.end
	end if
	select case act
	case 1'����npcʹ�õ��ǽ���������������أ�ÿһ���������һ����Ҵ�ҿ��Ը�jbmoney=1���Լ��ı���
		'ȡ��npc�ȼ����㻽������Ҫ�Ľ����
		if npcsl>=Application("aqjh_maxnpc") Then 
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ��" & aqjh_name & "�����涨����npc�����������["& Application("aqjh_maxnpc") &"]��!');location.href = 'seenpc.asp?id=" & id & "';</script>"
			Response.end
		end if
		rs.open "Select  * from [npc] where id=" & id,conn,1,1
		If rs.bof  Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ�������ҵ�NPC��������!');window.close();</script>"
			Response.end
		end if
		jbsl=(int(sqr(int(rs("n����")/50)))*jbmoney)
		name=rs("n����")
		if InStr(";" & Application("aqjh_npc"), ";" & name & "|")<>0 then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ����NPC�����Ѿ������ˣ��㲻�ܽ��л��Ѳ���!');location.href = 'seenpc.asp?id=" & id & "';</script>"
			Response.end
		end if
		zuren=0
		says="<font color=#ff0000>��NPC��Ϣ��</font>["&aqjh_name&"]�ڶ�ħ���ҵ���<a href=javascript:parent.sw(\'[" & name & "]\'); target=f2>" & name &"</a>���ý�Ǯ��"&name&"���������Լ���С��,����"&name&"�İ���["&aqjh_name&"]ʵ������!</font>��"
		if rs("n����")=session("aqjh_name") then
			jbsl=int(jbsl/2+1)
			says="<font color=#ff0000>��NPC��Ϣ��</font>["&aqjh_name&"]�ڶ�ħ�������ҵ����Լ����ϲ���<a href=javascript:parent.sw(\'[" &name & "]\'); target=f2>" & name &"</a>������"&name&"�İ���["&aqjh_name&"]ʵ������!</font>��"
		end if
		rs.close
		'�ж��Լ��Ľ������,������npc���ҵģ��һ�һ��Ǯ�ٳ���
			rs.open "Select * from [�û�] where ����='"& session("aqjh_name") &"'",conn,1,3
			if rs("���")<jbsl then
				rs.close
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('��ʾ����û���㹻�Ľ�ͱң����ܻ��Ѵ�NPC!');location.href = 'seenpc.asp?id=" & id & "';</script>"
				Response.end
			end if
                        if rs("ְҵ")<>"ħ��ʦ" then
				rs.close
				set rs=nothing
				conn.close
				set conn=nothing
				Response.Write "<script Language=Javascript>alert('��ʾ��ֻ��ħ��ʦ���ܻ��Ѵ�NPC!');location.href = 'seenpc.asp?id=" & id & "';</script>"
				Response.end
			end if
			rs("���")=rs("���")-jbsl
			rs.update
			rs.close
		rs.open "Select * from [npc] where id=" & id,conn,1,3
		rs("n����")=session("aqjh_name")
                id=rs("id")
		name=rs("n����")
                sex=rs("n�Ա�")
                jhtx=rs("nͼƬ")
                j=int(sqr(int(rs("n����")/50)))
                rs("n����")=j*100
                rs("n����")=j*150
                rs("n�书")=j*800
                rs("n����")=j*5000
                rs("n����")=j*150
                rs("n����")=j*150000
                rs("n������")=rs("n������")+1
                rs("n����ʱ��")=now()
                rs("n����ʱ��")=now()
                myonline = name & "|" & sex & "|NPC|"&rs("n����")&"|" & jhtx & "|" &j& "|"&id&"|0|1|"&"����"
		rs("n����")="��"
		rs.update
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
'�л��ɹ����������
Application.Lock
js=1
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")
if yjl=0 and StrComp(onuser(0),aqjh_name,1)=1 then	'�����ֺ���ƴ������
	Redim Preserve newonlinelist(js+1)
	yjl=1
	newonlinelist(js)=myonline
	newonlinelist(js+1)=onlinelist(i)
	js=js+2
else
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=onlinelist(i)
	js=js+1
end if
next 
if yjl=0 then
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=myonline
end if
Application("aqjh_onlinelist"&nowinroom)=newonlinelist
Application("aqjh_useronlinename"&nowinroom)=Application("aqjh_useronlinename"&nowinroom)&" "&name&" "
Application("aqjh_npc"&nowinroom)=Application("aqjh_npc"&nowinroom)&";"&name&"|"
Application("aqjh_npc")=Application("aqjh_npc")&";"&name&"|"
erase  newonlinelist
erase  onlinelist
Application.UnLock
Session("SayCount")=Application("SayCount")
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
	Application.Lock()
		Application("SayCount")=Application("SayCount")+1
		i="SayStr"&YuShu(Application("SayCount"))
		Application(i)=Str
	Application.UnLock()
End Sub
'�л��ɹ����������
call showchat(says)
		Response.Write "<script Language=Javascript>alert('��ʾ���ɹ��г�NPC�����������������ˣ������\n�˲�սʱ����������!');location.href = 'seenpc.asp?id=" & id & "';</script>"
		Response.end
	case 2,3
		rs.open "Select  * from [npc] where id=" & id,conn,2,3
		If rs.bof  Then
			rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script Language=Javascript>alert('��ʾ�������ҵ�NPC��������!');window.close();</script>"
			Response.end
		end if
                id=rs("id")
		name=rs("n����")
                sex=rs("n�Ա�")
                jhtx=rs("nͼƬ")
                nccc=rs("n������")
                j=int(sqr(int(rs("n����")/50)))
                rs("n����")=j*100
                rs("n����")=j*150
                rs("n�书")=j*800
                rs("n����")=j*5000
                rs("n����")=j*150
                rs("n����")=j*150000
                if act=2 then rs("n������")=rs("n������")+1
                rs("n����")="��"
                rs("n����")="��"
                rs("n����ʱ��")=now()
                rs("n����ʱ��")=now()
		rs.update
                myonline = name & "|" & sex & "|NPC|"&rs("n����")&"|" & jhtx & "|" &j& "|"&id&"|0|1|"&"����"
		rs.close
		if act=2 then
			if InStr(";" & Application("aqjh_npc"), ";" & name & "|")<>0 then
			Response.Write "<script Language=Javascript>alert('��ʾ����NPC�Ѿ������ˣ������ٷų�!');location.href = 'seenpc.asp?id=" & id & "';</script>"
			Response.end
			end if
                        npc_lasttime=DateDiff("n",Application("aqjh_npc_lasttime"),now())
                        if Application("aqjh_npc_lasttime")<>"" and npc_lasttime<30 and aqjh_name<>Application("aqjh_user") then
			Response.Write "<script Language=Javascript>alert('��ʾ���ٸ�ÿ��Сʱֻ�ܷ���һ��NPC��վ���������ƣ�');location.href = 'seenpc.asp?id=" & id & "';</script>"
			Response.end
                        end if
                        Application("aqjh_npc_lasttime")=now()
           says = "<bgsound src=readonly/cd.mid loop=1><font color=black>��NPC���빫�桿<a href=javascript:parent.sw(\'[" & name & "]\'); target=f2>" & name &"</a></font>"&nccc
          show="NPC���ųɹ�!"
'�������
Application.Lock
js=1
onlinelist=Application("aqjh_onlinelist"&nowinroom)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")
if yjl=0 and StrComp(onuser(0),aqjh_name,1)=1 then	'�����ֺ���ƴ������
	Redim Preserve newonlinelist(js+1)
	yjl=1
	newonlinelist(js)=myonline
	newonlinelist(js+1)=onlinelist(i)
	js=js+2
else
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=onlinelist(i)
	js=js+1
end if
next 
if yjl=0 then
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=myonline
end if
Application("aqjh_onlinelist"&nowinroom)=newonlinelist
Application("aqjh_useronlinename"&nowinroom)=Application("aqjh_useronlinename"&nowinroom)&" "&name&" "
Application("aqjh_npc"&nowinroom)=Application("aqjh_npc"&nowinroom)&";"&name&"|"
Application("aqjh_npc")=Application("aqjh_npc")&";"&name&"|"
erase  newonlinelist
erase  onlinelist
Application.UnLock
Session("SayCount")=Application("SayCount")
Function Yushu(a)
	Yushu=(a and 31)
End Function
Sub AddMsg(Str)
	Application.Lock()
		Application("SayCount")=Application("SayCount")+1
		i="SayStr"&YuShu(Application("SayCount"))
		Application(i)=Str
	Application.UnLock()
End Sub
'�������
call showchat(says)
		else 
                       if InStr(";" & Application("aqjh_npc"), ";" & name & "|")=0 then
				Response.Write "<script Language=Javascript>alert('��ʾ����NPC�������㲻���߳���!');location.href = 'seenpc.asp?id=" & id & "';</script>"
				Response.end
			end if
'ɾ������
nowinroom=session("nowinroom")
Application.Lock
onlinelist=Application("aqjh_onlinelist"&session("nowinroom"))
onno=ubound(onlinelist)
for i=1 to onno
if InStr(onlinelist(i),name & "|")=1 then
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  exit for
 end if
next
Application("aqjh_onlinelist"&nowinroom)=onlinelist
Application("aqjh_useronlinename"&nowinroom)=Replace(Application("aqjh_useronlinename"&nowinroom)," " & name & " ","")
Application("aqjh_npc"&nowinroom)=Replace(Application("aqjh_npc"&nowinroom),";"&name&"|","")
Application("aqjh_npc")=Replace(Application("aqjh_npc"),";"&name&"|","")
Application.UnLock
            says = "<bgsound src=wav/gf.wav loop=1><font color=black>���߳�NPC��</font>����Ա["&aqjh_name&"]��NPC����["&name&"]�������߳�������"&name&"�˳��˽���!"
            show="NPC�߳��ɹ�!"
call showchat(says)
		end if
		Response.Write "<script Language=Javascript>alert('��ʾ��"& show &"');location.href = 'seenpc.asp?id=" & id & "';</script>"
		Response.end
	end select
end if
rs.open "Select  * from [npc] where id=" & id,conn,1,1
If rs.bof  Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�������ҵ�NPC��������!');window.close();</script>"
	Response.end
end if
name=rs("n����")
%>
<html>
<head>
<title>�鿴NPC[<%=name%>]����ϸ״̬</title>
<LINK href="jh.css" rel=stylesheet>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body background="bg.gif" text="#FFFF00" leftmargin="0" topmargin="4" oncontextmenu=window.event.returnValue=false>
<table border="1" width="99%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#00215a" align="center">
<tr align="center">
<td colspan="4">�鿴NPC[<%=name%>]����ϸ״̬
<div id="Layer1" style="position:absolute; width:87px; height:73px; z-index:1; left: 404px; top: 370px;"><img src="<%=rs("nͼƬ")%>"></div>
</td>
</tr>
<tr>
<td width="25%" align="left">���ƣ�<%=name%></td>
<td width="25%" align="left">�Ա�<%=rs("n�Ա�")%></td>
<td width="25%" align="left">�ȼ���<%=int(sqr(int(rs("n����")/50)))%></td>
<td width="25%" align="left">������<%=rs("n������")%>��</td>
</tr>
<tr>
<td align="left">������<%=rs("n����")%></td>
<td align="left">������<%=rs("n����")%></td>
<td align="left">�书��<%=rs("n�书")%></td>
<td align="left">������<%=rs("n����")%></td>
</tr>
<tr>
<td align="left">������<%=rs("n����")%></td>
<td align="left">������<%=rs("n����")%></td>
<td align="left">�����ʣ�<%=rs("n������")%></td>
<td align="left">�����ʣ�<%=rs("n������")%></td>
</tr>
<tr>
<td align="left">��:<%=rs("n����ʱ��")%></td>
<td align="left">��:<%=rs("n����ʱ��")%></td>
<td align="left">���ˣ�<%=rs("n����")%></td>
<td align="left"><%online=InStr(";" & Application("aqjh_npc"), ";" & name & "|")
If online<>0 Then%>��NPC����!<%else%>��NPC������<%end if%></td>
</tr>
<tr>
<td colspan="4" align="center">
<input type="submit" name="Submit" onclick="javascript:<%If online<>0 Then%>alert('��NPC�Ѿ����߲����ٻ���!');<%else%>if(confirm('���Ѵ�NPC��Ҫ��ң�<%=(int(sqr(int(rs("n����")/50)))*jbmoney)%>�������������������ֻ�軨һ���Ǯ \n����������������ˣ���ȷ�����л��Ѳ�����')){location.href='seenpc.asp?id=<%=id%>&act=1';}<%end if%>" value="����NPC">&nbsp;&nbsp;&nbsp;
<%if session("aqjh_grade")>=application("aqjh_npcff") or Session("aqjh_name") = Application("aqjh_user") then%>
<input type="submit" name="Submit2" onclick="javascript:<%If online<>0 Then%>alert('��NPC�Ѿ����߲����ٷ���!');<%else%>location.href='seenpc.asp?id=<%=id%>&act=2';<%end if%>" value="����NPC">&nbsp;&nbsp;&nbsp;
<input type="submit" name="Submit3" onclick="javascript:<%If online=0 Then%>alert('��NPC�������߲������߳�!');<%else%>location.href='seenpc.asp?id=<%=id%>&act=3';<%end if%>" value="�߳�NPC">
<%end if%>
</td>
</tr>
<tr>
<td colspan="4" align="left">
<p><strong>1.NPC�������Ʒ���㡣</strong><br>
<font color="#FFFFFF">�ڴ���NPC����Ʒ�ı���ֵ��NPC��������ɱ���ߵȼ���NPC�ȼ����м��㣺(NPC�ȼ�/ɱ���ߵȼ�)*1000Ȼ��ȡ���������ô�ֵ��aqjh_npcwp���бȽϣ�����˼����ֵ���ڵ���aqjh_npcwp��ֵ�������Ʒ�����������˳���ֹ���������Ҫ��Ϊ�˴���߼���Ա��ͼ�NPC�ı��������ͨ���˱Ƚ���ƷҲ����%100�����ģ������NPC���Ĵ����뵱ǰ��Ʒֵ���м�������ó���
�� <br>
aqjh_npcwp���ڵ�ֵ��:<strong><font color="#00FF00"><%=application("aqjh_npcwp")%></font></strong>
������:70��NPC��65��.(65/70)x1000=928�����������<strong><font color="#00FF00"><%=application("aqjh_npcwp")%></font></strong>,�����Ҵ����npc������Ʒ����.С����û��!<br>
</font> <strong>2.NPC�Զ���������.</strong><font color="#FFFFFF"><br>
        NPC������ڽ�����ѡ��һ���˽��й��������������ڱ����У������ǳ����ˣ��ٸ����ˣ����еȼ�<21���ģ���ֹͣ���������򱻹������˽���ʧȥ����500,����200,����NPC���㽫�����ĵ���,���˲��������Լ���NPC,NPC�е��˺���ô�����ǹ������˵ģ�ֻ�����ĵ��˲��ܴ�!NPC�����Ѻ�������״̬��ָ����!<br>
</font><strong>3.������NPC�������!</strong><font color="#FFFFFF"><br>
ͬʱ����ٻ���NPC������<strong><font color="#00FF00"><%=application("aqjh_maxnpc")%></font></strong>��<br>
NPC��Ʒ�������������棺<strong><font color="#00FF00"><%=application("aqjh_npcwp")%></font></strong>ֵ<br>
������ټ��Ŀ����ֶ����ż��߳�NPC��<strong><font color="#00FF00"><%=application("aqjh_npcff")%></font></strong>��<br>
<br>
</font>���������Խ���NPC��ʲô���顢�뷨��Ҫ��������������<br>
��̳��ַ��<a href=http://bbs.7758530.com target=_blank><font color=yellow>http://www.7758530.com</font></a><br>��ϵ���׵��� QQ:865240608</p>
</td>
</tr>
</table>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<%
'��������Ϣ
Sub showchat(mess)
says=mess   '��������
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway&","&nowinroom&");parent.f2.document.af.npc.value='"&Application("aqjh_npc")&"';parent.m.location.reload();<"&"/script>"
addmsg saystr
end sub
Function Yushu1(a)
 Yushu1=(a and 31)
End Function
Sub AddMsg(Str)
Application.Lock()
Application("SayCount")=Application("SayCount")+1
i="SayStr"&YuShu1(Application("SayCount"))
Application(i)=Str
Application.UnLock()
End Sub
%>