<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../config.asp"-->
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.Expires = 0
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
chatroomsn=session("nowinroom")	'��ǰ���ڷ������
mychatroomsn=Request.QueryString("roomsn")	'Ҫȥ�ķ������
chatroomname=Application("sjjh_chatroomname"&mychatroomsn)
mm=Application("sjjh_chatroomname"&chatroomsn)
if chatroomsn=mychatroomsn then
  response.write "<Script language=javascript>alert('���Ѿ��ڡ�"&chatroomname&"�������ظ����룡');parent.m.location.reload();</script>"
  response.end
end if
if (sjz>72020 and sjz<72050) and mm="����E��" and sjjh_grade<6 then
	response.write "<Script language=javascript>alert('��Ȼ�Ѿ��߽���"&chatroomname&"������������·���ˣ��������˳������ң�');parent.m.location.reload();</script>"
	response.end
end if
	n=Year(date())
	y=Month(date())
	r=Day(date())
	s=Hour(time())
	f=Minute(time())
	m=Second(time())
	weekdate=weekday(date())
	sjz=weekdate*10000+s*100+f
	'��������20:00��Ϊ6*10000+20*100+0=62000
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t

chatroominfo=split(Application("sjjh_room"),";")
chatroomnum=ubound(chatroominfo)-1
i=mychatroomsn
online=split(trim(Application("sjjh_useronlinename"&i))," ")	'ȡ������Ա����
onlinenum=ubound(online)+1
sj_chat_info=split(chatroominfo(i),"|")	'����Ҫȥ�ķ�������
if sj_chat_info(0)="����E��" and sjjh_grade<>10 then
	if (sjz>71200 and sjz<72020) or sjz>72030 then
		erase sj_chat_info
		erase online
		erase chatroominfo
		response.write "<Script language=javascript>alert('�����˶ᱦ֮�գ�ֻ��������20:20-20:30֮����룡');parent.m.location.reload();</script>"
		response.end
	else
		if sjz>71200 and sjz<72020 then
			erase chatroominfo
			erase sj_chat_info
			erase online
			response.write "<Script language=javascript>alert('�����˶ᱦ֮�գ�����ʱ��Ϊ����20:20��20:30�������ȡ��ʼʱ��Ϊ20:30����');parent.m.location.reload();</script>"
			response.end
		else
			if sjz>72030 then
				erase sj_chat_info
				erase online
				erase chatroominfo
				response.write "<Script language=javascript>alert('�ѹ��˶ᱦ����ʱ�䣬��ʱ��Ϊÿ��������20:20��20:30���ᱦ������ʼʱ��Ϊ20:00�����´μ���������Ӵ��');parent.m.location.reload();</script>"
				response.end
			end if
		end if
	end if
end if
if (sjz>=72020 and sjz<=72030) and (sjjh_grade>=6 and sjjh_name<>"һ����") and chatroomname="����E��" then
		Response.Write "<script language=JavaScript>{alert('��ʾ���ٸ���Ա���ɲ���ᱦ��');}</script>"
		Response.End
	end if

num=int(sj_chat_info(1))
if onlinenum>=num then
	erase chatroominfo
	erase sj_chat_info
	erase online
	response.write "<Script>alert('��"&chatroomname&"�����䵱ǰ�������������ܽ��룡');;</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
if cstr(sj_chat_info(2))=1 then
	if sjjh_grade<7 then
		rs.open "select id,����,�ȼ�,���,����,������,������,����ͷ��,�Ա�,��Ա�ȼ�,����ʱ��,w1 from �û� where ����='" & sjjh_name &"'"&" and "&sj_chat_info(4),conn,2,2
	else
		rs.open "select id,����,�ȼ�,���,����,������,������,����ͷ��,�Ա�,��Ա�ȼ�,����ʱ��,w1 from �û� where ����='" & sjjh_name &"'",conn,2,2
	end if
else
	rs.open "select id,����,�ȼ�,���,����,������,������,����ͷ��,�Ա�,��Ա�ȼ�,����ʱ��,w1 from �û� where ����='" & sjjh_name &"'",conn,2,2
end if
if cstr(sj_chat_info(2))=1 and sjjh_grade<7 then
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><meta http-equiv='pragma' content='no-cache'></head><body bgcolor=" + Application("sjjh_chatbgcolor") + " background=" + Application("sjjh_chatimage") + " bgproperties=fixed>"
		Response.Write "<script language=JavaScript>{alert('���롺"&sj_chat_info(0)&"���������ǣ�"&sj_chat_info(3)&"');}</script>"
		erase chatroominfo
		erase sj_chat_info
		erase online
		Response.End
	end if
end if
sjjh_id=rs("id")
jhsf=rs("���")
jhmp=rs("����")
jhtx=rs("����ͷ��")
sex=rs("�Ա�")
hydj=rs("��Ա�ȼ�")
sj=DateDiff("n",rs("����ʱ��"),now())
tlsx=rs("�ȼ�")*sjjh_tlsx+5260+rs("������")
nlsx=rs("�ȼ�")*sjjh_nlsx+2000+rs("������")
w1=rs("w1")
rs.close
if sj<3 and sjjh_grade<6 and sj_chat_info(0)<>"����E��" then
	s=3-sj
	set rs=nothing	
	conn.close
	set conn=nothing
	erase chatroominfo
	erase sj_chat_info
	erase online
	Response.Write "<script language=JavaScript>{alert('��ʾ��ת���������["&s&"����]�ٲ�����');parent.m.location.reload();}</script>"
	Response.End
end if
if sjz>=72020 and sjz<=72030 and jhmp="����" and chatroomname="����E��" then
        Response.Write "<script language=JavaScript>{alert('��ʾ�������˲��ɲ���ᱦ��');parent.m.location.reload();}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if (sjz>72020 and sjz<72030) and sjjh_grade<6 and sj_chat_info(0)="����E��" then
	a=mywpsl(w1,"ǧ���˲�")
	b=mywpsl(w1,"������֥")
	butl=a*5000
	bunl=b*5000
	tlyj=int(tlsx/5000)
	nlyj=int(nlsx/5000)
	if butl>tlsx or bunl>nlsx then
		set rs=nothing	
		conn.close
		set conn=nothing
		erase chatroominfo
		erase sj_chat_info
		erase online
		Response.Write "<script language=JavaScript>{alert('�μӶᱦ����ʱֻ���������һ������������ǧ���˲μ�������֥�����ҩ̫�ֻ࣬�ܴ�["&tlyj&"]��ǧ���˲Σ�["&nlyj&"]��������֥��');parent.m.location.reload();}</script>"
		Response.End
	end if
end if
'����,�Ա�,����,���,ͷ��,�ȼ�,id
myzanli=0
if Instr(LCase(application("sjjh_zanli")),LCase("!"&sjjh_name&"!"))>0 then myzanli=1
if myzanli=1 and (sj_chat_info(0)="����E��" or sj_chat_info(0)="E�߽���") then
	set rs=nothing	
	conn.close
	set conn=nothing
	erase chatroominfo
	erase sj_chat_info
	erase online
	Response.Write "<script language=JavaScript>{alert('��ʾ�����������У������Խ���˷��䣡');parent.m.location.reload();}</script>"
	Response.End
end if
myonline = sjjh_name & "|" & sex & "|" & jhmp & "|" & jhsf & "|" & jhtx & "|" & sjjh_jhdj& "|" & sjjh_id& "|" & hydj&"|"&myzanli&"|"&"����"
if sj_chat_info(7)<>0 then
	conn.Execute "update �û� set ����=false,����ʱ��=now() where ����='" & sjjh_name &"'"
else
	conn.Execute "update �û� set ����=true,����ʱ��=now() where ����='" & sjjh_name &"'"
end if
set rs=nothing	
conn.close
set conn=nothing
'�˳�ԭ����
Application.Lock
mychatroomname=Application("sjjh_chatroomname"&chatroomsn)
onlinelist=Application("sjjh_onlinelist"&chatroomsn)
onno=ubound(onlinelist)
for i=1 to onno
 if InStr(onlinelist(i),sjjh_name & "|")=1 then
	  for j=i to onno-1
		onlinelist(j)=onlinelist(j+1)
	  next
	  ReDim Preserve onlinelist(onno-1)
	  Application("sjjh_onlinelist"&chatroomsn)=onlinelist
	  exit for
 end if
next
Application("sjjh_useronlinename"&chatroomsn)=Replace(Application("sjjh_useronlinename"&chatroomsn)," " & sjjh_name & " ","")
'�����·���
dim newonlinelist()
js=1
onlinelist=Application("sjjh_onlinelist"&mychatroomsn)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")
'if yjl=0 and StrComp(onuser(2),jhmp,1)=1 then		'���������ֺ���ƴ������
'if yjl=0 and len(onuser(2))<len(jhmp) then			'���������ֳ��ȣ�������ǰ
'if yjl=0 and len(onuser(2))>len(jhmp) then			'���������ֳ���,�̵���ǰ
if yjl=0 and StrComp(onuser(0),sjjh_name,1)=1 then	'�����ֺ���ƴ������
'if yjl=0 and len(onuser(0))<len(sjjh_name) then	'�����ֳ��ȣ�������ǰ
'if yjl=0 and len(onuser(0))>len(sjjh_name) then	'�����ֳ���,�̵���ǰ
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
if yjl=0  then
	Redim Preserve newonlinelist(js)
	newonlinelist(js)=myonline
end if
Application("sjjh_onlinelist"&mychatroomsn)=newonlinelist
Application("sjjh_useronlinename"&mychatroomsn)=Application("sjjh_useronlinename"&mychatroomsn)& " "&sjjh_name & " "
erase  newonlinelist
erase  onlinelist
session("nowinroom")=mychatroomsn
Application.UnLock	

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
act="�˳�"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
if (sjz>72020 and sjz<72030) and sj_chat_info(0)="����E��" then
	says="<font color=black>��������</font><font color=#009933>һ��һ�εĶᱦʱ�����ڵ��ˣ�<font color=red>" & sjjh_name & "</font>�������ŵش�<font color=red>��" & mychatroomname & "��</font>����������<font color=red>��" &chatroomname & "��</font>���䣬�����Ǵ��ף�����˰ɣ�</font><font class=t>(" & time() & ")</font><bgsound src='readonly/cd.mid' loop='1'>"
else
	says="<font color=black>��������</font><font color=#009933><font color=red>" & sjjh_name & "</font>ʩչ�����貨΢�����Ṧ��ת�ۼ���<font color=red>��" & mychatroomname & "��</font>��ʧ�ˣ�ԭ����ȥ<font color=red>��" &chatroomname & "��</font>�ˡ�</font><font class=t>(" & time() & ")</font><bgsound src='readonly/cd.mid' loop='1'>"
end if
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &",0," & chatroomsn & ");<"&"/script>"
addmsg saystr
'session("SayCount")=Application("SayCount")
act="����"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
if (sjz>72020 and sjz<72030) and sj_chat_info(0)="����E��" then
	says="<font color=#cc0000>�����ߡ�</font><font color=#009933>һ��һ�εĶᱦʱ�����ڵ��ˣ�<a href=javascript:parent.sw('��" & sjjh_name & "��'); target=f2>" & sjjh_name & "</a>�������ŵش�<font color=red>��" & mychatroomname & "��</font>����������<font color=red>��" &chatroomname & "��</font>���䣬�����Ǵ��ף�����˰ɣ�</font><font class=t>(" & t & ")</font><bgsound src='readonly/cdcd.wav' loop='1'>"	'�ܻ���
else
	says="<font color=#cc0000>�����ߡ�</font><font color=#009933><a href=javascript:parent.sw('��" & sjjh_name & "��'); target=f2>" & sjjh_name & "</a>ʩչ�����貨΢�����Ṧ��ת�ۼ�ӡ�" & mychatroomname & "�������ˡ�" &chatroomname & "����</font><font class=t>(" & t & ")</font><bgsound src='readonly/cdcd.wav' loop='1'>"	'�ܻ���
end if

says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))

saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &",0," & session("nowinroom") & ");<"&"/script>"
addmsg saystr
For i = 1 to Application("SayCount")-Session("SayCount")
	Response.Write Application("SayStr"&YuShu((Session("SayCount")+i)))
Next
session("SayCount")=Application("SayCount")
%>
<Script >
parent.crm='<%=Application("sjjh_chatroomname"&session("nowinroom"))%>';
parent.myroom=<%=session("nowinroom")%>
parent.m.location.reload();
parent.f2.location.href='f2.asp?id=1';
alert('��ӭ������<%=chatroomname%>��');
</script>