<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../const3.asp"-->
<%Response.Buffer=true
Response.Expires = 0
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
chatroomsn=session("nowinroom")
mychatroomsn=Request.QueryString("roomsn")
chatroomname=Request.QueryString("chatroomname")
chatroomname=Application("aqjh_chatroomname"&mychatroomsn)
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
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t
if chatroomsn=mychatroomsn then
  response.write "<Script language=javascript>alert('���Ѿ���["&chatroomname&"]�����ظ�����!');parent.m.location.reload();parent.r.location.reload();</script>"
  response.end
end if
chatroominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(chatroominfo)-1
i=mychatroomsn
online=split(trim(Application("aqjh_useronlinename"&i))," ")
onlinenum=ubound(online)+1
sj_chat_info=split(chatroominfo(i),"|")
num=int(sj_chat_info(1))
if onlinenum>=num then
  response.write "<Script>alert('["&chatroomname&"]���䵱ǰ�������������ܽ��룡');parent.r.location.reload();</script>"
  response.end
end if
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0 and chatroomname<>"�ݵ�ר��" and aqjh_grade<10 then
   response.write "<Script>alert('��ʾ:�ݵ���,�޷�����!');parent.r.location.reload();</script>"
   response.end
end if
'��ս�����ж�
if application("aqjh_user")<>aqjh_name then
if chatroomname=aqjh_chat2 and (Weekday(date())<>7 or Hour(time())<>21 or (Hour(time())=21 and Minute(time())>50)) then
	response.write "<Script>alert('��ս����ֻ����ÿ������21:00-21:50�ֲſɽ��� \n����ʱ��"&time()&"��ʱ��δ�������Եȣ�');parent.r.location.reload();</script>"
	response.end
end if
if chatroomname=aqjh_chat2 and aqjh_grade>5 then
	response.write "<Script>alert('��ʾ���ٸ����˲��ò��룡');parent.r.location.reload();</script>"
	response.end
end if
end if
if chatroomname=aqjh_chat2 and Instr(Application("aqjh_admin_send"),"|" & aqjh_name & "|")<>0 then
	response.write "<Script>alert('��ʾ���㶼�ǲ����ˣ��Ͳ��������ˣ�');parent.r.location.reload();</script>"
	response.end
end if
'��ս�����ж�END
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,�Ա�,����,���,����ͷ��,��Ա�ȼ�,����ʱ��,ͨ��,ת��,��ż from �û� where ����='" & aqjh_name &"'",conn,2,2
aqjh_id=rs("id")
jhsf=rs("���")
if Instr(Application("aqjh_guibin"),"|" & aqjh_name & "|")<>0 then
 jhsf="���"
end if
if Instr(Application("aqjh_admin_send"),"|" & aqjh_name & "|")<>0 then 
 jhsf="����"
end if
if rs("��ż")=Application("aqjh_user") and rs("�Ա�")="Ů" then
 jhsf="վ������"
end if
if Application("aqjh_mengzhu")=aqjh_name then 
 jhsf="��������"
end if
if rs("ͨ��")=True then
	jhmp="ͨ����"
else
	jhmp=rs("����")
end if
jhtx=rs("����ͷ��")
sex=rs("�Ա�")
hydj=rs("��Ա�ȼ�")
myzs=rs("ת��")
mypeiou=rs("��ż")
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<1 and aqjh_grade<6 then
	s=1-sj
	Response.Write "<script language=JavaScript>{alert('��ʾ��ת���������["&s&"����]�ٲ�����');parent.m.location.reload();parent.r.location.reload();}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if cstr(sj_chat_info(2))=1 and aqjh_grade<7 then
	rs.close
	rs.open "select id from �û� where ����='" & aqjh_name &"'"&" and "&sj_chat_info(4),conn
	if rs.eof or rs.bof  then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<html><head><meta http-equiv='Content-Type' content='text/html; charset=gb2312'><meta http-equiv='pragma' content='no-cache'></head><body bgcolor=" + Application("aqjh_chatbgcolor") + " background=" + Application("aqjh_chatimage") + " bgproperties=fixed>"
		Response.Write "<script language=JavaScript>{alert('����["&sj_chat_info(0)&"]�������ǣ�"&sj_chat_info(3)&"');parent.r.location.reload();}</script>"
		Response.End
	end if
	rs.close
end if
'����,�Ա�,����,���,ͷ��,�ȼ�,id
myzanli=0
if Instr(LCase(application("aqjh_zanli")),LCase("!"&aqjh_name&"!"))>0 then myzanli=1
myonline = aqjh_name & "|" & sex & "|" & jhmp & "|" & jhsf & "|" & jhtx & "|" & aqjh_jhdj& "|" & aqjh_id& "|" & hydj&"|"&myzanli&"|"&"����"&"|"&mypeiou&"|"&myzs
if sj_chat_info(7)<>0 or jhmp="����" then
	conn.Execute "update �û� set ����=false,����ʱ��=now() where ����='" & aqjh_name &"'"	
else
	conn.Execute "update �û� set ����=true,����ʱ��=now() where ����='" & aqjh_name &"'"	
end if
set rs=nothing	
conn.close
set conn=nothing
'�˳�ԭ����
Application.Lock
mychatroomname=Application("aqjh_chatroomname"&chatroomsn)
onlinelist=Application("aqjh_onlinelist"&chatroomsn)
onno=ubound(onlinelist)
for i=1 to onno
 if InStr(onlinelist(i),aqjh_name & "|")=1 then
'myonline=onlinelist(i)
  for j=i to onno-1
   onlinelist(j)=onlinelist(j+1)
  next
  ReDim Preserve onlinelist(onno-1)
  Application("aqjh_onlinelist"&chatroomsn)=onlinelist
  exit for
 end if
next
Application("aqjh_useronlinename"&chatroomsn)=Replace(Application("aqjh_useronlinename"&chatroomsn)," " & aqjh_name & " ","")
'�����·���
dim newonlinelist()
js=1
onlinelist=Application("aqjh_onlinelist"&mychatroomsn)
onlineno=ubound(onlinelist)
yjl=0
for i=1 to onlineno
onuser=split(onlinelist(i),"|")
'if yjl=0 and StrComp(onuser(2),jhmp,1)=1 then		'���������ֺ���ƴ������
'if yjl=0 and len(onuser(2))<len(jhmp) then			'���������ֳ��ȣ�������ǰ
'if yjl=0 and len(onuser(2))>len(jhmp) then			'���������ֳ���,�̵���ǰ
if yjl=0 and StrComp(onuser(0),aqjh_name,1)=1 then	'�����ֺ���ƴ������
'if yjl=0 and len(onuser(0))<len(aqjh_name) then	'�����ֳ��ȣ�������ǰ
'if yjl=0 and len(onuser(0))>len(aqjh_name) then	'�����ֳ���,�̵���ǰ
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
Application("aqjh_onlinelist"&mychatroomsn)=newonlinelist
Application("aqjh_useronlinename"&mychatroomsn)=Application("aqjh_useronlinename"&mychatroomsn)& " "&aqjh_name & " "
erase  newonlinelist
erase  onlinelist
' onlinelist=Application("aqjh_onlinelist"&mychatroomsn)
' onlineno=ubound(onlinelist)+1
' ReDim Preserve onlinelist(onlineno)
'onlinelist(onlineno)=myonline
'Application("aqjh_onlinelist"&mychatroomsn)=onlinelist
'Application("aqjh_useronlinename"&mychatroomsn)=Application("aqjh_useronlinename"&mychatroomsn)& " "&aqjh_name & " "
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
says="<bgsound src=wav/folder.wav loop=1><font color=black>��������</font><font color=#009933><font color=red>" & aqjh_name & "</font>ʩչ�����貨΢�����Ṧ��ת�ۼ��ӡ�" & mychatroomname & "����ʧ�ˣ�ԭ����ȥ��" &chatroomname & "���ˡ�</font><font class=t>(" & time() & ")</font>"
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &",0," & chatroomsn & ");<"&"/script>"
addmsg saystr
'session("SayCount")=Application("SayCount")
act="����"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
says="<bgsound src=wav/folder.wav loop=1><font color=black>��������</font><font color=#009933><a href=javascript:parent.sw('[" & aqjh_name & "]'); target=f2>" & aqjh_name & "</a>ʩչ�����貨΢�����Ṧ��ת�ۼ�ӡ�" & mychatroomname & "�������ˡ�" &chatroomname & "����</font><font class=t>(" & t & ")</font><bgsound src='readonly/cd.mid' loop='1'>"	'�ܻ���
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
saystr="<"&"script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &",0," & session("nowinroom") & ");<"&"/script>"
addmsg saystr
For i = 1 to Application("SayCount")-Session("SayCount")
	Response.Write Application("SayStr"&YuShu((Session("SayCount")+i)))
Next
session("SayCount")=Application("SayCount")
%>
<Script >
parent.crm='<%=Application("aqjh_chatroomname"&session("nowinroom"))%>';
parent.myroom=<%=session("nowinroom")%>
parent.f2.document.af.mdsx.checked=true;
parent.f2.document.af.sytemp.focus();
parent.f2.document.af.towho.value="���";
parent.m.location.reload();
//parent.f2.location.href='f2.asp?id=1';
parent.f2.location.reload();
parent.qp();
alert('��ӭ������<%=chatroomname%>����');
</script>