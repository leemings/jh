<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../pass.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
name=trim(request.form("name"))
pass=trim(request.form("pass"))
for each element in Request.Form
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��鿴��');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
end if
next
if name="" or pass="" then
	Response.Write "<script Language=Javascript>alert('��ʾ���ǲ����뿪���ˣ��������Ϳ����������');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
aqjh_roominfo=split(Application("aqjh_room"),";")
chatroomnum=ubound(aqjh_roominfo)-1
for i=0 to chatroomnum	
	ydl=1
	if Instr(LCase(Application("aqjh_useronlinename"&i))," "&LCase(name)&" ")=0 then ydl=0
	if ydl=1 then exit for
next 
if ydl=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ڸ�ʲôѽ���㲢û�п��ڽ���������ǲ���ѡ����ˣ�');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
inroom=i
pass=md5(pass)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,���� from �û� where ����='" & name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������������������Ҳ���!\n��鿴������������룡');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id=rs("id")
if trim(pass)<>rs("����") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ�����벻��ѽ��������ô�죡');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ�ͨ�������ǵ�������');window.close();</script>"
'aqjh_name=name
userip=Request.ServerVariables("REMOTE_ADDR")
'kickname=name
'kickwhy="�Ҳ�С�ߵ����ˣ�û�취��"
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
'*************************
Application.Lock
onlinelist=Application("aqjh_onlinelist"&inroom)
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
Application("aqjh_onlinelist"&inroom)=onlinelist
Application("aqjh_useronlinename"&inroom)=Replace(Application("aqjh_useronlinename"&inroom)," " & name & " ","")
Application.UnLock
'****************
says="<font color=black>�����ߡ�</font><font color=blue>����ID["& id &"]�����Լ����Լ���һ�ţ�Ϊʲô�������ˡ���</font><font class=t>(" & t & ")</font>"			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & "ϵͳ����" & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ",0);<"&"/script>"
addmsg saystr
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
%>