<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="func.ASP"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if aqjh_jhdj<30 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����ȡ�����Ҫ30���ſ��Բ�����');}</script>"
	Response.End
end if
temp=int(abs(clng(Application("aqjh_zdff"))))
if temp<>tl then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������ʲô�����س���������!');</script>"
	response.end
end if
zzfff=Application("aqjh_zzfafang")
if zzfff="" then
	Response.Write "<script language=JavaScript>{alert('����������ʾ��Ŀǰû��վ�����š�');}</script>"
	Response.End 
end if
fffdata=split(zzfff,"|")
ffsl=abs(int(clng(fffdata(0))))
ffsj=fffdata(1)
erase fffdata
nowsj=DateDiff("s",ffsj,now())
if nowsj>=10 then
        Application.Lock
	Application("aqjh_zzfafang")=""
        Application.UnLock
	Response.Write "<script language=JavaScript>{alert('��ʾ��վ��������Чʱ��10�룬�������ϡ�');}</script>"
	Response.End
end if
name=aqjh_name
myip=Request.ServerVariables("REMOTE_ADDR")
if application("aqjh_zzffjl")<>"" then
	ff=split(application("aqjh_zzffjl"),";")
	jsq=ubound(ff)-1
	for i=0 to jsq
		fsz=split(ff(i),"|")
		if fsz(0)=aqjh_name and clng(fsz(1))=hour(time()) then
			erase fsz
			erase ff
			Response.Write "<script language=JavaScript>{alert('��ʾ�����Ѿ��������ˡ�');}</script>"
			Response.End
		end if
		if clng(fsz(1))=hour(time()) and fsz(2)=myip then
			erase fsz
			erase ff
			Response.Write "<script language=JavaScript>{alert('��ʾ���Բ���ͬһIPֻ����ȡһ�ν�ҡ�');}</script>"
			Response.End
		end if
		erase fsz
	next
	erase ff
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����ʱ�� from �û� where ����='"&aqjh_name&"'",conn,2,2
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
xianzaisj=s & ":" & f & ":" & m
tt="<font class=t>(" & xianzaisj & ")</font>"
woczsj=DateDiff("s",rs("����ʱ��"),now())
if  woczsj<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��["&aqjh_name &"]��ղ���������ˣ�������Ҫ�ˣ�');window.close();</script>"
	response.end
end if
newjl=application("aqjh_zzffjl")&aqjh_name&"|"&hour(time())&"|"&myip&";"
application.lock()
application("aqjh_zzffjl")=newjl
application.unlock()
fffff="<marquee width=100% scrollamount=8><font color=green>����ȡ�ȼá�</font><font color=red>"&aqjh_name&"</font>�쵽��վ�����ŵ�<font color=red><b>"&ffsl&"���</b>��վ��֧��������������Ϸ����Ҫ��С��ѽ</font>��</marquee>"&tt
conn.execute "update �û� set ���=���+"&ffsl&",����ʱ��=now()+(31/86400) where ����='"&aqjh_name&"'"
conn.close
set conn=nothing
says=fffff
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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