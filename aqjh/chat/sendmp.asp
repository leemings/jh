<%@ LANGUAGE=VBScript codepage ="936" %><script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
<%Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=Trim(Request.QueryString("id"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select ����ʱ��,����,grade,���,��ż from �û� where ����='" & aqjh_name &"'",conn
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<30 then
	s=30-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ����Ҫ�����ŵ��˷�Ǯ���["&s&"]�������������ܷ�Ǯ��[����/����]�����š����ŷ���[�ɷ�]�������ų��ϣ�');window.close();</script>"
	Response.End
end if
conn.Execute "update �û� set ����ʱ��=now() where ����='" & aqjh_name &"'"
fmp=rs("����")
fgrade=rs("grade")
fsf=rs("���")
fpo=rs("��ż")
rs.close
smoney=0
if (fgrade>=5 and fsf="����") or (fgrade=4 and fsf="����") then
	smoney=1
end if
if fpo<>"��" then
rs.open "select ����,grade,���,��ż from �û� where ����='"& fpo &"'",conn
	if rs("����")=fmp and rs("grade")=5 and rs("���")="����" then
		smoney=1
	end if
rs.close
end if
if smoney=0 then
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��������˼�������ݲ��ܷ�Ǯ[����/����]�����ܷ�Ǯ��[����/����]�����š����ŷ���[�ɷ�]�������ų��ϣ�');window.close();</script>"
	Response.End
end if
useronlinename=Application("aqjh_useronlinename"&nowinroom)
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
randomize timer
if id=1 then
	s=int(rnd*5000)
	diaox="<bgsound src=wav/zt.mid loop=1><font color=green>��"&fmp&"�书ָ����</font>"& aqjh_name &"<B>���Ը������ұ����ɵ���ָ���书�������еĵ�������/�������ӣ�<font color=#ff0000>"& s &"�㣡</font></b>"
else
	s=int(rnd*200000)
	diaox="<bgsound src=wav/zt.mid loop=1><font color=green>��"&fmp&"��Ǯ��</font>"& aqjh_name &"<B>���Լ����ɵ���ÿһ���˷��Ž���������<font color=#ff0000>"& s &"��<img src=img/251.gif>��</font></b>"
end if

for i=0 to x
if id=1 then
	conn.Execute "update �û� set ����=����+" & s & ",����=����+" & s & " where ����='" & online(i) & "' and ����='"& fmp &"'"
else
	conn.Execute "update �û� set ����=����+" & s & " where ����='" & online(i) & "' and ����='"& fmp &"'"
end if
next
conn.close
set conn=nothing
says=diaox

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
Response.Redirect "../ok.asp?id=705"
%>