<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select boy,boysex,w1,��ż,���,ι��ʱ�� from �û� where ����='"&aqjh_name&"'",conn,3,3
myhua=rs("w1")
zz=rs("��ż")
if isnull(rs("boy")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������û��С�����Ͽ�����һ���ɣ�');</script>"
	response.end
end if
sj=DateDiff("s",rs("ι��ʱ��"),now())
if sj<30 then
s=30-sj
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���ӹ���Ҫ30��һ�Σ����"& s &"��,�ɱ����ź��ӣ�');parent.f3.location.href = 'boy.asp';</script>"
response.end
end if
zt=split(rs("boy"),"|")
if UBound(zt)<>7 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ��С�����ݳ��������¹���');</script>"
	response.end
end if

mypic=rs("boysex")
if DateDiff("h",zt(7),now())<5 then
	zg=clng(zt(6))
else
	zg=0
end if
if clng(zt(3))<3000 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ��С����������3000�������������ܴ��ˣ�');parent.f3.location.href = 'boy.asp';</script>"
response.end
end if
if DateDiff("d",zt(2),now())<15 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ�����ӻ�û���󣬲��ܰ�æ׬Ǯ��15����ܴ�׬Ǯ��');parent.f3.location.href = 'boy.asp';</script>"
response.end
end if
if clng(zt(5))<2000 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('��ʾ��С�������ģ����ܴ��ˣ���Ϊ��ĸ�ľ�����ȥ��ɣ�');parent.f3.location.href = 'boy.asp';</script>"
response.end
end if
if zg>=(4+int(DateDiff("d",zt(2),now())/10)) then
	Response.Write "<script Language=Javascript>alert('��ʾ�����Ѿ��չ�"&(4+int(DateDiff("d",zt(2),now())/10))&"���ˣ����5Сʱ�ٲ�����');parent.f3.location.href = 'boy.asp';</script>"
	response.end
end if
if zt(1)="��" then
'����|���|����|����|����|����|�չ�����|�չ�ʱ��
'����|����|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29|2002-6-28 11:04:29|2002-6-28 11:04:29
temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-250)&"|"&(clng(zt(4))-10)&"|"&(clng(zt(5))-10)&"|"&(zg+1)&"|"&now()  
 Response.Write "<script Language=Javascript>alert('��ʾ��С����׬Ǯ��ɣ���:-250 ��:-10 ��:-10 ����������2ö!');parent.f3.location.href = 'boy.asp';</script>"
conn.execute "update �û� set ���=���+2,boy='"&temp&"',ι��ʱ��=now() where  ����='"&aqjh_name&"'"
conn.execute "update �û� set boy='"&temp&"' where  ����='"&zz&"'" 
says="<font color=red><b>�����Ӵ򹤡�</b></font><font color=blue>"&aqjh_name&"<font color=green>�����Ѿ������С˧��<img src=/chat/boy/IMAGES/boy.gif><font color=blue>"&zt(0)&"</font>�Ѿ���Ϊ��ĸ��׬Ǯ�ˣ��ø��ˣ�<font color=blue>"&zt(0)&"</font>�︸ĸ<font color=blue>"&zz&"</font>��<font color=blue>"&aqjh_name&"</font>��׬����<font color=red>2</font>ö��ң�����������250�㣬��Ϊ��ĸ���������ƣ�Ͳ���ʲô~~"
else
'����|���|����|����|����|����|�չ�����|�չ�ʱ��
'����|����|2002-6-28 11:04:29|100|100|100|0|2002-6-28 11:04:29|2002-6-28 11:04:29|2002-6-28 11:04:29
temp=zt(0)&"|"&zt(1)&"|"&zt(2)&"|"&(clng(zt(3))-250)&"|"&(clng(zt(4))-10)&"|"&(clng(zt(5))-100)&"|"&(zg+1)&"|"&now()  
 Response.Write "<script Language=Javascript>alert('��ʾ��С�����ҽ�׬Ǯ����:-250 ��:-10 ��:-100 ����������3ö!');parent.f3.location.href = 'boy.asp';</script>"
conn.execute "update �û� set ���=���+3,boy='"&temp&"',ι��ʱ��=now() where  ����='"&aqjh_name&"'"
conn.execute "update �û� set boy='"&temp&"' where  ����='"&zz&"'" 
says="<font color=red><b>�����ӽ��顿</b></font><font color=blue>"&aqjh_name&"<font color=0088FF>�����Ѿ�ͤͤ������СѾͷ<img src=/chat/boy/IMAGES/GIRL.gif><font color=blue>"&zt(0)&"</font>��Ȼ��ͻ�С����ѧ���Ѿ��ظ��ˣ�<font color=blue>"&zt(0)&"</font>������Сѧѧ����Ӣ��ҽ̣�����׬����<font color=red>3</font>ö��ң�����������250�㣬����Т˳�ĺ��Ӱ�~~"
end if
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