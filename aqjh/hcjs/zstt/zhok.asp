<%@ LANGUAGE=VBScript%>

<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
jifen=rs("allvalue")
if rs("ת��")<3  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����ʹ������ܣ�����ת��3�Σ�');location.href = 'javascript:history.go(-1)';}</script>"
		response.end
end if
if rs("����")<50000000  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����ʹ������ܣ�������5ǧ���ֽ�');location.href = 'javascript:history.go(-1)';}</script>"
		response.end
end if
if rs("���")<10  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����ʹ������ܣ�����������5����ң�');location.href = 'javascript:history.go(-1)';}</script>"
		response.end
end if
if rs("����")<200000  then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('��ʾ����ʹ������ܣ������з���200000 ��');location.href = 'javascript:history.go(-1)';}</script>"
		response.end
end if
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<300 then
	ss=300-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('300����һ�Σ������ţ��ٵ�"&ss&"��ɣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
neili=int(jifen*.25)
tili=int(jifen*.25)
wg=int(jifen*.05)
conn.execute "update �û� set ����=����+10000,����=����+10000,������=������+100,������=������+200,����=����-50000000,���=���-10,����=����-200000,����ʱ��=now()  where ����='"&aqjh_name&"'"
rs.close
set rs=nothing	
set conn=nothing
says="<font color=#ff0000><b>������ܡ����齭��" & aqjh_name & "</b>������ʹ������ܳɹ������ӹ���10000�ͷ���10000������������������100��200��,����Ŭ���ɣ���ת���˾��Ǻÿ���������</font>"			'��������
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
Response.Write "<script language=JavaScript>{alert('��ʾ�������ʩչ��ϣ������������������');location.href = 'javascript:history.go(-1)';}</script>"
response.end
%>
