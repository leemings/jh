<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../../mywp.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then 
Response.Write "<script language=javascript>{alert('��ʾ����û�е�½���Ѿ���ʱ�Ͽ����ӣ������µ�½��');parent.history.go(-1);}</script>" 
Response.End 
end if
http = Request.ServerVariables("HTTP_REFERER")
if InStr(http,"mhmj/syl.asp")=0 then 
Response.Write "<script language=javascript>{alert('�ύ���ݷǷ�������ܾ�ִ�У�');parent.history.go(-1);}</script>" 
Response.End 
end if

if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ���㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
gwm="��ħ֮��"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb") 
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
nl=rs("����")
fl=rs("����")
hydj=rs("��Ա�ȼ�")
select case hydj
	case 0
		jgsj=320
	case 1
		jgsj=260
	case 2
		jgsj=220
	case 3
		jgsj=180
	case 4
		jgsj=130
	case 5
		jgsj=100
	case 6
		jgsj=60
end select
sj=DateDiff("s",rs("����ʱ��"),now())
if sj<jgsj then
	s=jgsj-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������"&hydj&"����Ա��"&jgsj&"��ɽ���һ�Σ����"&s&"��������Գ���');parent.history.go(-1);}</script>" 
	Response.End
end if
if fl<110000 then 
	Response.Write "<script Language=Javascript>alert('��ʾ���ٻ�������Ҫ110000�㷨����');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select * from ���� where ����='"&gwm&"'",conn
zt=rs("״̬")
if zt="����" then
	Response.Write "<script Language=Javascript>alert('��ʾ���ù����ѱ����ѣ��㲻��Ҫ�ظ�������');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
conn.execute "update ���� set ����=18000000,ɱ��=8000000,����=7000000,����=5000000,�ȼ�=100,״̬='����'  where ����='"&gwm&"'"
conn.execute "update �û� set ����=����-110000 where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=RED><b>�����ѹ��</b></font><font color=green>[<font color=red><b>"& aqjh_name &"</b></font>]�ķ�[<font color=red><b>110000</b></font>]�㷨���ɹ��Ľ��Ѿ�������[<font color=red><b>"&gwm&"</b></font>]���ѣ����������ֽ���һ��Ѫ���ȷ磡��֪���ֽ��ж��ٽ�����ʿ��ǻ�Ұ������</font>"			'��������
says=replace(says,"'","\'")
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
<script Language=Javascript>alert('��ʾ��<%=aqjh_name%>���������дʣ��ķ�110000�㷨����<%=gwm%>�ɹ����ѣ�');location.href = 'javascript:history.go(-1)';</script>
