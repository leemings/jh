<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
st=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if st="" then Response.Redirect "../../error.asp?id=440"
if Application("aqjh_klt4")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ�������Ѿ������ˣ�');</script>"
	Response.end
end if
tempjs=int(abs(clng(Application("aqjh_klt4"))))
'if tempjs<>r then
	'Response.Write "<script Language=Javascript>alert('��ʾ�����������ʲô�����س���������!');</script>"
	'response.end
'end if
Application.Lock
Application("aqjh_klt4")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update �û� set ����=����-"&tempjs&" where ����='"&st&"'"
rs.open "select ���� FROM �û� WHERE ����='"&st&"'",conn
if rs("����")<0 then
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='����|"&fn1&"' where ����='"&st&"'"
conn.execute "insert into l(b,a,c,e,d) values ('" & st & "',now(),'����','������������','����')"
	kl="<img src='pic/kl.gif'>̫�����ˣ�["&st&"]Ϊ�˵�������Ϯ��������ȡ�壬��������Զ�������˲ʺ�֮�£��������ֻ��ǧ�������ڱ���������â�Ĵ�˵������"
	Response.Write "<script Language=Javascript>alert('��ʾ����Ӣ�¿ɼΣ���������̫��������С����ȥ����ɣ�û���ͱ��Ҵ���');</script>"
	call boot(st,"��������ߣ����������")
else
	conn.execute "update �û� set ֪��=֪��+"& tempjs *10&" where ����='" & st & "'"
	kl="<font color=#0000FF>"&st&"</font>�ܲ���,��������ʱʤ����,��������,"&st&"��ţ��Ϊ�����������ؽ���֪��"&tempjs*10&"�㣡"
	end if
	rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>��������</b></font>"&kl	
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
