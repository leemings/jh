<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<!--#include file="../sjfunc/func.asp"-->
<!--#include file="../../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Application("aqjh_kl")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����С���ǲ���û������ѽ�������й���');</script>"
	response.end
end if
id=LCase(trim(Request.QueryString("id")))
tsk=int(abs(clng(Request.QueryString("tsk"))))
tempjs=int(abs(clng(Application("aqjh_kl"))))
if tempjs<>tsk then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������ʲô�����س���������!');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM �û� where ״̬='��ʬ��' and ���='��ʬ��' order by ���� ",conn
sf=rs("���")
attackname=rs("����")
tl=rs("����")
if aqjh_grade>1 then
	Response.Write "<script language=JavaScript>{alert('��ʾ������2�����������ʸ��ʬ��');window.close();}</script>"
	Response.End 
end if
Application.Lock
Application("aqjh_kl")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
conn.execute "update �û� set ����=����-"&tempjs&" where ����='" & aqjh_name &"'"
rs.open "select ����,w6 FROM �û� WHERE ����='" & aqjh_name &"'",conn
if rs("����")<-100000 then
	conn.execute "update �û� set ״̬='��ʬ��',���='��ʬ��',�¼�ԭ��='��ʬ��|"&fn1&"' where ����='" & aqjh_name &"'"
    conn.execute "update �û� set ����=5000,����=5000,�¼�ԭ��='��',״̬='����',���='����' where ����='" & attackname &"'"
  	conn.execute "insert into l(b,a,c,e,d) values ('" & aqjh_name & "',now(),'��ʬ��','���´�ʬ������','����')"
	kl="<img src='pic/kl.gif'>["&aqjh_name&"]̫˥�ˣ�Ϊ�˸���ʬ��" & attackname & "��������Լ���������ʬ��"&id&"ҧ��,["&aqjh_name&"]����˽�ʬ��." & attackname & "��ʽ�����λ��Ϻ��ϲ����" & attackname & "�����ڽ�ʬ��þ���..��������"
	call boot(aqjh_name,"��ʬ���������ߣ���ʬ��������ѽ������")
else
dim js(10)
js(0) ="��ͷ"
js(1) ="����ʯ"
js(2) ="�챦ʯ"
js(3) ="��ʬѪ"
js(4) ="��ʬ��"
js(5) ="ˮ��ʯ"
js(6) ="�̱�ʯ"
js(7) ="���"
js(8) ="����"
js(9)="ľ̿"
randomize()
myxy = Int(Rnd*10)
zstemp=add(rs("w6"),js(myxy),2)
conn.execute "update �û� set ����=����+"& tempjs*1 &",w6='"&zstemp&"' where ����='" & aqjh_name & "'"
kl="��call,Ӣ�ۣ�����Ӣ�ۣ�["&aqjh_name&"]�뽩ʬ��" & attackname & "�����֡�����ʬ�����ڵ����ˣ������Լ�Ҳ�˵Ĳ��ᣬ�ٸ�����<img src='img/251.gif'>"&tempjs*1&"�����ڴ�ܽ�ʬ����"&aqjh_name&"�õ���[<b><font color=red>"&js(myxy)&"</font></b>]2������"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>����ʬ����</b></font>"&kl
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