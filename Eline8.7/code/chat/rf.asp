<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="sjfunc/func.asp"-->
<!--#include file="../mywp.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
if Application("sjjh_rf")=0 then
	Response.Write "<script Language=Javascript>alert('��ʾ����~~�����������������');</script>"
	response.end
end if
tl=int(abs(clng(Request.QueryString("tl"))))
tempjs=int(abs(clng(Application("sjjh_rf"))))
if tempjs<>tl then
	Response.Write "<script Language=Javascript>alert('��ʾ�����������ʲô�����س���������!');</script>"
	response.end
end if
Application.Lock
Application("sjjh_rf")=0
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
conn.execute "update �û� set ����=����-"&tempjs&" where ����='" & sjjh_name &"'"
rs.open "select ����,w6 FROM �û� WHERE ����='" & sjjh_name &"'",conn
if rs("����")<-100 then
	conn.execute "update �û� set ״̬='��',�¼�ԭ��='��ʬ|"&fn1&"' where ����='" & sjjh_name &"'"
	conn.execute "insert into l(b,a,c,e,d) values ('" & sjjh_name & "',now(),'��ʬ','���´��ο�������','����')"
	rf="<img src='img/kl.gif'>̫�����ˣ�["&sjjh_name&"]����û����ɣ��˷���(ren fanzi)������Ҳ��ɱ���ң��ٺ�~������"
	call boot(sjjh_name,"�˷��������ߣ��˷���ץ��ѽ������")
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
zstemp=add(rs("w6"),js(myxy),1)
conn.execute "update �û� set ����=����+"& tempjs*40 &",����=����+200,w6='"&zstemp&"' where ����='" & sjjh_name & "'"
rf="������ɣ��˷���(ren fanzi)�������Ȼɱ���ң��ѵ��㲻��֪����Щ���ӵ�������ô���ٸ�����<img src='img/251.gif'>"&tempjs*55&"������������200�㣬����"&sjjh_name&"�õ���[<b><font color=red>"&js(myxy)&"</font></b>]1������"

end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>����������</b></font>"&rf

says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& nowinroom & ");<"&"/script>"
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