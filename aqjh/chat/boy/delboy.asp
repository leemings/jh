<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../sjfunc/sjfunc.asp"-->
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * FROM �û� WHERE ����='" & aqjh_name &"'",conn
if rs("boy")="" and rs("boysex")="" then
   	Response.Write "<script Language=Javascript>alert('��ʾ���㲢û��С�����������ΰɣ�');history.back();</script>"
	response.end
end if
if rs("����")<10000 then
	Response.Write "<script Language=Javascript>alert('��ʾ����ĵ������Ѿ��ܵ��ˣ���ͻ���°ɣ�');history.back();</script>"
	response.end
else
 conn.execute "update �û� set ����=����-10000,boy='',boysex='' where ����='" & aqjh_name &"'"
 kl=aqjh_name&"���Լ���Ӥ��������·�ߣ������������ȥ����ʧ����<font color=red>10000</font>��."
 Response.Write "<script Language=Javascript>alert('��������Թ������');history.back();</script>"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>������Թ����</b></font>"&kl
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