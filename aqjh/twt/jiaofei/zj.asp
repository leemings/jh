<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if aqjh_grade<7  then
	Response.Write "<script Language=Javascript>alert('��ʾ�����ǹٸ����ˣ������Բ���!');window.close();</script>"
	response.end
end if
if Application("aqjh_dalie")="����" then
	Response.Write "<script Language=Javascript>alert('��ʾ�����˻����أ�');window.close();</script>"
	response.end
end if
Application.Lock
Application("aqjh_dalie")="����" 
Application.UnLock
says="<bgsound src=wav/tufei.wav loop=1><font color=red>����Ϣ��</font><font color=red>ͻȻһȺ����<img src=img/jiao.gif>���뽭�����٣�������ǿ�ȥ�˷˰���</font>"			'��������
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
Response.Write "<script Language=Javascript>alert('��ʾ���ٸ�������ɹ���');window.close();</script>"
response.end
%>