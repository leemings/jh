<%
'��ַ�ύ�ж�
sub ChkPost()
	dim server_v1,server_v2
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		Response.Write "<script Language=Javascript>alert('��ʾ���벻Ҫ��վ���ύ���ݣ�');window.close();</script>"
		Response.End
	end if
end sub
'�����ҷ���ģ��
sub lts(actz,towho,saysj,gs,room)
'actz:����ģʽ��towho:�ܻ�����saysj���������ݣ�gs:���Ļ���˽�ģ�room�������
	toname=towho
	if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(towho)&" ")=0 then
		toname="���"
	end if
	Application.Lock
	says=saysj
	says=replace(says,chr(39),"\'")
	says=replace(says,chr(34),"\"&chr(34))
	act=actz
	towhoway=gs
	towho=toname
	addwordcolor="660099"
	saycolor="008888"
	addsays="��"
	saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& room & ");<"&"/script>"
	addmsg saystr
end sub
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