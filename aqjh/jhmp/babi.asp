<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0

aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
id=trim(request.querystring("id"))
if not isnumeric(id) then 
	Response.Write "<script language=JavaScript>{alert('��ʾ��ʲô����ˣ�С���㻹��������ѽ��');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
id=trim(request.querystring("id"))
if InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where id=" & id & " and ������='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�����Ҳ������������ģ���С������ʲô��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
toname=rs("����")
if rs("����2")="��Ƥ"&Month(date()) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('call,������ʲô�����Ѿ��ǹ�Ƥ�ˣ�������ʲô��');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
yudian=rs("mvalue")
conn.Execute "update �û� set ����2='"&"��Ƥ"&Month(date())&"' where id="&id
conn.Execute "update �û� set allvalue=allvalue+"&int(yudian*0.05)&" where ����='"&aqjh_name&"'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>���ǵ㡿��["&aqjh_name&"]��("&toname&")���ϰǵ��ݵ�:"&int(yudian*0.05)&"�㣬Ŭ���ɣ������ˣ�</font>"		'��������
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
Response.Write "<script Language=Javascript>alert('��ʾ��["&aqjh_name&"]��("&toname&")���ϰǵ��ݵ�:"&int(yudian*0.05)&"�㣬Ŭ���ɣ������ˣ�');location.href = 'javascript:history.go(-1)';</script>"
%>
