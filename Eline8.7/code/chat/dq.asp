<%@ LANGUAGE=VBScript codepage ="936" %>
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
if Instr(LCase(Application("sjjh_useronlinename"&session("nowinroom")))," "&LCase(sjjh_name)&" ")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ��������������Ҳſ��Բ�����');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if InStr(Application("sjjh_diuqi"),"|")=0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ���������ˡ���');}</script>"
	Response.End
end if
zt=split(Application("sjjh_diuqi"),"|")
'2����
if ubound(zt)=2 then
	if zt(2)=sjjh_name then
		Response.Write "<script language=JavaScript>{alert('��ʾ���Լ�����������ҲҪ����');}</script>"
		Response.End 
	end if
elseif ubound(zt)=4 then
	if zt(4)=sjjh_name then
		Response.Write "<script language=JavaScript>{alert('��ʾ���Լ���������ƷҲҪ����');}</script>"
		Response.End 
	end if
end if
Application.Lock
	Application("sjjh_diuqi")=""
Application.UnLock
if left(zt(0),1)<>"w" and left(zt(0),1)<>"v" then
	if not isnumeric(zt(0)) then 
		Response.Write "<script language=JavaScript>{alert('��ʾ�����������ⲻ���Բ�����');}</script>"
		Response.End 
	end if
	yin=int(abs(clng(zt(0))))
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("sjjh_usermdb")
	conn.execute "update �û� set ����=����+"&yin&" where  ����='"&sjjh_name&"'"
	conn.close
	set conn=nothing
	diuqi="<font color=red><b>��������Ϣ��</b></font>##�������ˣ��õ���<font color=blue>["&zt(2)&"]</font>����������[<font color=red><b>"&yin&"</b></font>]������"
else
if left(zt(0),1)="v" then
	if not isnumeric(zt(2)) then 
		Response.Write "<script language=JavaScript>{alert('��ʾ�����������ⲻ���Բ�����');}</script>"
		Response.End 
	end if
	value=int(abs(clng(zt(2))))
	name=zt(1)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	conn.open Application("sjjh_usermdb")
	conn.execute "update �û� set allvalue=allvalue+"&value&" where  ����='"&sjjh_name&"'"
	diuqi="<font color=red><b>��������Ϣ��</b></font>##�������ˣ�["&name&"]���������ң�û�б���Ĵ�㣺</b>"&value&"</b>�㣬��##�����ˡ�</font>]������"
	conn.close
	set conn=nothing
else
	if not isnumeric(zt(2)) then 
		Response.Write "<script language=JavaScript>{alert('��ʾ�����������ⲻ���Բ�����');}</script>"
		Response.End 
	end if
	wusl=int(abs(clng(zt(2))))
	wpname=zt(1)
	lb=zt(0)
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("sjjh_usermdb")
	rs.open "select "&lb&" from �û� where ����='"&sjjh_name&"'",conn,3,3
	zstemp=add(rs(lb),wpname,wusl)
	conn.execute "update �û� set "&lb&"='"&zstemp&"' where  ����='"&sjjh_name&"'"
	diuqi="<font color=red><b>��������Ϣ��</b></font>##�������ˣ��õ���<font color=blue>["&zt(4)&"]</font>��������Ʒ[<font color=red><b>"&wpname&"����"&wusl&"��</b></font>]������"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
end if
zj="<a href=javascript:parent.sw('[" & sjjh_name & "]'); target=f2>"& sjjh_name & "</a>"
br="<a href=javascript:parent.sw('[" & name & "]'); target=f2>" & name & "</a>"
diuqi=Replace(diuqi,"##",zj,1,3,1)
diuqi=Replace(diuqi,"%%",br,1,3,1)
says=diuqi
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
