<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
	Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');}</script>" 
	Response.End 
end if 
name=LCase(trim(Request.QueryString("name")))
toname=LCase(trim(Request.QueryString("toname")))
fn1=abs(int(Request.QueryString("fn1")))
m1=abs(int(Request.QueryString("m1")))
m2=abs(int(Request.QueryString("m2")))
m3=abs(int(Request.QueryString("m3")))
for each element in Request.QueryString
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.QueryString(element),"'")<>0 or instr(Request.QueryString(element)," ")<>0 or instr(Request.QueryString(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��鿴��');</script>"
		Response.End
end if
next
if toname<>aqjh_name then
		Response.Write "<script Language=Javascript>alert('��ʾ���˼�Ҳ����Ҫ����Զ����ʲôȤ��');</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select �Ṧ from �û� where ����='"&aqjh_name&"'",conn,2,2
if rs("�Ṧ")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ������Ṧ���񲻹�ѽ��');}</script>"
	Response.End
end if
rs.close
rs.open "select �Ṧ from �û� where ����='"&name&"'",conn,2,2
if rs("�Ṧ")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&name&"]����û����ô����Ṧ��');}</script>"
	Response.End
end if
randomize()
m11 = Int(6 * Rnd + 1)
m22 = Int(6 * Rnd + 1)
m33 = Int(6 * Rnd + 1)
myds=m11+m22+m33
tods=m1+m2+m3
if myds<>tods then
	if myds>tods then
		temp="<font color=green>���Զ��Ṧ��</font>[<font color=red>"&aqjh_name&"</font>]ҡ��:<img src='../jhimg/"&m11&".gif'><img src='../jhimg/"&m22&".gif'><img src='../jhimg/"&m33&".gif'>"&myds&"��,{<font color=red>"&name&"</font>}ҡ��:<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>"&tods&"�㣬<font color=red><b>Ӯ��</b></font>"&fn1&"��!������˼��лл�ֵܵĺ񰮣�"
		conn.execute "update �û� set �Ṧ=�Ṧ+" & (fn1-1) & " where ����='"&aqjh_name&"'"		
		conn.execute "update �û� set �Ṧ=�Ṧ-" & fn1 & " where ����='"&name&"'"		
	else
		temp="<font color=green>���Զ��Ṧ��</font>[<font color=red>"&aqjh_name&"</font>]ҡ��:<img src='../jhimg/"&m11&".gif'><img src='../jhimg/"&m22&".gif'><img src='../jhimg/"&m33&".gif'>"&myds&"��,{<font color=red>"&name&"</font>}ҡ��:<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>"&tods&"�㣬<font color=red><b>����</b></font>"&fn1&"��!û��ϵ���´�һ��Ӯ�㣡"
		conn.execute "update �û� set �Ṧ=�Ṧ-" & fn1 & " where ����='"&aqjh_name&"'"
		conn.execute "update �û� set �Ṧ=�Ṧ+" & (fn1-1) & " where ����='"&name&"'"
	end if
else
temp="<font color=green>���Զ��Ṧ��</font>[<font color=red>"&aqjh_name&"</font>]ҡ��:<img src='../jhimg/"&m11&".gif'><img src='../jhimg/"&m22&".gif'><img src='../jhimg/"&m33&".gif'>"&myds&"��,{<font color=red>"&name&"</font>}ҡ��:<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>"&tods&"�㣬<font color=red><b>ƽ��</b></font>!ΰ���˼�����ǲ�ı���ϡ���"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red><b>��������Ϣ��</b></font>"&temp		'��������
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
