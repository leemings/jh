<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
nowinroom=session("nowinroom")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
	Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');}</script>" 
	Response.End 
end if 
name=LCase(trim(request.querystring("name")))
to1=LCase(trim(request.querystring("toname")))
fn1=abs(int(request.querystring("fn1")))
m1=abs(int(request.querystring("m1")))
m2=abs(int(request.querystring("m2")))
m3=abs(int(request.querystring("m3")))
for each element in request.querystring
if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('��ʾ���������������⣬��鿴��');</script>"
		Response.End
end if
next
if to1<>sjjh_name then
		Response.Write "<script Language=Javascript>alert('��ʾ���˼�Ҳ����Ҫ����Զ����ʲôȤ��');</script>"
		Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ���� from �û� where ����='"&sjjh_name&"'",conn,2,2

if rs("����")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���Զ�����Ӧ����10��С��300��');}</script>"
	Response.End
end if
rs.close
rs.open "select ���� from �û� where ����='"&name&"'",conn
if rs("����")<fn1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ��["&name&"]����������󲻹�ѽ����ô���˼Ҷ�??��');}</script>"
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
		temp="<font color=green>���Զ�������</font>[<font color=red>"&sjjh_name&"</font>]ҡ��:<img src='../jhimg/"&m11&".gif'><img src='../jhimg/"&m22&".gif'><img src='../jhimg/"&m33&".gif'>"&myds&"��,{<font color=red>"&name&"</font>}ҡ��:<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>"&tods&"�㣬<font color=red><b>Ӯ��</b></font>"&fn1&"��!�ٺ�~~~������˼��[<font color=red>"&sjjh_name&"</font>]��лл�ֵܵĺ񰮣����ó���!!<font color=green>�ٸ��ն�˰5%!���˵���������������!</font>"
	
		conn.execute "update �û� set ����=����+" & int((fn1-1)*0.95) & ",����=����-20 where ����='"&sjjh_name&"'"		
		conn.execute "update �û� set ����=����-" & fn1 & ",����=����-10 where ����='"&name&"'"		
	else
		temp="<font color=green>���Զ�������</font>[<font color=red>"&sjjh_name&"</font>]ҡ��:<img src='../jhimg/"&m11&".gif'><img src='../jhimg/"&m22&".gif'><img src='../jhimg/"&m33&".gif'>"&myds&"��,{<font color=red>"&name&"</font>}ҡ��:<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>"&tods&"�㣬<font color=red><b>����</b></font>"&fn1&"��!��?��ô��������?����~~~��,û��ϵ�������,[<font color=red>"&sjjh_name&"</font>]���´�һ��Ӯ��!!!<font color=green>�ٸ��ն�˰5%!���˵���������������!</font>"
	
		conn.execute "update �û� set ����=����-" & fn1 & ",����=����-10 where ����='"&sjjh_name&"'"
		conn.execute "update �û� set ����=����+" & int((fn1-1)*0.95) & ",����=����-20 where ����='"&name&"'"
	end if
else
	temp="<font color=green>���Զ�������</font>[<font color=red>"&sjjh_name&"</font>]ҡ��:<img src='../jhimg/"&m11&".gif'><img src='../jhimg/"&m22&".gif'><img src='../jhimg/"&m33&".gif'>"&myds&"��,{<font color=red>"&name&"</font>}ҡ��:<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>"&tods&"�㣬<font color=red><b>ƽ��</b></font>!ΰ���˼�����ǲ�ı���ϡ���"
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