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
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
you=trim(request("you"))
if instr(you,"OR")<>0 or instr(you,"Or")<>0 or instr(you,"oR")<>0 or instr(you,"or")<>0 or instr(you,">")<>0 or instr(you,"<")<>0 or instr(you,"=")<>0 or instr(you,chr(13))<>0 or instr(you,"'")<>0 or instr(you," ")<>0 or instr(you,"%20")<>0 or instr(you,chr(13))<>0 then
	Response.Write "<script language=JavaScript>{alert('���ؾ��棬��Ҫ����');window.close();}</script>"
	Response.End
end if
if you=aqjh_name then
	Response.Write "<script language=javascript>alert('�������Ѹ����ѳ����ɵģ���');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT ���� FROM �û� WHERE ���='����' and ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('��ʾ���㲻������������ţ���');window.close();</script>"
	response.end
end if
rs.close
rs.open "select ����,��� from �û� where ����='" & aqjh_name &"'",conn
if rs("����")="�ٸ�" and rs("���")="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���ٸ������Ų�����Ҫ�߳����ɵĶ��󣡣���');}</script>"
	Response.End
end if
mp=rs("����")
rs.close
rs.open "SELECT * FROM �û� WHERE ����='"&you&"' and ����='"&mp&"'" ,conn
if rs.eof or rs.bof then	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('��ʾ��["&you&"]�����������ɵ��ӣ�����');window.close();</script>"
	response.end
end if
message="�ɹ��İ�" & you & "���" & id & "�������ĵ��ӻ��ǲ��յĺã�"
conn.execute "update p set c=c-1 where a='" & mp & "'"
conn.execute "update �û� set ���ɻ���=0,����='����',���='����',grade=1 where ����='" & you & "'"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=#ff0000>��������ӡ�" & aqjh_name & "���Լ����ӣ�"& you &"�ϳ����Լ�������"&mp&"��</font>"		'��������
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
%>
<html>
<head>
<title>��<%=Application("aqjh_chatroomname")%>��</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body background="../jhimg/bk_hc3w.gif" text="#000000" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center"><font color="#FF0000" size="+3">���ӹ���</font><br>
<br>
������ʾ��<%=message%> </div>
<hr>
<div align="center">��Ȩ���С����ֽ�����վ��</div> 
</body>
</html>