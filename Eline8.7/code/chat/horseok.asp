<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
if sjjh_name="" then Response.Redirect "../error.asp?id=440"
sjjh_roominfo=split(Application("sjjh_room"),";")
chatinfo=split(sjjh_roominfo(session("nowinroom")),"|")
if chatinfo(9)<>0 then
	Response.Write "<script language=JavaScript>{alert('��ʾ����["&chatinfo(0)&"]���䲻�ܹ����������');}</script>"
	Response.End
end if
money=int(clng(Request.form("money")))
radiobutton=Request.form("radiobutton")
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"chat")=0 then 
Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');parent.history.go(-1);}</script>" 
Response.End 
end if 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "select ��� FROM �û� WHERE ����='" & sjjh_name &"'",conn,2,2
if rs("���")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ����Ǯ���񲻹���');}</script>"
	Response.End
end if
rs.close
rs.open "select top 1 * FROM g WHERE c='��' and a='"& sjjh_name &"' and f=true",conn,2,2
if not(rs.eof) or not(rs.bof) then
	temp="��ʾ��"&sjjh_name&"������["&rs("b")&"]����ƣ�"&rs("d")&"�������ſ�ʼ��!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('"&temp&"');}</script>"
	Response.End 
end if
rs.close
rs.open "select top 1 id,a FROM g WHERE c='��' and f=false",conn,2,2
tempdu=0
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('��ʾ���������ڿ�ʼ�ˣ���һ���~~!');}</script>"
	tempdu=1
	rs.close
end if
if 	tempdu=0 then
	id=rs("id")
	conn.execute "update �û� set ���=���-"&money&" where ����='" & sjjh_name &"'"
	conn.execute "update g set a='"&sjjh_name&"',f=true,d="&money&",b='"&radiobutton&"' where id="&id&" and f=false"
	rs.close
end if
tmprs=conn.execute("Select count(*) As ���� from g where c='��' and f=true")
dars=tmprs("����")
'��ʼ����
if dars>=10 then
	randomize(timer())
	mmm = Int(4 * Rnd + 1)
	yinma=""
	rs.open "select * FROM g WHERE c='��' and b='"&mmm&"' and f=true",conn,2,2
	do while not rs.bof and not rs.eof
		yinma=yinma&"["&rs("a")&" "& rs("d") &"] "
		conn.execute "update �û� set ���=���+"&rs("d")*3&",����=����+500,����=����+500 where ����='"& rs("a") &"'"
		rs.movenext
	loop
	conn.execute "update g set f=false where f=true and c='��'"
	says="<font color=blue><b>[������ʼ]</b></font>���ż��ҵ�������ʼ�ˡ�ֻ��5ƥ��ͷ������ٱ��ܡ���������<font color=red>"&mmm&"</font>����<img src=horse/"&mmm&".gif>��������ǰ�棬�õ��˵�һ������ң�<font color=blue><b>"&yinma&"</b></font>�������Ѻ��,ÿ�˵õ���ע��<font color=red>3</font>��������,����Ѻ�������������������<font color=red>+500</font>��"
	call chat(says)
	set rs=nothing
	conn.close
	set conn=nothing
	Response.End 
end if
set rs=nothing
conn.close
set conn=nothing
says="<font color=blue><b>[������]</b></font>"&sjjh_name&"�ڽ�������ѡ��һƥ����<img src=horse/"&radiobutton&".gif>["&radiobutton&"]�ţ�����ƥ������ע:<font color=red>"&money&"</font>�������ڻ���<font color=blue>"&(10-dars)&"</font>������"
call chat(says)
Response.End 

sub chat(says)
says=replace(says,chr(39),"\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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