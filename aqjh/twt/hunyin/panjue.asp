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
if aqjh_grade<>10  then
	Response.Write "<script Language=Javascript>alert('��ʾ������������ʲô������ѽ��');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
name=request("name")
my=request("my")
rs.open "select * from h where a='" & name & "'and b='" & my &"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������Ҫ�о�������["&my&"]�����ڣ���');history.go(-1);</script>"
	response.end
end if
fgyin=rs("d")
rs.close
rs.open "select * from �û� where ����='" & name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ʾ������Ҫ�о�������["&name&"]�����ڣ���');history.go(-1);</script>"
	response.end
else
	myboy=rs("boy")
	if isnull(myboy) or myboy="" then
		myboy=""
	else
		zt=split(myboy,"|")
		if UBound(zt)<>7 then
			myboy=""
		end if
	end if
end if
if myboy<>"" then
 conn.execute "insert into gry(boy,fm1,fm2) values ('"&myboy&"','"&name&"','"&my&"')"
end if
conn.execute "update �û� set ��ż='��',boy='',boysex='',����=����+"&fgyin&" where ����='" & my & "'"
conn.execute "update �û� set ��ż='��',boy='',boysex='' where ����='" & name & "'"
conn.execute "delete * from h where a='" & name & "' or b='" & name & "' "
conn.execute "delete * from t where b='" & name & "' or c='" & name & "'"
message="[���]" & my & "��" & name & "���ٸ��о����ɹ���Ϊ�����ػ����ɹ��ƣ�"
if myboy<>"" then message=message&"˫����С���Ѿ��͵��¶�Ժ���������޹��İ�������"
rs.close
set rs=nothing
conn.close
set conn=nothing
says="<font color=red>����Ϣ��</font><b><font color=blank>'"& message &"'</font></b>"		'��������
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
<body bgcolor="#000000" background="../../bg.gif">
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=message%></p>
<p align=center><a href=Yuanou.asp>����</a></p>
</td></tr>
</table>
</td></tr>
</table></body>