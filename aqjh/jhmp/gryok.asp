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
%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../bg.gif">
<%my=aqjh_name
money=abs(request.form("money"))
if money<>1000 and  money<>10000 and money<>100000 and money<>1000000 then 
	Response.Write "<script Language=Javascript>alert('������ʲô����');window.close();</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	coon.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
	response.end
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�Բ���ϵͳ���ƣ����["&s&"����]�ٲ�����');}</script>"
	Response.End
end if	
if rs("����")<money then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��û��Ǯѽ��������ʲô����');window.close();</script>"
	response.end
end if
conn.execute "update �û� set ����=����+"&int(money/500)&",����=����-"& money &",����ʱ��=now() where ����='"&aqjh_name&"'"
Response.Write "<script Language=Javascript>alert('����¶�Ժ������"& money &"������������"&int(money/500)&"�㣡');window.close();</script>"
rs.close
conn.close
set rs=nothing
set conn=nothing
says="<font color=red>���¶�Ժ��Ϣ��["&aqjh_name&"]</font><font color=green>��¶�Ժ������"&money&"��,�������ǣ�"&int(money/500)&"��,��Ҵ������Ǹ�л����!</font>"
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