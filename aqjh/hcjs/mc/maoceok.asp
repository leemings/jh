<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
%>
<body leftmargin="0" topmargin="0" bgcolor="#CCCCCC" background="../../bg.gif">
<%my=aqjh_name
money=abs(request.form("money"))
if  money<>10000000  then 
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
	Response.Redirect "../../error.asp?id=210"
	response.end
end if
sj=DateDiff("n",rs("����ʱ��"),now())
if sj<1 then
	s=1-sj
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
	Response.Write "<script Language=Javascript>alert('��û����ô��Ǯѽ��������ʲô����');window.close();</script>"
	response.end
end if
if rs("���")=0 then
  rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��Ǯ����?���Ϊ��,����ʲô��ѽ��');window.close();</script>"
	response.end
end if
conn.execute "update �û� set ���=0,����=����-"& money &",����ʱ��=now() where ����='"&aqjh_name&"'"
says="<font color=red>�������⡿["&aqjh_name&"]</font><font color=green>�򽭺��շ�վ����"&money&"����,����������!</font>"
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & aqjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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
Response.Write "<script Language=Javascript>alert('���򽭺��շ�վ����"& money &"�������������ϣ�');window.close();</script>"
rs.close
conn.close
set rs=nothing
set conn=nothing
%>