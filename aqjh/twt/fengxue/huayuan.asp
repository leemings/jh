<!--#include file="func213.asp"-->
<!--#include file="sub.asp" -->
<SCRIPT LANGUAGE=JavaScript>if(window.name!='aqjh_win'){var i=1;while(i<=50){window.alert('ˢǮ����ϲ�����𣿵㰡��ˢ������');i=i+1;}top.location.href='../../exit.asp'}</script>
<%
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sh=request.form("sh")
'sy=request.form("sy")
if request.form("h")="1" then
my=session("aqjh_name")
set conn=server.createobject("adodb.connection")
conn.open Application("aqjh_usermdb")
set rs=server.CreateObject ("adodb.recordset")
sql="select * from �û� where ����='" & aqjh_name & "'"
set rs=conn.execute(sql)
if rs("����")<0 or rs("״̬")="��" then
Response.Redirect "../../error.asp?id=421"
	rs.close
        set rs=nothing
        conn.close
        set conn=nothing
	response.end
end if
tl=rs("����")
if rs("����")<20 then	
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('��ƣ�ͳ̶��ѳ�����Χ��Ϊ�����⣬�����뿪�˵�Ϊ�ϣ�');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
sql="update �û� set ����=����-20 where ����='" & aqjh_name & "'"
	conn.execute sql
        rs.close
        set rs=nothing
	conn.close
        set conn=nothing
        message=huayuan(my,sh,sy)
end if
says="<font color=#ff0000>����ѩ�轣��</font>"&message
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
Response.Write "<script language=JavaScript>{alert('��ʾ��"&message&"');location.href='FENGXUE.ASP'}</script>"
	Response.End 
%>