<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
nickname=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
chatroomsn=session("nowinroom")
lasttime=session("aqjh_lasttime")
if application("aqjh_mp")="" then 
	Response.Write "<script Language=Javascript>alert('�Բ���û�����������ѽ');window.close();</script>"
response.end
end if
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
yn=LCase(trim(request.querystring("yn")))
id=LCase(trim(request.querystring("id")))
dui=split(application("aqjh_mp"),"|")
to1=dui(0)
xiazhu=dui(2)
from=dui(3)
pai1=dui(4)
pai2=dui(5)
xiazhu=abs(int(xiazhu2))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
	Response.End
end if
if nickname<>to1 then
      	Response.Write "<script Language=Javascript>alert('�����˼�û���������ѽ');window.close();</script>"
      abc=0
response.end
else
if yn=0 then
Response.Write "<script Language=Javascript>alert('����Ϊ��������ʤ������,һͳ������ҵָ�տɴ�,��Ӧ��ʤ׷��,�ܾ��Է������.');window.close();</script>"
abc=1
duidu="<b><font color=green>["&nickname&"]</font></b>��<b><font color=black>["&from&"]</font></b>˵����������ʤ������,һͳ������ҵָ�տɴ�,��Ӧ��ʤ׷��."
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr 
sql="select q,h FROM p WHERE a='"&pai1&"'"
Set Rs=conn.Execute(sql)
kujin1=rs("h")
enemy1=rs("q")
if kujin1<xiazhu then
Response.Write "<script Language=Javascript>alert('���ź����������û����ô����⳥���');window.close();</script>"
duidu="<b><font color=green>["&nickname&"]</font></b>��<b><font color=black>["&from&"]</font></b>˵����������û����ô��Ŀ��,����ʲô�⳥������"
abc=0
conn.close
set conn=nothing
set rs=nothing
response.end
else
sql="select h,q FROM p WHERE a='"&pai2&"'and b='"&nickname&"'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
Response.Write "<script Language=Javascript>alert('�㲻�Ǹ����ɵ�����,ƾʲô�������ɽ����⳥');window.close();</script>"
duidu="["&nickname&"]��������������,��Ȩ���ܶԷ����ɵ��⳥��"
abc=1
conn.close
set conn=nothing
set rs=nothing
response.end
else
enemy2=rs("q")
 if InStr("|" & enemy2 & "|", "|" & pai1& "|")=0 then
Response.Write "<script language=JavaScript>{alert('�Է������Ѿ����������,����Ҫ�����');}</script>"
      set rs=nothing
      conn.close	
	set conn=nothing
       Response.End
       end if
 if InStr("|" & enemy1 & "|", "|" & pai2& "|")=0 then
Response.Write "<script language=JavaScript>{alert('�Է������Ѿ����������,����Ҫ�����');}</script>"
    set rs=nothing
    conn.close	
   set conn=nothing
       Response.End
       end if

Response.Write "<script Language=Javascript>alert('�㱧������̫ƽ�ľ���,ͬ�����������');window.close();</script>"
abc=1
duidu="<b><font color=green>["&nickname&"]</font></b>���岿�¼���,ͬʱ�ֿ���"&pai1&"���ɵ����Ǻ��г���,���������ַ���Լ."
enemy1=replace(enemy1,pai2&"|","")
enemy2=replace(enemy2,pai1&"|","")
   conn.execute"update p set q='"&enemy1&"',h=h+'"&xiazhu&"' where a='" & pai1 & "'"
    conn.execute"update p set q='"&enemy2&"',h=h-'"&xiazhu&"' where a='" & pai2 & "'"
    conn.close
set conn=nothing
set rs=nothing
end if
end if
end if
end if
session("aqjh_lasttime")=sj2
application("aqjh_mp")=""
if abc=1 then
says="<font color=#ff0000>��Ϣ</font>"& aqjh_name &"��"&duidu			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
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
end if
%>
