<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
nickname=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
chatroomsn=session("nowinroom")
lasttime=session("sjjh_lasttime")
if application("sjjh_mp")="" then 
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
dui=split(application("sjjh_mp"),"|")
to1=dui(0)
xiazhu=dui(2)
from=dui(3)
men1=dui(4)
men2=dui(5)
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
Response.Write "<script Language=Javascript>alert('����Ϊ��������ʤ��������һͳ������ҵָ�տɴ�����Ӧ��ʤ׷�����ܾ��Է�����͡�');window.close();</script>"
abc=1
duidu="<b><font color=green>["&nickname&"]</font></b>��<b><font color=black>["&from&"]</font></b>˵����������ʤ��������һͳ������ҵָ�տɴ�����Ӧ��ʤ׷����"
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("sjjh_usermdb")
conn.open connstr 
sql="select q,h FROM p WHERE a='"&men1&"'"
Set Rs=conn.Execute(sql)
kujin1=rs("h")
shidi1=rs("q")
if kujin1<xiazhu then
Response.Write "<script Language=Javascript>alert('���ź����������û����ô����⳥���');window.close();</script>"
duidu="<b><font color=green>["&nickname&"]</font></b>��<b><font color=black>["&from&"]</font></b>˵����������û����ô��Ŀ������ʲô�⳥�����ɣ�"
abc=1
conn.close
set conn=nothing
set rs=nothing
response.end
else
sql="select h,q FROM p WHERE a='"&men2&"'and b='"&nickname&"'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
Response.Write "<script Language=Javascript>alert('�㲻�Ǹ����ɵ����ţ�ƾʲô�������ɽ����⳥');window.close();</script>"
duidu="["&nickname&"]�����Ͳ��Ǹ��������ţ���Ȩ���ܶԷ����ɵ��⳥��"
abc=1
conn.close
set conn=nothing
set rs=nothing
response.end
else
kujin1=rs("h")
shidi2=rs("q")
 if InStr("|" & shidi2 & "|", "|" & men1& "|")=0 then
Response.Write "<script language=JavaScript>{alert('�Է������Ѿ���������ˣ�����Ҫ�����');}</script>"
      set rs=nothing
      conn.close	
	set conn=nothing
       Response.End
       end if
 if InStr("|" & shidi1 & "|", "|" & men2& "|")=0 then
Response.Write "<script language=JavaScript>{alert('�Է������Ѿ���������ˣ�����Ҫ�����');}</script>"
    set rs=nothing
    conn.close	
   set conn=nothing
       Response.End
       end if

Response.Write "<script Language=Javascript>alert('�㱧����������ĸ�����٣�ͬ�����������');window.close();</script>"
abc=1
duidu="["&nickname&"]</font></b>���岿�¼��࣬ͬʱ�ֿ���"&men1&"���ɵ����Ǻ��г��⣬���°���֮�ˣ������ַ���"
shidis1=replace(shidi1,men2&"|","")
shidis2=replace(shidi2,men1&"|","")

   conn.execute"update p set h=h+'"&xiazhu&"',q='"&shidis1&"' where a='" & men1 & "'"
   conn.execute"update p set h=h-'"&xiazhu&"',q='"&shidis2&"' where a='" & men2 & "'"
   conn.close
set conn=nothing
set rs=nothing
end if
end if
end if
end if
session("sjjh_lasttime")=sj2
application("sjjh_mp")=""
if abc=1 then

says="<font color=#ff0000>��Ϣ</font>"& sjjh_name &"��"&duidu			'��������
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="��Ϣ"
towhoway=0
towho="���"
addwordcolor="660099"
saycolor="008888"
addsays="��"
saystr="<script>parent.sh("& chr(39) & addwordcolor & chr(39) &","& chr(39) & saycolor & chr(39) &","& chr(39) & act & chr(39) &","& chr(39) & sjjh_name & chr(39) &","& chr(39) & addsays & chr(39) &","& chr(39) & towho & chr(39) &"," & chr(39) & says & chr(39) &"," & towhoway &  ","& session("nowinroom") & ");<"&"/script>"
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
