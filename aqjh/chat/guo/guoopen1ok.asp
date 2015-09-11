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
if application("aqjh_guo")="" then 
	Response.Write "<script Language=Javascript>alert('对不起，没有人向你求和呀');window.close();</script>"
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
dui=split(application("aqjh_guo"),"|")
to1=dui(0)
xiazhu=dui(2)
from=dui(3)
guo1=dui(4)
guo2=dui(5)
xiazhu=abs(int(xiazhu2))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End
end if
if nickname<>to1 then
      	Response.Write "<script Language=Javascript>alert('错误：人家没有向你求和呀');window.close();</script>"
      abc=0
response.end
else
if yn=0 then
Response.Write "<script Language=Javascript>alert('你认为你们国家胜利在望,一统江湖大业指日可待,理应乘胜追击,拒绝对方的求和.');window.close();</script>"
abc=1
duidu="<b><font color=green>["&nickname&"]</font></b>对<b><font color=black>["&from&"]</font></b>说：我们国家胜利在望,一统江湖大业指日可待,理应乘胜追击."
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr 
sql="select d,jin FROM guo WHERE g='"&guo1&"'"
Set Rs=conn.Execute(sql)
kujin1=rs("jin")
enemy1=rs("d")
if kujin1<xiazhu then
Response.Write "<script Language=Javascript>alert('很遗憾，求和的国家没有那么多的赔偿库金');window.close();</script>"
duidu="<b><font color=green>["&nickname&"]</font></b>对<b><font color=black>["&from&"]</font></b>说：你们国家连这么点钱都没有,你拿什么赔偿给本国！！"
abc=0
conn.close
set conn=nothing
set rs=nothing
response.end
else
sql="select jin,d FROM guo WHERE g='"&guo2&"'and j='"&nickname&"'"
Set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
Response.Write "<script Language=Javascript>alert('你不是该国家的国王,凭什么代表国家接受赔偿');window.close();</script>"
duidu="["&nickname&"]根本就就不是国家的君王,还想滥竽充数接受对方国家的赔偿。真丢人！
abc=1
conn.close
set conn=nothing
set rs=nothing
response.end
else
enemy2=rs("d")
 if InStr("|" & enemy2 & "|", "|" & guo1& "|")=0 then
Response.Write "<script language=JavaScript>{alert('对方好像已经和你求和了,不需要求和了');}</script>"
      set rs=nothing
      conn.close	
	set conn=nothing
       Response.End
       end if
 if InStr("|" & enemy1 & "|", "|" & guo2& "|")=0 then
Response.Write "<script language=JavaScript>{alert('对方好像已经和你求和了,不需要求和了');}</script>"
    set rs=nothing
    conn.close	
   set conn=nothing
       Response.End
       end if

Response.Write "<script Language=Javascript>alert('你抱着天下太平的精神,同意了他的求和');window.close();</script>"
abc=1
duidu="<b><font color=green>["&nickname&"]</font></b>深体战争只会给百姓带来疾苦,同时又看到"&guo1&"国的将士们很有诚意,逐定下永不侵犯条约.该事件也从此载入史册，供后人效尤。。。。"
enemy1=replace(enemy1,guo2&"|","")
enemy2=replace(enemy2,guo1&"|","")
   conn.execute"update guo set d='"&enemy1&"',jin=jin+'"&xiazhu&"' where g='" & guo1 & "'"
    conn.execute"update guo set d='"&enemy2&"',jin=jin-'"&xiazhu&"' where g='" & guo2 & "'"
    conn.close
set conn=nothing
set rs=nothing
end if
end if
end if
end if
session("aqjh_lasttime")=sj2
application("aqjh_guo")=""
if abc=1 then
says="<font color=#ff0000>消息</font>"& aqjh_name &"："&duidu			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
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
