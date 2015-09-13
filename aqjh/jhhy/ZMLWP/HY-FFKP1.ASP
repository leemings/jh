<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../../pass.asp"-->
<!--#include file="../../mywp.asp"-->
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
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
psw=trim(Request.form("psw"))
rz=trim(Request.Form("psw2"))
rzjm=trim(request.form("regjm"))
for each element in Request.Form
	if instr(element,"'")<>0 or instr(element,"|")<>0 or instr(element," ")<>0 or instr(Request.Form(element),"'")<>0 or instr(Request.Form(element)," ")<>0 or instr(Request.Form(element),"|")<>0 then 
		Response.Write "<script Language=Javascript>alert('江湖提示：输入数据有问题，请查看！');location.href = 'javascript:history.go(-1)';</script>"
		Response.End
	end if
next
if len(psw)<5 then Response.Redirect "error.asp?id=52"
psw=md5(psw)
n=Year(date())
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r
if Weekday(date())<>1 then
	Response.Write "<script Language=Javascript>alert('对不起，会员卡片发放时间不到！星期日发放一次！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT id,w7,mvalue,等级,会员等级 FROM 用户 WHERE 姓名='"&session("aqjh_name")&"' and 密码='"&psw&"'",conn
If Rs.Bof OR Rs.Eof Then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你现在还不是我们江湖的人吧！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("mvalue")<=20000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('月积分在20000以上的才可以领取物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if rs("等级")<=60 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('为防止作弊，只有等级在60级以上的才可以领取物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
if hy=rs("会员等级")and hy<1 then
		rs.close
		set rs=nothing
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你不是等级制会员，请回吧!');history.go(-1)';}</script>"
		response.end
end if
myid=rs("id")
wpdate=rs("w7")
hydj=rs("会员等级")
aqjh_name=session("aqjh_name")
rs.close
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
conn1.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("sdlw.asp")
rs1.open "select * from a where yhid="&myid,conn1
if not(rs1.eof or rs1.bof) then
	rs1.close
	set rs1=nothing
	conn1.close
	set conn1=nothing
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('你已经领过礼品了！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs1.close
lw=""
randomize timer
r=int(rnd*5)+1
select case r
	case 1 
		n=999*hydj
                m=300*hydj
                k=500*hydj
                l=1000*hydj
                p=7000*hydj
		temp=add(wpdate,"美满生活",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&m&",土=土+"&k&",法力=法力+"&l&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="美满生活"&n&"个,金"&m&",木"&k&",水"&k&",火"&m&",土"&k&",法力"&l&",轻功"&l&",武功"&p&""
	case 2 
		n=999*hydj
                m=300*hydj
                k=500*hydj
                l=1000*hydj
                p=80000*hydj
		temp=add(wpdate,"求爱玫瑰",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&m&",土=土+"&k&",法力=法力+"&l&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="求爱玫瑰"&n&"个,金"&m&",木"&k&",水"&k&",火"&m&",土"&k&",法力"&l&",轻功"&l&",武功"&p&""	
	case 3 
		n=999*hydj
                m=400*hydj
                k=600*hydj
                l=800*hydj
                p=6000*hydj
		temp=add(wpdate,"鹊桥相会",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&l&",土=土+"&k&",法力=法力+"&l&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="鹊桥相会"&n&"个,金"&m&",木"&k&",水"&k&",火"&l&",土"&k&",法力"&l&",轻功"&l&",武功"&p&""
       case 4 
		n=999*hydj
                m=700*hydj
                k=600*hydj
                l=500*hydj
                p=6000*hydj
		temp=add(wpdate,"玫瑰衷情",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&l&",土=土+"&k&",法力=法力+"&l&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="玫瑰衷情"&n&"个,金"&m&",木"&k&",水"&k&",火"&l&",土"&k&",法力"&l&",轻功"&m&",武功"&p&""
	case 5 
		n=999*hydj
                m=300*hydj
                k=500*hydj
                l=900*hydj
                p=7000*hydj
		temp=add(wpdate,"跳舞的玫瑰",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&k&",土=土+"&l&",法力=法力+"&k&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="跳舞的玫瑰"&n&"个,金"&m&",木"&k&",水"&k&",火"&k&",土"&k&",法力"&k&",轻功"&l&",武功"&p&""	
	case 6 
		n=999*hydj
                m=600*hydj
                k=500*hydj
                l=700*hydj
                p=5000*hydj
		temp=add(wpdate,"想飞的玫瑰",n)
		sql="update 用户 set 金=金+"&m&",木=木+"&m&",水=水+"&m&",火=火+"&l&",土=土+"&m&",法力=法力+"&k&",轻功=轻功+"&m&",武功=武功+"&p&",w7='"&temp&"' where id="&myid
		lw="想飞的玫瑰"&n&"个,金"&m&",木"&k&",水"&k&",火"&l&",土"&m&",法力"&k&",轻功"&l&",武功"&p&""	
	
end select
set rs=conn.execute(sql)
conn1.execute "insert into a(yhid,姓名,礼物) values ('"&myid&"','"&aqjh_name&"','"&lw&"')"
set rs=nothing
conn.close
set conn=nothing
says="【<font color=#ff0000>周末礼品</font>】<font color=blue>恭喜<font color=red><b>【"& aqjh_name &"】</b></font>从站长那里得到了【<font color=red><b>"& lw &"</b></font>】，大家快去<font color=red><b>【礼品屋】</b></font>领礼品啊！去晚了就没有了啊！<font color=red size=+1><b>祝各位玩家周末快乐!</b></font>"			'聊天数据
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
act="消息"
towhoway=0
towho="大家"
addwordcolor="660099"
saycolor="008888"
addsays="对"
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
set rs1=nothing
conn1.close
set conn1=nothing
%>
<html>
<head>
<title>周末礼品屋</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>

<body bgcolor="#FF9900" background="bg.gif">
<table width=401 height="227" border=3 align=center cellpadding="5" cellspacing="10" bordercolor="#6633CC">
  <tr bgcolor="#FFFFFF" align="center"> 
      <td width="367" height="252" bgcolor="#FF9900" background="bg.gif"><div align="center"><b>恭喜你从<font color="#FF0000"><%=application("aqjh_user")%></font>站长那里得到了<br>
        <font color="#66FF00"><%=lw%></font></b><br>
<br>
<br>
<br>
<br>
[快乐江湖]</div></td>
  </tr>
</table>
</body>
</html>
