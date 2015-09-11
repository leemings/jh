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


Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
'以上是对联盟网吧判断
wbname=""
wbpf=0
ip=Request.ServerVariables("REMOTE_ADDR")
rs.open "SELECT barname FROM bar WHERE ip='"&ip&"'",conn
if Not(rs.Eof and rs.Bof) then
wbname=rs("barname")
wbpf=1
end if
rs.close
rs.open "Select 会员等级 from 用户 where 姓名='"&aqjh_name&"'",conn
hy=rs("会员等级")
if hy<=0 then
    Response.Write "<script Language=Javascript>alert('提示：你的条件不符合，必须是江湖等级会员，拉来的人泡点江湖等级40级以上并且拉人数超过20人才有资格进这里领取！');</script>"
response.end
end if
tmprs=conn.execute("Select count(*) As 数量 from 用户 where 等级>=40 and 介绍人='"& aqjh_name &"'")
lr=tmprs("数量")
set tmprs=nothing
if jsren<>20  then
Response.Write "<script language=JavaScript>{alert('拉来的人泡点必须是江湖等级40级以上并且拉人数超过20人才有资格每天来领取奖励!');location.href = 'javascript:history.back()';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
tmprs=conn.execute("Select count(*) As 数量 from 用户 where allvalue>100 and 介绍人='"&aqjh_name&"'")
jsren=tmprs("数量")
rs.open "Select * from 用户 where 姓名='"&aqjh_name&"'",conn
mdate=rs("领钱")
'if CDate(mdate)=now() then
'if day(mdate)>=day(now()) and month(mdate)=month(now()) and year(mdate)=year(now()) then 
if  DateDiff("d",rs("领钱"),now())=0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('天呀，["&aqjh_name &"]今天你来领过钱的！忘记了？');window.close();</script>"
	response.end
end if
yin=rs("银两")+rs("存款")
if yin>(rs("会员等级")+1)*500000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('天呀，["&aqjh_name &"]你现在有这么多的银两你还想要呀！！');window.close();</script>"
	response.end
end if
dj=rs("等级")
hy=rs("会员等级")
if trim(rs("身份"))="掌门" then
	'如是掌门统计人数
	tmprs=conn.execute("Select count(*) As 数量 from 用户 where 门派='"&rs("门派")&"'")
	renshu=tmprs("数量")
	'算钱数
	money=(hy+1)*200000+cint(renshu)*10000+jsren*500000+dj*250000+wbpf*10000000+20000
	conn.execute("update 用户 set 银两=银两+"&money&",领钱=now() where 姓名='"&aqjh_name&"'")
else
	'个人按会员乘10万
	money=(rs("会员等级")+1)*200000+jsren*500000+dj*250000+wbpf*10000000+20000
	conn.execute("update 用户 set 银两=银两+"&money&",领钱=now() where 姓名='"&aqjh_name&"'")
end if
if hy>=1 then
	jin=(rs("会员等级"))*10
	conn.execute("update 用户 set 金币=金币+"&jin&" where 姓名='"&aqjh_name&"'")
end if
says="<marquee width=100% scrollamount=8><font color=red><b>【拉人奖励】</b></font><font color=blue>"&aqjh_name&"发财了，拉人超过二十人，风的季节给你每天"&money&"块，金币"&jin&"块奖励哦，哈哈，因为你拉人多，钱就多多，记得天天上来拿呀，拉人就是好哦！</font></marquee>"			'聊天数据
says=replace(says,chr(39),"\'")
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

%>
<html>
<head>
<title>拉人奖励</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>
TD {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
BODY {
	FONT-FAMILY: "宋体";
	FONT-SIZE: 9pt;
	background-color: #004080;
}
SELECT {FONT-FAMILY: "宋体"; FONT-SIZE: 9pt}
A {COLOR: #FFC106; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: none}
A:hover {COLOR: #cc0033; FONT-FAMILY: "宋体"; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
body,td,th {
	color: #FFFFFF;
}
.style1 {color: #FFFFFF}
.style2 {color: #FF3300}
</STYLE>
</head>
<body bgcolor="#800000" style="background-color: #800000">
<div align="center"> 
  <p><b><span class="style2"><%=Application("aqjh_chatroomname")%><font size="+3">拉人奖励</font></span></b> 
    <br>
    <font size=+1> <b><%=aqjh_name%></b></font>你拉人：<font color="#FF0000"><b><%=jsren%>个</b></font>&nbsp; 
 <br>
    <span class="style1">
    <%if wbpf<>0 then%>
    您在联盟网吧：[<%=wbname%>]上网，多领20万块！</span><br>  
    <%end if%>
  今天你从江湖领取到了拉人奖励银子<b><font color="#FF0000"><%=money%>两</font></b>，金币<b><font color="#FF0000"><%=jin%>块</font></b>保存好哦，多多努力，多多拉人！   
    <%
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
    <br>
  </p>
<p align="center"><br>
<font color="#FF0000" size="+1"><b>拉人奖励计算方法</b></font><br>
<span class="style1">掌门：会员x2万+弟子数x1万+介绍人数x10000+战斗等级x2000+2万<br>
弟子：会员等级x20000+介绍人数x10000+江湖战斗等级x2000+2万</span></p>
  <span class="style1">金币：会员等级x10 
  </span>
  <p align="center"><span class="style1"><%=Application("aqjh_chatroomname")%></span></p>
<p align="center">
<input type=button value=关闭窗口 onClick='window.close()' name="button">
</p>
</div>
</body>
</html>

