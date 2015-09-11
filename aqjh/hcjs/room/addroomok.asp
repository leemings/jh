<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="roomconfig.asp"-->
<%Response.Expires=0
Response.Buffer=true
aqjh_id=Session("aqjh_id")
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
nowinroom=session("nowinroom")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
if sfkf=0 then
	response.write "门派房间创建功能已关闭。"
	response.end
end if
if session("aqjh_inthechat")<>"1" then
	Response.Write "<script language=JavaScript>{alert('你必须在进入聊天室后才可以创建本门聚义厅！！');window.close();}</script>"
	Response.End
end if
if nowinroom<>0 then
	Response.Write "<script language=JavaScript>{alert('你必须在大厅时才可以创建本门聚义厅！！');window.close();}</script>"
	Response.End
end if
function jc(z)
	if not isnumeric(z) or (z<>0 and z<>1) then
		jc=true
	else
		jc=false
	end if
end function
h=clng(request.form("bh"))	'保护h
f=clng(request.form("fzxz"))	'发招f
g=clng(request.form("sjxz"))	'事件g
i=clng(request.form("kp"))	'卡片i
j=clng(request.form("db"))	'赌博j
c=clng(request.form("xz"))	'进入c
if jc(bh) or jc(fzxz) or jc(sjxz) or jc(kp) or jc(db) or jc(xz) then
	Response.Write "<script language=JavaScript>{alert('严重警告，提交数据错误！！');window.close();}</script>"
	Response.End 
end if

'固定值
b=200	'最高可以进入多少人

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select 门派,身份,银两,金币,会员金卡,grade from 用户 where 姓名='"&aqjh_name&"'",conn,2,2
mymp=rs("门派")
mysf=rs("身份")
myjinbi=rs("金币")
myyinliang=rs("银两")
myjinka=rs("会员金卡")
rs.close
if mysf<>"掌门" or aqjh_grade<5 or mymp="天网" or mymp="出家" or mymp="游侠" or mymp="官府" then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "<font size=2>你不是掌门或者你是天网、出家、游侠、官府的！</font>"
	response.end
end if
e="门派=''"&mymp&"''"
if xz=1 then
	d="只有"&mymp&"门派弟子方可进入"
	k=mymp&"重地，非本派人员擅闯者，杀无赫！！！"
else
	d="任何人都可以进入"
	k=mymp&"聚义厅，广交天下朋友，欢迎各们光临。"
end if
if myjinbi<jinbi or myyinliang<yinliang or myjinka<jinka then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "创建门派聚义厅需要花费 银两："&yinliang&"，金币："&jinbi&"，会员金卡："&jinka&"，你的钱不够，不能创建。"
	response.end
end if
rs.open "select count(姓名) as 数量 from 用户 where 门派='"&mymp&"' and 等级>="&dzdj,conn
zs=rs("数量")
rs.close
if zs<dzrs then
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "你门派弟子数等级在<font color=red>"&dzdj&"</font>级以上的只有<font color=red><b>"&zs&"</b></font>人，不足<font color=red><b>"&dzrs&"</b></font>，不可以创建房间。"
	response.end
end if
rs.open "select * from r where a='"&mymp&"'",conn,2,2
if not(rs.eof or rs.bof) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：一个门派只可以创建一个房间，你们派已经创建过房间了！');window.close();}</script>"
	Response.End
end if
rs.close
conn.execute "update 用户 set 银两=银两-"&yinliang&",金币=金币-"&jinbi&",会员金卡=会员金卡-"&jinka&" where 姓名='"&aqjh_name&"'"
conn.execute "insert into r(a,b,c,d,e,f,g,h,i,j,k) values ('"&mymp&"',"&b&","&c&",'"&d&"','"&e&"',"&f&","&g&","&h&","&i&","&j&",'"&k&"')"
rs.open "select * from r order by id",conn,2,2
application.lock()
roomming=""
do while Not rs.Eof
	rooming=rooming&rs("a")&"|"&rs("b")&"|"&rs("c")&"|"&rs("d")&"|"&rs("e")&"|"&rs("f")&"|"&rs("g")&"|"&rs("h")&"|"&rs("i")&"|"&rs("j")&"|"&rs("l")&"|"&rs("m")&";"
	if rs("a")=mymp then
		id=rs("id")
	end if
	rs.MoveNext
loop
application("aqjh_room")=rooming
aqjh_roominfo=split(Application("aqjh_room"),";")
dim nameindex(0)
for roomsn=0 to ubound(aqjh_roominfo)-1
	fenroom=split(aqjh_roominfo(roomsn),"|")
	if fenroom(0)=mymp then
		application("aqjh_chatroomname"&roomsn)=fenroom(0)
		Application("aqjh_onlinelist"&roomsn)=nameindex
		Application("aqjh_useronlinename"&roomsn)=""
	end if
next
application.unlock()
gg="恭喜<font color=blue>"&mymp&"</font>掌门<font color=red><b>"& aqjh_name &"</b></font><font color=blue><b>和他的弟子努力发展本门，"&dzdj&"级以上的弟子数达到"&dzrs&"以上，</b></font>站长特奖励其创建了本门的门派聚义厅：<font color=red><b>"&mymp&"</b></font>，希望"&aqjh_name&"以后继续努力，创建一个更加辉煌的业绩。<br>重进聊天室后即可进入"&mymp&"聚义大厅。"
ltssay="<table bgcolor=FFFF00 width=80% align=center border=1 cellspacing=0 cellpadding='0'><tr><td bgcolor=00FF00 align=center>门派聚义大厅创建公告</td><tr><td><div align='center'><font size=-1>"&gg&"</font></div></td></tr></table>"
says=ltssay		'聊天数据
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
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="../../hc3w_Admin/setup.css">
<title>门派聚义大厅创建</title></head>
<body text="#FFFFFF" background="../../jhimg/bk_hc3w.gif" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<div align="center">
  <font color="#000000" size="2">
  <p>江湖各门派聚义厅创建</p>
  </font></div>
<table border="1" cellspacing="1" align="center" cellpadding="0" bordercolor="#000000"
bgcolor="#006699" width="75%">
  <tr> 
    <td width="67"> 
      <div align="center"><font color="#FFFFFF" size=2>房间ＩＤ</font></div>
    </td>
    <td width="197"><font size="2"> <%=id%></font></td>
    <td width="81"> 
      <div align="center"><font size="2" color="#FFFFFF">房 间 名</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2"><%=mymp%></font></td>
  </tr>
  <tr> 
    <td width="67"> 
      <div align="center"><font color="#FFFFFF" size="2">发招限制</font></div>
    </td>
    <td width="197">
      <%if f=0 then%>
      允许
      <%else%>
      禁止
      <%end if%>
    </td>
    <td width="81" nowrap> 
      <div align="center"><font color="#FFFFFF" size="2">事件限制</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2">
      <%if g=0 then%>
      允许
      <%else%>
      禁止
      <%end if%>
      </font></td>
  </tr>
  <tr> 
    <td width="67" nowrap> 
      <div align="center"><font color="#FFFFFF" size="2">保&nbsp;&nbsp;护</font></div>
    </td>
    <td width="197"><font color="#FFFFFF" size="2">
      <%if h=0 then%>
      允许
      <%else%>
      禁止
      <%end if%>
      </font></td>
    <td width="81"> 
      <div align="center"><font color="#FFFFFF" size="2">卡&nbsp;&nbsp;片</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2">
      <%if h=0 then%>
      允许
      <%else%>
      禁止
      <%end if%>
      </font></td>
  </tr>
  <tr> 
    <td width="67"> 
      <div align="center"><font size="2" color="#FFFFFF">赌&nbsp;&nbsp;博</font></div>
    </td>
    <td width="197"><font color="#FFFFFF" size="2">
      <%if j=0 then%>
      允许
      <%else%>
      禁止
      <%end if%>
      </font></td>
    <td width="81"> 
      <div align="center"><font color="#FFFFFF" size="2">限&nbsp;&nbsp;制</font></div>
    </td>
    <td width="224"><font color="#FFFFFF" size="2"><%=d%></font></td>
  </tr>
  <tr> 
    <td colspan="4">
      <div align="center"> <font size="2"> <br>
        恭喜你成功创建了自已的门派聚义厅：<font color=red><b><%=mymp%></b></font> ，你现在只要重进入一次聊天室就可以看到你派的房间</font><br>
        <br>
        <input  onClick="javascript:window.close()" value="关 闭" type=button name="close">
      </div>
    </td>
  </tr>
</table>