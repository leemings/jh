<!--#include file="conn.asp"-->
<!--#include file="z_fy_Conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_fy_const.asp"-->
<!-- #include file="inc/FORUM_CSS.asp" --> 

<%
'=========================================================
' Plusname: 法院监狱第三版
' File: z_fy_tanjian.asp
' For Version: 6.0sp2
' Date: 2003-2-10
' Script Written by Hero20000
'=========================================================
if not founduser then 
  Errmsg=Errmsg+"<br>"+"<li>游客不能探视囚犯。"
   call dvbbs_error()
else
  call getfyconfig()
  dim fanren,jailsj,sjpic,sjstr
  call checkjail()
  if not founderr then
    select case checkstr(request("action"))
      case "save" 
        call savetj()
      case "qiang"
        call qiang()       
      case else
        call main()
    end select
  end if
end if
rs.close
set rs=nothing

sub main()  
%>
<html>
<head>
<title><%=membername%>探望囚犯<%=fanren%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body background="Images/img_fy/bg001.gif" >
<form action=z_fy_tanjian.asp?action=save&name=<%=fanren%> method=post>
<table border=0  cellspacing=1 cellpadding=3 align="center" width="60%" ><center>
  <tr> 
    <td align=center > 
<%
randomize 
jailsj=Int((15*rnd)+1)
select case jailsj
 case 1
  sjpic="Images/img_fy/sj1.gif"
  sjstr="他正在狱中痛哭流涕，后悔自己的所作所为。"
 case 2
  sjpic="Images/img_fy/sj2.gif"
  sjstr="他开始参加培训班，学习社区法律法规。"
 case 3
 sjpic="Images/img_fy/sj3.gif"
  sjstr="他已经开始自学少林寺的武功精要，准备出狱后伺机报复。"
 case 4
 sjpic="Images/img_fy/sj4.gif"
  sjstr="他在狱中仰望长天，就这样过了一天。"
 case 5
 sjpic="Images/img_fy/sj5.gif"
  sjstr="他今天好像把监狱的铁门撞开了，看守所动用了坦克才将他制服。"
 case 6
 sjpic="Images/img_fy/sj6.gif"
  sjstr="罪犯企图越狱，被监狱长抓回，现正在隔离审查。"
 case 7
 sjpic="Images/img_fy/sj7.gif"
  sjstr="他好像和其他狱友相处的不是很好！"
 case 8
 sjpic="Images/img_fy/sj8.gif"
  sjstr="他在里面生活的很好，每天唱歌跳舞。"
 case 9
 sjpic="Images/img_fy/sj9.gif"
  sjstr="今天社区监狱的饭菜可是非常丰盛。"
 case 10
 sjpic="Images/img_fy/sj10.gif"
  sjstr="由于他在监狱期间表现良好，法官可能给他减刑。"
 case 11
 sjpic="Images/img_fy/sj11.gif"
  sjstr="他好像特别能吃，已经吃了我们半年的口粮，典狱长正考虑提请法官减少刑期，以便缩减开支。"
 case else
 sjpic="Images/img_fy/sj1.gif"
  sjstr="他正在狱中痛哭流涕，后悔自己的所作所为。"
end select
response.write "<img src="&sjpic&" width=350 height=345 >"
%>
    </td>
  </tr>
  <tr><td align=center><%=sjstr%>
</td></tr>
  <tr><td align=center>留言：<input name=tjly size=30 maxlength=20 > 最多250字符</td></tr></center>   
 </table>
<div align="center">【<a href=z_fy_looktj.asp?name=<%=fanren%> title='点击查看留言'>看看<%=fanren%>的留言本</a>】<br><br>
<% if master or jymaster then %>
<input type="submit" value='管理员留言'  name="submit">&nbsp;&nbsp;<input type="button" value="懒得理他" onClick="window.close()" name="button">
<%else%>
<input type="submit" value='花<%=tj_set(1)%>元留言'  name="submit">&nbsp;&nbsp;<input type="button" value="懒得理他" onClick="window.close()" name="button">
<%end if%>
</form>
<%if qj_set(0)="1" then%>
<form action=z_fy_tanjian.asp?action=qiang&name=<%=fanren%> method=post>
&nbsp;<input type="submit" value="嘿嘿，你也有今天，抢你没商量！" name="submit2">
</form>
<%end if%>
</div>
</body>   
</html>   
<%
end sub

sub savetj()
  if isnull(request.form("tjly")) or request.form("tjly")="" then
   content="默默无语两眼泪，耳边响起驼铃声......"
  else
   content=checkstr(request.form("tjly"))
  end if
  mysign="2|0"
  call logs("监狱","探监",membername,fanren)
  if (not master) and (not jymaster) then
  sql="update [user] set userwealth=userwealth-"&tj_set(1)&" where username='"&membername&"'"
  conn.execute sql
  end if
  %>
  <html>
<head>
<title><%=membername%>探望囚犯<%=fanren%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body background="Images/img_fy/bg001.gif" >
<table border=0  cellspacing=1 cellpadding=3 align="center" valign="middle" width="60%" ><center>
  <tr> 
    <td align=center  ><p><br><br><%=membername%>探监并留言:<%=content%><br><br><%=fanren%>哽咽地说：~~要我说什么好呢？<br><br>出狱后我一定报答你！</td></tr></table>
<div align="center"><input type="button" value="关闭窗口" onClick="window.close()" name="button">
</div>
</body>   
</html> 
<%  
end sub

sub qiang()
dim qjjg,userw,usere,userc,qjtype,mess1,qjs,qds
if qj_set(4)="1" then
  if session("qjhave")<>"" then
	if session("qjhave")="haveqj" then
	  Errmsg=Errmsg+"<br>"+"<li>不要太贪心，20分钟内只能抢一次哦。"
         call dvbbs_error()
         exit sub
	end if
  end if
end if
session("qjhave")="haveqj"
  if qj_set(0)="0" then
   mess1="社区监狱现在警卫森严，您什么也没抢到！"
   qds="0"
  else
   sql="select userwealth,userep,usercp from [user] where username='"&fanren&"'"
   set rs=conn.execute(sql)
   userw=rs(0)
   usere=rs(1)
   userc=rs(2)
   randomize 
   qjjg=Int((30*rnd)+1)
   select case qjjg
    case 1
     mess1="小小虾米喽，嘻嘻！抢到现金 "
     qjtype="userwealth"
     qjs=0.01
     qds=clng(userw*qjs)
    case 8
     mess1="聊胜于无，哈哈，抢到经验值 "
     qjtype="userep"
     qjs=0.01
     qds=clng(usere*qjs)
    case 15
     mess1="没白费力气，抢到魅力值 "
     qjtype="usercp"
     qjs=0.01
     qds=clng(userc*qjs)
    case 4
     mess1="不错不错，抢到现金 "
     qjtype="userwealth"
     qjs=0.05
     qds=clng(userw*qjs)
    case 18
     mess1="解气呀，抢到经验值 "
     qjtype="userep"
     qjs=0.05
     qds=clng(usere*qjs)
    case 29
     mess1="都进了大狱了，要嘛魅力，抢到魅力值 "
     qjtype="usercp"
     qjs=0.05
     qds=clng(userc*qjs)
    case 5
     mess1="歹势了，呵呵，竟然抢到现金 "
     qjtype="userwealth"
     qjs=0.1
     qds=clng(userw*qjs)
    case 26
     mess1="耶！被我抢到经验值 "
     qjtype="userep"
     qjs=0.1
     qds=clng(usere*qjs)
    case 23
     mess1="我是热血强人，抢到魅力值 "
     qjtype="usercp"
     qjs=0.1
     qds=clng(userc*qjs)
    case 10
     mess1="福星高照，嘿嘿，抢到现金 "
     qjtype="userwealth"
     qjs=0.2
     qds=clng(userw*qjs)
    case 21
     mess1="运气真不错，抢到经验值 "
     qjtype="userep"
     qjs=0.2
     qds=clng(usere*qjs)
    case 7
     mess1="要晕倒了，我抢到了魅力值 "
     qjtype="usercp"
     qjs=0.2
     qds=clng(userc*qjs)
    case 11
     mess1="天哪，我竟然合法抢到现金 "
     qjtype="userwealth"
     qjs=0.3
     qds=clng(userw*qjs)
    case else
     mess1="嗯？这小子钱都放哪儿了！什么也没抢到!"
     qjtype="none"
     qjs=1
     qds=0
  end select
  if qjtype<>"none" then
     sql="update [user] set "&qjtype&"="&qjtype&"-round("&qds&") where username='"&fanren&"'"  
     conn.execute sql
     sql="update [user] set "&qjtype&"="&qjtype&"+round("&qds&") where username='"&membername&"'"
     conn.execute sql
     if qj_set(1)="1" then
        sql="update [user] set usercp=usercp-"&qj_set(2)&" where username='"&membername&"'"
        conn.execute sql
     end if
     select case qjtype
      case "userwealth"
           qjtype="现金"
      case "userep"
           qjtype="经验"
      case "usercp"
           qjtype="魅力"
     end select
     content="祸不单行！"&fanren&"被强盗抢走 "&qjtype&":"&qds
     mysign="3|0"
     call logs("监狱","抢劫",membername,fanren)
  end if
end if
  
  %>
  <script>
confirm('<%=mess1%><%=qds%>')
</script>

  <html>
<head>
<title><%=membername%>抢劫<%=fanren%>了！<%=fanren%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

</head>
<body background="Images/img_fy/bg001.gif" >
<table border=0  cellspacing=1 cellpadding=3 align="center" valign="middle" width="60%" ><center>
  <tr> 
    <td align=center  ><p><br><br><%=mess1%><%if qds<>"0" then response.write(qds)%><br><br><%if clng(userw)*clng(qjs)<50 or clng(usere)*clng(qjs)<10 or clng(userc)*clng(qjs)<10 then
            response.write " "&fanren&"骂道：还抢！还抢？我就剩个裤头了，呜呜~~<br><br>"
    else%>
<%=fanren%>哽咽地说：人倒霉真是什么事都能碰到<br><br><%end if%>出狱后我一定“报答”你！</td></tr></table>
<div align="center"><input type="button" value="关闭窗口" onClick="window.close()" name="button">
</div>
</body>   
</html> 
<%  
end sub

sub checkjail()
if tj_set(0)="0" then
    Errmsg=Errmsg+"<br>"+"<li>社区监狱现在关闭了探监功能。"
    founderr=true
    call dvbbs_error()
    exit sub
end if
fanren=checkstr(request("name"))
sql="select userclass from [user] where username='"&fanren&"' and lockuser=1"
set rs=conn.execute(sql)
if rs.eof then
  Errmsg=Errmsg+"<br>"+"<li>无此囚犯或该囚犯已出狱！"
  founderr=true
  call dvbbs_error()
  exit sub
end if
sql="select userwealth from [user] where username='"&membername&"'"
set rs=conn.execute(sql)
if rs(0)<clng(tj_set(1)) and not master and not jymaster then
  Errmsg=Errmsg+"<br>"+"<li>您的探视费不够！"
  founderr=true
  call dvbbs_error()
  exit sub
end if
end sub
%>
