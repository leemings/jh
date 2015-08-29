<%
response.buffer=true
'session("tclock")=empty

onep=request.cookies("endtime")("oneplay")
response.write onep

 if  session("tclock")>7 or onep="no"    then 
response.redirect "endclock.asp"
end if
if session("waityou")="" then  session("waityou")=1
session("waityou")=session("waityou")+1


 if application("reno")="" or application("reno")=0 then application("reno")=1
'	application("com1")="1"
'	application("com2")="2" 
'	application("com3")="3" 
'	application("com4")="4"   

if  application("com1")=""  or application("com2")=""  then 
'if  application("com1")=""  or application("com2")=""   then 

response.write "<body background=gif/msgbg.gif>"
response.write "<center><font color=green>等待玩家进入游戏</font></center>"
response.write "<center><a href=nan.asp><font color=red>等待玩家</font></a></center>"
response.write "<p>"
response.write "<font size=2 > (如果正在玩时出现此页面,就是中途有人退出游戏请从首页重新进入游戏)</font>"

response.write "</body >"
response.end
end if
reno=Trim(cstr(session("players")))
%>

<html>
<head>

<title>E线江湖--游戏在线</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="forum.css">
<style type="text/css">
BODY {
scrollbar-face-color:#efefef; 
scrollbar-shadow-color:#000000; 
scrollbar-highlight-color:#000000;
scrollbar-3dlight-color:#efefef;
scrollbar-darkshadow-color:#efefef;
scrollbar-track-color:#efefef;
scrollbar-arrow-color:#000000;
}
</style>
</head>


<body background="gif/msgbg.gif">
  
<table width="100%" border="1" bordercolorlight="#8E8EFF">
  <tr>
      <td> 
<%  for i=10 to 1 step-1
response.write application("msgg"&i)&"<br>"

next %>
&nbsp;
</td>
    </tr>
  </table>
  
 
  <div align="center">
<% if  application("play"&reno) ="4" or application("play"&reno) ="9"  or application("play"&reno) ="14"  or application("play"&reno) ="20"  or application("play"&reno) =25  or application("play"&reno) =27  or application("play"&reno) ="28"  or application("play"&reno) ="34"  or application("play"&reno) ="38" then %>

<form method="post" action="doing.asp">
<select name="charp"
    style="BACKGROUND-COLOR: #B3D9FF; COLOR: #000000; FONT-SIZE: 12px">

	<option value="siyi" selected>使用卡片</option>

	<% if session("stopp")<>"" then %>
	<option value="stopp" >停留卡</option>
	<% end if%>
	
	<%if session("nomoney")<>"" then	%>
	<option value="nomoney">免税卡</option>
	<% end if%>

	<%if session("addmoney")<>"" then 	%>
	<option value="addmoney">查税卡</option>
	<%end if%>

	<%if session("myhouse")<>""  then 	%>
	<option value="myhouse">购地卡</option>
	<%end if%>

	<%if session("pmoney")<>"" then  %>
	<option value="pmoney">均富卡</option>
	<%end if %>
</select>
对
<select name="ddwho"
    style="BACKGROUND-COLOR: #B3D9FF; COLOR: #000000; FONT-SIZE: 12px">
<option value="">无</option>
<option value="1">小丹尼</option>
<option value="2">宫本宝藏</option>
<option value="3">钱夫人</option>
<option value="4">沙隆巴斯</option>
</select>

<br>
        <input  class=buttonface  type=submit name=Submit2 value=确定>
        <input  class=buttonface  type=reset name=Submit value=取消>

</form>
<%
'108
end if%>

<table width="100%" border="0" cellspacing="0">
  <tr>
    <td><b><font color="#339933">现金:<%=formatcurrency(application("money"&reno))%> &nbsp;位置:<%=application("play"&reno)%></font></b></td>
  </tr>
</table>
 </div>

    <%
         '22
          if application("reno")>2 then   application("reno")=1
          if application("reno")=session("players") then 
    %> 
 
<table width="100%" border="0" cellspacing="0">
    <tr>
      
    <td>
      <div align="center"><a href="doing.asp"><img src="gif/nan.gif" width="54" height="48" alt="扔塞子" border="0"></a></div>
    </td>
    </tr>
  </table>
   <% else%> 
<table width="100%" border="0" cellspacing="0" align="center">
  <tr> 
    <td>  
      <div align="center"><a href=nan.asp><img src="gif/nan<%=session("weiz")%>.gif" width="54" height="48"   alt="等待塞子"  border="0"></a></div>
    </td>
  </tr>
</table>

 <% end if %> 
<form method="post" action="doing.asp">
<%
'146
if session("waityou")>10 then %>
  <table width="100%" border="0" cellspacing="0">
    <tr>
      <td>
        <div align="center"><font color="#B16565">你是否要重新开始游戏</font><font color="#287979">?</font> 
          (注意抢位置)</div>
      </td>
    </tr>
    <tr> 
      <td> 
        <div align="center"><img src="gif/bye.gif" width="103" height="54" border="0" usemap="#Map" href="replay.asp"><map name="Map"><area shape="rect" coords="6,4,44,48" href="replay.asp"><area shape="rect" coords="57,5,98,48" href="play.asp"><area shape="rect" coords="95,46,99,49" href="#"></map></div>
      </td>
    </tr>
  </table>
<% end if%>
  <table width="100%" border="0" cellspacing="0" height="45">
    <tr> 
      <td></td>
    </tr>
    <tr> 
      <td> 
        <input type="text" name="saymsg">
        <input  class=buttonface  type=submit name=Submit22 value=闲聊>
        <input  class=buttonface  type=reset name=Submit3 value=取消>
      </td>
    </tr>
  </table>
</form>
<table width="100%" border="0" cellspacing="0" height="36">
  <tr> 
    <td ><font color="#339966"><b><font color="#339933">房产清单:</font></b></font></td>
    <td><a href="hermsg.asp"><b><font color="#339933">对手资料</font></b></a></td>
  </tr>
  <tr> 
    <td width="50%"> <%	

'183
for i = 1 to 15
	whohouse=Trim(application("house"&i))
	whohouse=mid(whohouse,1,1)
	if   whohouse=reno  and reno<>"" then 
	houseTop=mid(trim(application("house"&i)),2,1)
	 response.write"房子"&i&  "~" &housetop& "级 "
response.write "<br >"

	end if
next
'200
%> 
&nbsp;
</td>
    <td width="50%"> <%	for i = 16 to 39
	whohouse=Trim(application("house"&i))
	whohouse=mid(whohouse,1,1)
	if   whohouse=reno  and reno<>"" then 
	houseTop=mid(trim(application("house"&i)),2,1)
	 response.write"房子"&i&  "~" &housetop& "级 "
response.write "<br >"

	end if
next
%>
&nbsp;
 </td>
  </tr>
</table>
</body>
</html>

