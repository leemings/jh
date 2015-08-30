<%
response.buffer=true
'判断是否非法进入

pp=request("player")
'response.write pp  
'response.write "<br>"
	'response.write "公有变量"
	'response.write application("com1")

if application("com1")=pp and pp<>""  or application("com2")=pp and pp<>"" then 

response.redirect "loading.asp"

end if

'清空所有变量
session("tclock")=empty
session("stopp")=""
session("zjstopp")=empty
session("nomoney")=""
session("zjnomoney")=empty
session("addmoney")=""
session("zjaddmoney")=empty
session("pmoney")=""
session("zjpmoney")=empty
session("myhouse")=""
session("zjmyhouse")=empty


session("players")=" "
session("players")=clng(request("player")) '表示第几个玩家,值为1,2,3,4
'dd=clng(request("player")) '表示第几个玩家,值为1,2,3,4
'response.write session("players")
'response.end

j=trim(cstr(session("players")))
application("com"&j)=pp
	response.write application("com"&j)   '表示第几个玩家已经进入游戏
application("msgg")=""    '信息
application("reno")=1   '表示玩的顺序

for i =1 to 4 	'现在暂时只有二人 玩家
application("play"&i)=""
application("play"&i)=0
'application("money"&i)=50000    '初始钱为5万元人民币
'response.write application("money"&i)
'response.write "<br>"
next

for i =1 to 40    '到房子个数,应该为41
application("house"&i)=""
application("msgg"&i)=""
next
 if trim(cstr(session("players")))<>"0"  then
 response.redirect "charp.asp"
%>
  <SCRIPT LANGUAGE="VBScript">

   '   Window.Open "charp.asp", "ChaWindow", "MENUBAR=0, SCROLLBARS=0, HEIGHT=550,WIDTH=800 "
  </SCRIPT>
<%end if%>



<html>
<head>
<title>快乐江湖--游戏在线</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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

<body bgcolor="#FFFFFF" background="gif/msgbg.gif">
<table width="789" border="0" cellspacing="0">
  <tr> 
    <td width="19%">　</td>
    <td width="17%"><img src="gif/more1.gif" width="154" height="159"></td>
    <td width="19%"><img src="gif/more2.gif" width="145" height="159"></td>
    <td width="45%"><img src="gif/more3.gif" width="150" height="159"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
      <div align="center"><a href="loading.asp"><img src="gif/look.gif" width="39" height="39" border="0" alt="刷新,查看空闲位置"></a></div>

    </td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0">
  <tr>
    <td width="8%">　</td>
    <td width="44%"> 
      <table width="291" border="0" cellspacing="0" height="281" align="center">
        <tr> 
          <td height="61">　</td>
          <td height="61" valign="bottom"> <%if application("com2")="" then	%> 
            <div align="center"><a href="loading.asp?player=2"><img src="gif/big2.gif" width="52" height="45" border="0" alt="宫本宝藏"></a></div>
            <%else%> 
            <div align="center"><img src="gif/big2come.gif" width="52" height="45" alt="宫本一郎"></div>
            <%end if %> </td>
          <td height="61">　</td>
        </tr>
        <tr> 
          <td> <%if application("com1")="" then	%> 
            <div align="right"><a href="loading.asp?player=1"><img src="gif/big1.gif" width="52" height="45" border="0" alt="小丹尼"></a></div>
            <% else %> 
            <div align="right"><img src="gif/big1come.gif" width="52" height="45" alt="小丹尼"></div>
            <% end if%> </td>
          <td>
          <p align="center">1号游戏</td>
          <td> <%if application("com3")<>"" then	%> 
            <div align="left"><a href="loading.asp?player=3"><img src="gif/big3.gif" width="52" height="45" border="0" alt="钱夫人"></a></div>
            <%else%> 
            <div align="left"><img src="gif/big3come.gif" width="52" height="45" border="0"></div>
            <%end if%> </td>
        </tr>
        <tr> 
          <td>　</td>
          <td valign="top"> <%if application("com4")<>"" then	%> 
            <div align="center"><a href="loading.asp?player=4"><img src="gif/big4.gif" width="52" height="44" usemap="#Map4" border="0" alt="沙隆巴斯"></a></div>
            <%else%> 
            <div align="center"><img src="gif/big4come.gif" width="52" height="45" alt="酋长"></div>
            <%end if %> </td>
          <td>　</td>
        </tr>
      </table>
    </td>
    <td width="48%"> 
      <table width="291" border="0" cellspacing="0" height="281" align="center">
        <tr> 
          <td height="57">　</td>
          <td height="57" valign="bottom"> <%if application("com2")="" then	%> 
            <div align="center"><img src="gif/big2.gif" width="52" height="45" border="0" usemap="#Map2Map"><map name="Map2Map"><area shape="rect" coords="3,1,62,46" href="loading.asp?player=2" alt="宫本一郎" title="宫本一郎"></map></div>
            <%else%> 
            <div align="center"><img src="gif/big2come.gif" width="52" height="45" alt="宫本一郎"></div>
            <%end if %> </td>
          <td height="57">　</td>
        </tr>
        <tr> 
          <td> <%if application("com1")="" then	%> 
            <div align="right"><img src="gif/big1.gif" width="52" height="45" border="0" usemap="#MapMap"><map name="MapMap"><area shape="rect" coords="2,2,54,46" href="loading.asp?player=1" alt="小丹尼" title="小丹尼"></map></div>
            <% else %> 
            <div align="right"><img src="gif/big1come.gif" width="52" height="45" alt="小丹尼"></div>
            <% end if%> </td>
          <td>
          <p align="center">2号游戏</td>
          <td> <%if application("com3")<>"" then	%> 
            <div align="left"><img src="gif/big3.gif" width="52" height="45" usemap="#Map3Map" border="0"><map name="Map3Map"><area shape="rect" coords="2,1,53,44" href="loading.asp?player=3" alt="钱夫人" title="钱夫人"></map></div>
            <%else%> 
            <div align="left"><img src="gif/big3come.gif" width="52" height="45" usemap="#Map3Map" border="0"><map name="Map3Map"><area shape="rect" coords="2,1,53,44" href="loading.asp?player=3" alt="钱夫人" title="钱夫人"></map></div>
            <%end if%> </td>
        </tr>
        <tr> 
          <td>　</td>
          <td valign="top"> <%if application("com4")<>"" then	%> 
            <div align="center"><img src="gif/big4.gif" width="52" height="45" usemap="#Map4Map" border="0"><map name="Map4Map"><area shape="rect" coords="37,12,38,13" href="#"><area shape="rect" coords="1,2,53,46" href="loading.asp?player=4" alt="酋长" title="酋长"></map></div>
            <%else%> 
            <div align="center"><img src="gif/big4come.gif" width="52" height="45" alt="酋长"></div>
            <%end if %> </td>
          <td>　</td>
        </tr>
      </table>
    </td>
  </tr>
 
</table>
<table width="100%" border="0" cellspacing="0" height="8">
  <tr> 
    <td colspan="2"> 
      <div align="center"><div align="center"><a href="delall.asp">清除所有玩家</a></div></div>
    </td>
  </tr>
</table>
</body>
</html>