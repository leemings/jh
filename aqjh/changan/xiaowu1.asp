<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
my=aqjh_name
sex=aqjh_jhdj
%>
<!--#include file="data1.asp"--><%
sql="select * from 房屋 where 户主='" & my & "' or 伴侣='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('您还没有买房子呢！');location.href = 'fangwu.asp';}</script>"
elseif rs("状态")<>"正常" then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('您的房屋是不是有点状况呀！');location.href = '../welcome.asp';}</script>"
else
lx=rs("类型")
%>
<html>
<head>
<title>我 的 小 屋</title>
<style>
p{font-size:9pt; color:#ffee00}
td,select,input{font-size:9pt; color:#000000; height:14pt}
textarea{font-size:9pt; color:#000000}
A:link {COLOR: #ffffff; FONT-SIZE: 9pt;FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:visited {COLOR: #ffffff;FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:active {FONT-SIZE: 9pt; FONT-STYLE: normal; FONT-WEIGHT: normal; TEXT-DECORATION: none}
A:hover {COLOR: #ffff00; FONT-SIZE: 9pt; TEXT-DECORATION: underline}
</style>
</head>
<body bgcolor=#990099>
<center>
<p align="center"><b><font style="font-size: 9pt">欢迎<%=my%>回到自己的小窝</font></b><br><br>
<TABLE width='95%' ALIGN=center CELLSPACING=2 BORDER=2 CELLPADDING=5 BGCOLOR='#90c088'><tr><td>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=100%>
<tr bgcolor="#C4DEFF">
<td align="center" width=10%><a href="xiaowu.asp"><font color="#000000"><%if lx="一般房屋" or lx="高级公寓" or lx="花园洋房" or lx="豪华别墅" then%>客厅</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu1.asp"><font color="#000000"><%if lx="一般房屋" or lx="高级公寓" or lx="花园洋房" or lx="豪华别墅" then%>卧室</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu4.asp"><font color="#000000"><%if lx="高级公寓" or lx="花园洋房" or lx="豪华别墅" then%>卫生间</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu5.asp"><font color="#000000"><%if lx="高级公寓" or lx="花园洋房" or lx="豪华别墅" then%>小金库</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu6.asp"><font color="#000000"><%if lx="花园洋房" or lx="豪华别墅" then%>花园</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu7.asp"><font color="#000000"><%if lx="豪华别墅" then%>游泳池</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu8.asp"><font color="#000000"><%if lx="豪华别墅" then%>健身房</font></a><%end if%></td>
<td align="center" width=10%><a href="xiaowu9.asp"><font color="#000000"><%if lx="花园洋房" or lx="豪华别墅" then%>书房</font></a><%end if%></td>
</tr></table></TABLE>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=678 height="281">
<tr bgcolor=#f7f7f7>
<td height="1" width="101" bgcolor="#FFFFFF" rowspan="11"><%if sex="女" then%><img src="image/logonan.jpg"><%end if%><%if sex="男" then%><img src="image/logonv.jpg"><%end if%></td>                 
<td height="1" width="401" bgcolor="#C4DEFF" rowspan="11"><%if lx="一般房屋" then%><img src="image/ws-01.jpg"><%end if%><%if lx="高级公寓" then%><img src="image/ws-02.jpg"><%end if%><%if lx="花园洋房" then%><img src="image/ws-03.jpg"><%end if%><%if lx="豪华别墅" then%><img src="image/ws-04.jpg"><%end if%> </td>                 
<td height="1" width="159" bgcolor="#C4DEFF">
<form method=POST action='pub1.asp'>
你想要休息：</td>                 
</tr>
<tr><td height="1" width="159" bgcolor="#C4DEFF">
<select name=date size=1>
<option value="0">零天
<option value="1">一天
<option value="2">两天
<option value="3">三天
</select>
<select name=time size=1> 
<option value="0">0小时
<option value="1">1小时
<option value="2">2小时
<option value="3">3小时
<option value="4">4小时
<option value="5">5小时
<option value="6">6小时
<option value="7">7小时
<option value="8">8小时
<option value="9">9小时
<option value="10">10小时
<option value="11">11小时
<option value="12">12小时
<option value="13">13小时
<option value="14">14小时
<option value="15">15小时
<option value="16">16小时
<option value="17">17小时
<option value="18">18小时
<option value="19">19小时
<option value="20">20小时
<option value="21">21小时
<option value="22">22小时
<option value="23">23小时
</select>
</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">
<INPUT TYPE="submit" VALUE="确定"><INPUT TYPE="reset" VALUE="重写">
</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">在卧室里休息每小时</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">加内力1000体力1000</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">不需要银两。</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">　</td>                 
</tr><tr>
<td height="1" width="159" bgcolor="#C4DEFF">
    <p align="center">&nbsp;</p>
  </td>                 
</tr></form></table>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=678 height="1"> 
<tr>
<td height="164" width="102" bgcolor="#C4DEFF">
<p align="center"><a href="/myhome/sheepboy/feedsheep.asp" target="_blank"><font color="#000000">照料孩子</font></a></p> 
</td>                 
<td height="164" width="403" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">当前位置：卧室</font></p>
</td>                 
<td height="164" width="157" bgcolor="#C4DEFF">
  <p align="center"><a href="../welcome.asp"><img src="../jhimg/return.gif" width="40" height="20" border="0"></a></p>
</td>                 
</tr>
</table>
<br><br><br><br><br><!--#include file="copy.inc"-->
</body></html>
<%end if
set rs=nothing
conn.close
set conn=nothing
%>