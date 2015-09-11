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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sql="select * from 用户 where 姓名='"&aqjh_name&"'"
set rs=conn.execute(sql)
sex=rs("性别")
ml=rs("魅力")
dd=rs("道德")
my=aqjh_name
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
td,select,input{font-size:9pt; color:#000000; height:9pt}
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
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=700 height="281">
<tr bgcolor=#f7f7f7>
<td height="1" width="101" bgcolor="#FFFFFF" rowspan="10"><%if sex="女" then%><img src="image/logonan.jpg"><%end if%><%if sex="男" then%><img src="image/logonv.jpg"><%end if%></td>                 
<td height="1" width="401" bgcolor="#C4DEFF" rowspan="10"><%if lx="花园洋房" then%><img src="image/sf-01.jpg"><%end if%><%if lx="豪华别墅" then%><img src="image/sf-02.jpg"><%end if%> </td>                 
</tr>

</table>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=700 height="1"> 
<tr>
<td height="164" width="102" bgcolor="#C4DEFF">
 <p align="center"><a href="/xywl/html/xiashuo/indexxiao.htm"><font color="#000000">看书</font></a></p> 
</td>                 
<td height="164" width="403" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">当前位置：书房</font></p>
</td>                 
<td height="164" width="157" bgcolor="#C4DEFF">
  <p align="center"><a href="../welcome.asp"><img src="../jhimg/return.gif" width="40" height="20" border="0"></a></p>
</td>
<td height="164" width="157" bgcolor="#C4DEFF">
<font color="#ff0000">魅力：<%=ml%><br>道德：<%=dd%></font>
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

