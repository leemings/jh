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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select * from 用户 where 姓名='"&aqjh_name&"'",conn
sex=rs("性别")
yin=rs("银两")
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
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
p{font-size:9pt; color:#000000}
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
<p align="center"><b><font style="font-size: 9pt" color="#ffee00">欢迎<%=my%>回到自己的小窝</font></b><br><br>
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
<td height="1" width="101" bgcolor="#FFFFFF" rowspan="10"><%if sex="女" then%><img src="image/logonan.jpg"><%end if%><%if sex="男" then%><img src="image/logonv.jpg"><%end if%></td>                 
  <td height="1" width="380" bgcolor="#C4DEFF" rowspan="10"> 
    <%if lx="豪华别墅" then%>
    <img src="image/35028_2.jpg" width="436" height="279"> 
    <%end if%>
  </td>                 
<td height="13" width="159" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">金库里有</font><font color="#ff0000"><%if sex="男" then%><%=rs("银两")%><%end if%><%if sex="女" then%><%=rs("伴侣银两")%><%end if%></font>两</p>
</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><%if sex="男" then%><form method="post" action="putmoney.asp"><%end if%><%if sex="女" then%><form method="post" action="putmoney.asp"><%end if%>
我要存<input type="TEXT" maxlength="10" value="0" name="money" style="width:80;border:1 solid #9a9999; font-size:9pt; background-color:#ffffff" size="10"><b>两</b></td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><INPUT TYPE="submit" VALUE="确定"><INPUT TYPE="reset" VALUE="重写"></td>                 
</tr></form>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><%if sex="男" then%><form method="post" action="putmoney.asp"><%end if%><%if sex="女" then%><form method="post" action="putmoney.asp"><%end if%>
我要取<input type="TEXT" maxlength="10" value="0" name="money" style="width:80;border:1 solid #9a9999; font-size:9pt; background-color:#ffffff" size="10"><b>两</b></td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF"><INPUT TYPE="submit" VALUE="确定"><INPUT TYPE="reset" VALUE="重写"></td>                 
</tr></form>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">现在的世道不好啊钱</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">还是放在家里比较安</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">全些。</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">　</td>                 
</tr>
<tr>
<td height="1" width="159" bgcolor="#C4DEFF">　</td>                 
</tr>
</table>
<table border=0 cellpadding=1 border=0 cellspacing=1 bgcolor="#51A8FF" width=678 height="1"> 
<tr>
<td height="164" width="102" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">：你回来了</font></p>
</td>                 
<td height="164" width="403" bgcolor="#C4DEFF">
  <p align="center"><font color="#000000">当前位置：金库</font></p>
</td>                 
<td height="164" width="157" bgcolor="#C4DEFF">
  <p align="center"><a href="../welcome.asp"><img src="../jhimg/return.gif" width="40" height="20" border="0"></a></p>
</td>
<td height="164" width="157" bgcolor="#C4DEFF">
<font color="#ff0000">您身上的银两：<%=yin%></font>
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

