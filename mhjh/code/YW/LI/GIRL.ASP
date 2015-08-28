
<%
nickname=Session("yx8_mhjh_username")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("yx8_mhjh_connstr")
conn.open connstr
sql="select * from 用户 where 姓名='" & nickname & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	  Response.Redirect "../error.asp?id=016"
	conn.close
	response.end
else
dj=rs("等级")
yl=rs("银两")
xingbie=rs("性别")
if dj<3 then
	response.write "你还是江湖小辈，就想来这种地方"
	conn.close
	response.end
else
if xingbie="女" then
	response.write "丽院的姑娘不接待女的，请止步！"
	conn.close
	response.end

else %>
<!--#include file="jiu.asp"-->
<% id=request("id")
sql="select * from 名妓 where ID=" & id
Set Rs=connt.Execute(sql)
if rs.eof or rs.bof then
response.write "你有没有搞错呀，本院哪有这个姑娘呀，是不是想姑娘想疯了"
	connt.close
	response.end
else
mingji=rs("姓名")
meimao=rs("美貌度")

if yl<meimao*1.5  then
	response.write "没钱也想来这种高级的地方呀，请止步！"
            conn.close
	response.end
else
sql="update 用户 set 内力=内力+"&meimao&"/5,体力=体力-100,银两=银两-"&meimao&"*1.5,道德=道德-1000 where 姓名='"&nickname&"' "
conn.execute sql
connt.close
conn.close


end if
end if
end if
end if
end if%>
<html>

<head>
<title>-女人画廊三温暖</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
.unnamed1 {  font-size: 9pt}
-->
</style>

<bgsound src="sm.mid" loop="-1">

</head>

<body background="a15.gif" oncontextmenu=self.event.returnValue=false>
<table width="52%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td height="17" bgcolor="#0000FF" align="center"><font size="2" color="#FFFFFF">现在请接受你的女伴给你提供令你意想不到的@#$%*服务，请配合音乐观赏。</font></td>
  </tr>
  <tr bgcolor="#66FF66"> 
    <td align="center" height="378" bgcolor="#FFCC66"> 
      <p><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="550" height="400">
          <param name="_cx" value="14552">
          <param name="_cy" value="10583">
          <param name="Movie" value="girl.swf">
          <param name="Src" value="girl.swf">
          <param name="WMode" value="Window">
          <param name="Play" value="0">
          <param name="Loop" value="-1">
          <param name="Quality" value="High">
          <param name="SAlign" value>
          <param name="Menu" value="-1">
          <param name="Base" value>
          <param name="Scale" value="ShowAll">
          <param name="DeviceFont" value="0">
          <param name="EmbedMovie" value="0">
          <param name="BGColor" value>
          <param name="SWRemote" value>
          <param name="Stacking" value="below"><embed src="girl.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="550" height="400"></object><font> </font></p>
    </td>
  </tr>
  <tr bgcolor="#0033CC"> 
    <td align="center" height="13" class="unnamed1" style="font-family: Abadi MT Condensed Light"><font size="2" color="#FFFFFF">&nbsp;</font></td>
  </tr>
  <tr bgcolor="#0033CC"> 
    <td align="center" height="13" class="unnamed1" bgcolor="#FFFFFF"><a href="index.asp"><img border="0" src="ps.gif" alt="返回三温暖"></a></td>
  </tr>
</table>
</body>
</html>