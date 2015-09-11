<%
if Session("aqjh_name")="" then 
Response.Write "<script Language=Javascript>alert('你没有登陆江湖，或者已经断开连接!');window.close();</script>"
response.end
end if%>
<script lanaguage=javascript>
if(window.name!="aqjh_win"){ var i=1;while (i<=50){window.alert("你想作什么呀，你倒是刷啊！哈！慢慢点50次！！");i=i+1;}top.location.href="../../exit.asp"}
</script>
<HTML><HEAD><TITLE><%=Application("aqjh_chatroomname")%>-抗击倭寇</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="../../css.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY background=../../bg.gif leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<center><br><font color=green size=6 face="隶书">抗 击 倭 寇</font><br><br><br>
<table width="80%" border="1" cellspacing="2" cellpadding="2" bordercolor="#CC9933" align="center" style=font-size:9pt>
  <tr> 
    <td colspan="3"> 
      <div align="center"><img src="badman.gif" width="179" height="150"></div>
    </td>
  </tr>
  <tr> 
    <td rowspan="3" width="33%"><b>在这个动荡的年代里，在南方沿海总有些日本倭寇来骚扰我们，做为一名正义的侠客，我们赶快加入<font color="#FF0000">抗击倭寇</font>的战斗中来吧</b> 
    </td>
    <td width="40%" height="21"><font color="#996633"><b><font color="#CC0033">※初级侠客抗击倭寇战斗场※</font></b></font></td>
    <td width="27%" height="21"><a href="wokou1.asp" target="_self">踏入战场</a></td>
  </tr>
  <tr> 
    <td width="40%" height="23"><font color="#996633"><b><font color="#CC0033">※中级侠客抗击倭寇战斗场※</font></b></font></td>
    <td width="27%" height="23"><a href="wokou2.asp" target="_self">踏入战场</a></td>
  </tr>
  <tr> 
    <td width="40%"><font color="#996633"><b><font color="#CC0033">※高级侠客抗击倭寇战斗场※</font></b></font></td>
    <td width="27%"><a href="wokou3.asp" target="_self">踏入战场</a></td>
  </tr>
</table>
<p align=center><font style=font-size:9pt color=#000000>版权所有『爱情江湖总站』</font></p>
</BODY></HTML>