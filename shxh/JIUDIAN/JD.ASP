<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file="dadata.asp"-->
<HTML><HEAD><TITLE>剑侠大酒店</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<SCRIPT src="jd.files/head.js"></SCRIPT>

<SCRIPT src="jd.files/the9.js"></SCRIPT>
<LINK href="jd.files/the9.css" rel=stylesheet>


<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff leftMargin=0 
text=#000000 topMargin=0>
<TABLE bgColor=#000000 border=0 cellPadding=0 cellSpacing=0 height=24 
  width=760 align="center">
  <TBODY> 
  <TR>
    <TD width=48><IMG src="jd.files/empty.gif"></TD>
    <TD width=34><IMG src="jd.files/position.gif" width="22" height="22"></TD>
    <TD class=head-l width=660>您当前位置：剑侠情缘-&gt;剑侠大酒店</TD>
    <TD align=right vAlign=top width=18><IMG 
  src="jd.files/rc.gif"></TD></TR></TBODY></TABLE>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=760 align="center">
  <TBODY> 
  <TR>
    <TD width=236><IMG height=119 src="jd.files/ggg.gif" width=236></TD>
    <TD background=jd.files/index_bg1.gif bgColor=#d70000 vAlign=top>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width=524>
        <TBODY> 
        <TR> 
          <TD height=8 width=120></TD>
          <TD width=62></TD>
          <TD width=47></TD>
          <TD width=100></TD>
          <TD width=74></TD>
          <TD width=42></TD>
          <TD width=79></TD>
        </TR>
        <TR> 
          <TD colspan="7">
            <div align="center"><img src="jd.files/tubiao.gif" width="300" height="50"></div>
          </TD>
        </TR>
        </TBODY> 
      </TABLE>
    </TD></TR></TBODY></TABLE>
<table border=0 cellpadding=0 cellspacing=0 width=760 align="center">
  <tbody> 
  <tr> 
    <td bgcolor=#d70000 height=379 valign=top width=180> 
      <table bgcolor=#d70000 border=0 cellpadding=0 cellspacing=0 width=180>
        <tbody> 
        <tr> 
          <td><img height=60 src="jd.files/index_t2.gif" width=149></td>
        </tr>
        <tr> 
          <td><a href=yiyi.asp
            ><img border=0 height=26 
            src="jd.files/index_left1.gif" width=149></a></td>
        </tr>
        <tr> 
          <td><a href=byby.asp
            ><img border=0 height=20 
            src="jd.files/index_left2.gif" width=149></a></td>
        </tr>
        <tr> 
          <td><a href=ssss.asp
            ><img border=0 height=19 
            src="jd.files/index_left3.gif" width=149></a></td>
        </tr>
        <tr> 
          <td><a href="smsm.asp" 
            ><img border=0 height=21 
            src="jd.files/index_left4.gif" width=149></a></td>
        </tr>
        <tr> 
          <td><a href=wcwc.asp 
            ><img border=0 height=20 
            src="jd.files/index_left5.gif" width=149></a></td>
        </tr>
        <tr> 
          <td><a href=lglg.asp 
            ><img border=0 height=19 
            src="jd.files/index_left6.gif" width=149></a></td>
        </tr>
        <tr> 
          <td><a href=lqlq.asp 
            ><img border=0 height=22 
            src="jd.files/index_left7.gif" width=149></a></td>
        </tr>
        <tr> 
          <td><a href=btbt.asp 
            ><img border=0 height=20 
            src="jd.files/index_left8.gif" width=149></a></td>
        </tr>
        <tr> 
          <td><a href=tctc.asp 
            ><img border=0 height=20 
            src="jd.files/index_left9.gif" width=149></a></td>
        </tr>
        <tr> 
          <td><img height=135 src="jd.files/index_t3.gif" 
        width=180></td>
        </tr>
        </tbody> 
      </table>
    </td>
    <td bgcolor=#faf4e9 valign=top width=39><img height=217 
      src="jd.files/index_xi.gif" width=39></td>
    <td bgcolor=#faf4e9 valign=top width=295><br>
      <table border=1 bgcolor="#FFFFFF" align=center width=100% cellpadding="0" cellspacing="1" bordercolor="#ff7010">
        <tr bgcolor="#006633"> 
          <td height="30" colspan="4" bgcolor="#FFFFFF"> 
            <p align=center class="p9"><font color="#FF0000"><img src="jd.files/djx.gif" width="150" height="25"></font></p>
          
        <tr bgcolor=#74E76D bordercolor="#000000"> 
          <td height="15" width="25%" bgcolor="#66FF66" bordercolor="#FF6600"> 
            <div align="center"><font size="2" color="#FF0000">宴会名</font></div>
          </td>
          <td height="15" width="24%" bgcolor="#66FF66" bordercolor="#FF6600"> 
            <div align="center"><font size="2" color="#FF0000">规格</font></div>
          </td>
          <td height="15" width="28%" bgcolor="#66FF66" bordercolor="#FF6600"> 
            <div align="center"><font size="2" color="#FF0000">售价</font></div>
          </td>
          <td height="15" width="23%" bgcolor="#66FF66" bordercolor="#FF6600"> 
            <div align="center"><font size="2" color="#FF0000">订购</font></div>
          </td>
        </tr>
        <!--#include file="dadata.asp"--> <%
sql="SELECT * FROM 宴会 where  类型='婚宴'"
Set Rs=connt.Execute(sql)
do while not rs.bof and not rs.eof
%> 
        <tr bgcolor=#DEAD63> 
          <td width="25%" bgcolor="#FFFFFF" class="calen-curr"><font size="2"> 
            <center>
              <%=rs("宴会名")%> 
            </center>
            </font> 
          <td width="24%" bgcolor="#FFFFFF" class="ct-def4"><font size="2"><%=rs("数量")%><span class="calen-curr">人参加</span></font></td>
          <td width="28%" bgcolor="#FFFFFF" class="calen-curr"><font size="2"><%=rs("售价")%> 
            两 </font></td>
          <td width="23%" bgcolor="#FFFFFF" class="calen-curr"><font size="2"> 
            <center>
              <font size="2"><a href=jiudian1.asp?id=<%=rs("id")%>><span class="calen-curr">定酒席</span></a></font> 
            </center>
            </font></td>
        </tr>
        <%
rs.movenext
loop
connt.close
set rs=nothing
%> 
      </table>
      <br>
      <img 
      border=0 height=44 name=Image90 src="jd.files/index_r1_c16.gif" 
      width=123> </td>
    <td background=jd.files/index_bg2.gif bgcolor=#faf4e9 width=223> 
      <table border="0" cellpadding="10" cellspacing="15" class="ct-def2" width="185">
        <tr bgcolor="#ffffff">
          <td height="150"><img height="1" src="jd.files/empty.gif" width="1">
 <div style="height: 37; width: 138">
              <marquee direction="up" height="150" onMouseOut="this.start()" onMouseOver="this.stop()" scrollamount="1" scrolldelay="30">在剑侠大酒店中，你可以定酒席，请你的亲朋好友来这参加你定下的酒席，酒席分为结婚宴，生日宴请功宴，请罪宴等等。不同的酒席，参加者所得到得东西是不同的，当然参加的人数也是不一样的。每个参加者可以获得体力，内力和精力的补充，参加不同的酒席可以获得不同程度的属性增加，前提是你的等级必须大于2级。</marquee>
              <br>
            </div>
          </td>
        </tr>
      </table>
    </td>
    <td bgcolor=#d70000 width=23>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor=#d70000 width=180><img height=44 src="jd.files/index_t4.gif" 
      width=149></td>
    <td bgcolor=#d70000 width=39>&nbsp;</td>
    <td bgcolor=#d70000 width=295>&nbsp;</td>
    <td bgcolor=#d70000 width=223>&nbsp;</td>
    <td bgcolor=#d70000 width=23>&nbsp;</td>
  </tr>
  </tbody> 
</table>
<TABLE border=0 cellPadding=0 cellSpacing=0 width=760 align="center">
  <TBODY> 
  <TR>
    <TD bgColor=#ff700d>&nbsp;</TD></TR></TBODY></TABLE>
</BODY></HTML>
