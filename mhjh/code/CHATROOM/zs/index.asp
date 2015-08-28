<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
my=session("yx8_mhjh_username")
if my="" then Response.Redirect "../error.asp?id=016"
if session("yx8_mhjh_inchat")="" then
Response.write "<script language='javascript'>alert ('你不能进入,请先进入聊天室再来。谢谢合作'); window.close();</script>"
Response.End 
        end if
%>
<HTML><HEAD><TITLE>绝世转生堂</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type><LINK 
href="jh/clubcom.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 5.0" name=GENERATOR></HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=0>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=596>
  <TBODY> 
  <TR>
    <TD background=jh/history_top_bg.gif vAlign=top width=596> 
      <TABLE align=center border=0 cellPadding=2 cellSpacing=0 class=p9 
      width="97%">
        <TBODY> 
        <TR>
          <TD height=25 width="31%">　</TD>
          <TD height=25 vAlign=top width="37%"> 
            <DIV align=center><font color="#FF0000" size="3">绝世转生堂</font></DIV>
          </TD>
          <TD height=25 vAlign=top 
width="32%">　</TD>
        </TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD background=jh/history_table_bg.gif height=158 vAlign=top width="596"> 
      <TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain 
      width=596>
        <TBODY> 
        <TR>
          <TD vAlign=top bgcolor="#000000"><BR>
            <TABLE align=center background=jh/table_bg.gif border=0 
            cellPadding=0 cellSpacing=0 class=p9 width="90%">
              <TBODY> 
              <TR>
                <TD align=middle vAlign=center> 
                  <P class=p9><font color="#FF0000"><b>【江湖提示】</b></font><br>
                          　　这儿是江湖中等级到100级,积分到400万的大侠转生的地方,目前,大侠可以转生四次,第一次转生为:转生人;第二次:强化人;第三次:觉醒人;第四次:真*无敌 
                          ;在这里转生后,积分和级别都变成0,但各种上限都大幅度增加,还将获得一些特殊能力,比如打穿别人的保护和隐身等等,自己去体会吧. 
                        </P>
                  <table width="550" border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr> 
                      <td width="207"><span style="font-size: 9pt">你有勇气转生吗？</span><font color="#FFFF00"><span style="font-size: 9pt">
                      </span> <a href="../../mp/zhuansheng.asp">
                      <style="font-size: 9pt">有！</font></td>
                      <td width="1"><align="right"><img border="0" src="JH/lg.gif"></td>
                    </tr>
                  </table>
                  <P class=p9>&nbsp; </P>
                  <P align=center class="p9"><a href="javascript:window.close()"><font color="#FFFFFF">关闭窗口</font></a></P>
                  </TD></TR></TBODY></TABLE>
            <P align=center>　</P></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=#847939 height=1 width="596"><IMG height=1 src="jh/page_point.gif" 
      width=1></TD>
  </TR></TBODY></TABLE></BODY></HTML>