<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("aqjh_jm")=session("aqjh_jm")+1
if session("aqjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
%>
<HTML><HEAD><TITLE>快乐江湖月薪领取—祝大家开心！^-^</TITLE>
<LINK href="lyy.css" rel=stylesheet></HEAD>
<BODY background=BG.gif oncontextmenu=self.event.returnValue=false>
<form method="POST" action="closeok.asp">
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=135 align=center border=0 height="1">
<TBODY><TR>
<TD width="11" height=15><IMG src="T1.gif" border=0></TD>
<TD width="373" background=M1.gif height=15>　</TD>
<TD width="224" height=15><IMG src="T2.gif" border=0></TD></TR>
  <TR><TD width="11" background=M2.gif rowSpan=3 height="1"></TD>
    <TD width="373" height=23><P align=center><SPAN lang=en>&copy;快乐会员月薪水领取说明</P></SPAN></TD> 
    <TD width="224" background=M2.gif height=1 rowSpan=3></TD></TR> 
  <TR> 
    <TD width="373" height=1> 
      <HR color=#682420 SIZE=1> 
    </TD></TR> 
  <TR><TD width="373" height="1"> 
<!=> 
<table width="250" align="center" height="1"> 
          <tr>     
            <td height="11" width="441">      
&nbsp;&nbsp;&nbsp;说明：月薪有多个限制，目的多个，如下
<p><font color="#0000FF">第一：金卡或金币超过4000的领不了</font></p>
<p><font color="#0000FF">第二：只有每个月的28号才可以领，28号将领不了日薪水</font></p>
<p><font color="#0000FF">第三：1级会员领20个金卡，2级和3级会员领50个，4级付费会员领150块金卡</font></p>
<p><font color="#0000FF">第四：金币领取是根据介绍人及等级，还有转生数计算的：介绍人*4+等级+转生*4 
当介绍人月积分超过20万时算一人</font></p>
<p><font color="#0000FF">第五：刚上线没达到15分钟的不能领</font></p>
</td>      
          </tr>      
<TR><TD align=middle width="441" height="1"> <SPAN lang=en>&copy;快乐江湖首创&nbsp;&nbsp;&nbsp;<a href="yxs.asp">返回上页</a>     
    </SPAN> </TD></TR>            
</table>            
<!=></TD></TR></form>            
<TR><TD width="11" height=1><IMG src="T3.gif" border=0></TD>            
    <TD width="373" background=M3.gif height=1>　</TD>     
    <TD width="224" height=1><IMG src="T4.gif"      
  border=0></TD></TR></TBODY></TABLE></FORM></BODY></HTML>