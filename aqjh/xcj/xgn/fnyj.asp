<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("aqjh_jm")=session("aqjh_jm")+1
if session("aqjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
%>
<HTML><HEAD><TITLE>弓箭手职业技能—祝大家开心！^-^</TITLE>
<LINK href="lyy.css" rel=stylesheet></HEAD>
<BODY background=BG.gif oncontextmenu=self.event.returnValue=false>
<form method="POST" action="closeok.asp">
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=1 align=center border=0 height="1">
<TBODY><TR>
<TD width="11" height=15><IMG src="T1.gif" border=0></TD>
<TD width="373" background=M1.gif height=15>　</TD>
<TD width="244" height=15><IMG src="T2.gif" border=0></TD></TR>
  <TR><TD width="11" background=M2.gif rowSpan=3 height="1"></TD>
    <TD width="373" height=1><P align=center><SPAN lang=en>&copy;快乐江湖职业技能说明</P></SPAN></TD> 
    <TD width="244" background=M2.gif height=1 rowSpan=3></TD></TR> 
  <TR> 
    <TD width="373" height=1> 
      <HR color=#682420 SIZE=1> 
    </TD></TR> 
  <TR><TD width="373" height="1"> 
<!=> 
<table width="250" align="center" height="1"> 
          <tr> 
            <td height="1" width="102">  
<img border="0" src="hs1.gif">   
</td>  
            <td height="1" width="780">  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font face="宋体" color="#0000FF">弓箭手最强大的杀伤职能，能杀人于无形，由于是远距离攻击，所以命中率只有三分之一，但已经足够了，一旦命中，直接死亡，金钱全部变0，体力武功变0，希望大家能珍惜使用!^-^</font>   
</td>  
          </tr>  
          <tr>  
            <td height="1" width="882" colspan="2">   
<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
愤怒一击条件说明</p>
</td>  
          </tr>  
          <tr>  
            <td height="38" width="882" colspan="2">   
使对方立刻死亡，三分之一的几率，体力武功为0，金钱为0，需要耗100点知质，200点智力，金币50个！  
</td>  
          </tr>  
<TR><TD align=middle width="882" height="1" colspan="2"> <SPAN lang=en>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;    
    &copy;快乐江湖首创&nbsp;&nbsp;<a href="gjs.asp">返回上页</a></SPAN> </TD></TR>    
</table>    
<!=></TD></TR></form>    
<TR><TD width="11" height=14><IMG src="T3.gif" border=0></TD>    
    <TD width="373" background=M3.gif height=14>　</TD>    
    <TD width="244" height=14><IMG src="T4.gif"     
  border=0></TD></TR></TBODY></TABLE></FORM></BODY></HTML>

