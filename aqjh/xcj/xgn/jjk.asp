<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("aqjh_jm")=session("aqjh_jm")+1
if session("aqjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
%>
<HTML><HEAD><TITLE>升点卡说明—祝大家开心！^-^</TITLE>
<LINK href="lyy.css" rel=stylesheet></HEAD>
<BODY background=BG.gif oncontextmenu=self.event.returnValue=false>
<form method="POST" action="closeok.asp">
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=166 align=center border=0 height="1">
<TBODY><TR>
<TD width="11" height=15><IMG src="T1.gif" border=0></TD>
<TD width="373" background=M1.gif height=15>　</TD>
<TD width="244" height=15><IMG src="T2.gif" border=0></TD></TR>
  <TR><TD width="11" background=M2.gif rowSpan=3 height="1"></TD>
    <TD width="373" height=23><P align=center><SPAN lang=en>&copy;快乐江湖加金卡</P></SPAN></TD> 
    <TD width="244" background=M2.gif height=1 rowSpan=3></TD></TR> 
  <TR> 
    <TD width="373" height=1> 
      <HR color=#682420 SIZE=1> 
    </TD></TR> 
  <TR><TD width="373" height="1"> 
<!=> 
<table width="250" align="center" height="1"> 
          <tr> 
            <td height="1" width="441">  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font face="宋体" color="#0000FF">加金卡在卡片屋也是买不到的，这个卡片的创立也是考虑奖励官府跟一些奖励的需要，，每个月评最优掌门，奖励金币或者加金卡，希望各位掌门能对每位弟子负责，做到更好~谢谢!^-^</font>   
</td>  
          </tr>  
          <tr>  
            <td height="1" width="441">   
<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
加金卡限制说明</p>
</td>  
          </tr>  
          <tr>  
            <td height="8" width="441">   
加金卡加金卡100块，本身金卡大于1000，金币大于2000的都不能使用，只有官府跟掌门能使用！  
</td>  
          </tr>  
<TR><TD align=middle width="441" height="1"> <SPAN lang=en>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
    &copy;快乐江湖首创&nbsp;&nbsp;<a href="kpsm.asp">返回上页</a></SPAN> </TD></TR>    
</table>    
<!=></TD></TR></form>    
<TR><TD width="11" height=14><IMG src="T3.gif" border=0></TD>    
    <TD width="373" background=M3.gif height=14>　</TD>    
    <TD width="244" height=14><IMG src="T4.gif"     
  border=0></TD></TR></TBODY></TABLE></FORM></BODY></HTML>