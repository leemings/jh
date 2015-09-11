<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("aqjh_jm")=session("aqjh_jm")+1
if session("aqjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
%>
<HTML><HEAD><TITLE>爱情方术师职能区—祝大家开心！^-^</TITLE>
<LINK href="lyy.css" rel=stylesheet></HEAD>
<BODY background=BG.gif oncontextmenu=self.event.returnValue=false>
<form method="POST" action="closeok.asp">
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=1 align=center border=0 height="1">
<TBODY><TR>
<TD width="11" height=15><IMG src="T1.gif" border=0></TD>
<TD width="373" background=M1.gif height=15>　</TD>
<TD width="244" height=15><IMG src="T2.gif" border=0></TD></TR>
  <TR><TD width="11" background=M2.gif rowSpan=3 height="1"></TD>
    <TD width="373" height=23><P align=center><SPAN lang=en>&copy;爱情江湖方术师职能说明区</P></SPAN></TD> 
    <TD width="244" background=M2.gif height=1 rowSpan=3></TD></TR> 
  <TR> 
    <TD width="373" height=1> 
      <HR color=#682420 SIZE=1> 
    </TD></TR> 
  <TR><TD width="373" height="1"> 
<!=> 
<table width="1" align="center" height="1"> 
          <tr> 
            <td height="1" width="415" colspan="3">  
&nbsp;&nbsp;<a href="kphc.asp">卡片合成</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="aqhc.asp">暗器合成&nbsp;</a>&nbsp;&nbsp;&nbsp;&nbsp; 
<a href="wqfy.asp">完全防御</a>     
</td>    
          </tr>    
          <tr>    
            <td height="1" width="109">     
<img border="0" src="hs4.gif">     
</td>    
            <td height="1" width="91">     
<img border="0" src="hs2.gif">     
</td>    
            <td height="1" width="310">     
<img border="0" src="hs3.gif">     
</td>    
          </tr>    
          <tr>    
            <td height="4" width="415" colspan="3">     
<p align="left">&nbsp;&nbsp;&nbsp; <font color="#0000FF">方术师是新兴的一门职业，拥有强大的守备能力能合成能力，能起到很大的辅助作用，但所耗的智力跟能源也不菲，希望这门职业能为玩家带来更大的娱乐性~谢谢~</font></p>   
</td>     
          </tr>     
          <tr>     
            <td height="4" width="415" colspan="3">      
<SPAN lang=en>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &copy;爱情江湖首创&nbsp;&nbsp;&nbsp;     
<a href="sm.asp">返回首页</a></SPAN>      
</td>       
          </tr>       
</table>        
<!=></TD></TR></form>        
<TR><TD width="11" height=1><IMG src="T3.gif" border=0></TD>        
    <TD width="373" background=M3.gif height=1>　</TD>        
    <TD width="244" height=1><IMG src="T4.gif"         
  border=0></TD></TR></TBODY></TABLE></FORM></BODY></HTML>