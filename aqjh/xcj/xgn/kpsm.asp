<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("aqjh_jm")=session("aqjh_jm")+1
if session("aqjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
%>
<HTML><HEAD><TITLE>爱情江湖门派卡片栏—祝大家开心！^-^</TITLE>
<LINK href="lyy.css" rel=stylesheet></HEAD>
<BODY background=BG.gif oncontextmenu=self.event.returnValue=false>
<form method="POST" action="closeok.asp">
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=166 align=center border=0 height="76">
<TBODY><TR>
<TD width="11" height=15><IMG src="T1.gif" border=0></TD>
<TD width="373" background=M1.gif height=15>　</TD>
<TD width="244" height=15><IMG src="T2.gif" border=0></TD></TR>
  <TR><TD width="11" background=M2.gif rowSpan=3 height="47"></TD>
    <TD width="373" height=23><P align=center><SPAN lang=en>&copy;爱情江湖门派卡片说明区</P></SPAN></TD> 
    <TD width="244" background=M2.gif height=47 rowSpan=3></TD></TR> 
  <TR> 
    <TD width="373" height=1> 
      <HR color=#682420 SIZE=1> 
    </TD></TR> 
  <TR><TD width="373" height="28"> 
<!=> 
<table width="250" align="center" height="1"> 
          <tr> 
            <td height="1" width="441" colspan="3">  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="jqk.asp">加钱卡</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="sdk.asp">升点卡</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<a href="jjk.asp">加金卡</a>    
</td>   
          </tr>   
          <tr>   
            <td height="1" width="16">    
<img border="0" src="1.gif">    
</td>   
            <td height="1" width="89">    
<img border="0" src="2.gif">    
</td>   
            <td height="1" width="336">    
<img border="0" src="3.gif">    
</td>   
          </tr>   
          <tr>   
            <td height="1" width="16">    
<p align="left"></p> 
</td>   
            <td height="1" width="89">    
</td>   
            <td height="1" width="336">    
              <div align="center">   
                　   
              </div>   
</td>   
          </tr>   
<tr><td align="center" width="16" height="14">   
  </td><td align="center" width="89" height="14">   
  </td><td align="center" width="336" height="14">   
  </td></tr>   
<TR><TD align=middle width="441" height="1" colspan="3"> <SPAN lang=en>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
    &copy;爱情江湖首创&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;   
    <a href="sm.asp">返回首页</a></SPAN> </TD></TR>    
</table>    
<!=></TD></TR></form>    
<TR><TD width="11" height=14><IMG src="T3.gif" border=0></TD>    
    <TD width="373" background=M3.gif height=14>　</TD>    
    <TD width="244" height=14><IMG src="T4.gif"     
  border=0></TD></TR></TBODY></TABLE></FORM></BODY></HTML>