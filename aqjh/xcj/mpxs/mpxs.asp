<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("aqjh_jm")=session("aqjh_jm")+1
if session("aqjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
%>
<HTML><HEAD><TITLE>���齭������нˮ��ȡ��ף��ҿ��ģ�^-^</TITLE>
<LINK href="lyy.css" rel=stylesheet></HEAD>
<BODY background=BG.gif oncontextmenu=self.event.returnValue=false>
<form method="POST" action="closeok.asp">
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=135 align=center border=0 height="1">
<TBODY><TR>
<TD width="11" height=15><IMG src="T1.gif" border=0></TD>
<TD width="373" background=M1.gif height=15>��</TD>
<TD width="224" height=15><IMG src="T2.gif" border=0></TD></TR>
  <TR><TD width="11" background=M2.gif rowSpan=3 height="1"></TD>
    <TD width="373" height=23><P align=center><SPAN lang=en>&copy;��������нˮ��ȡ</P></SPAN></TD> 
    <TD width="224" background=M2.gif height=1 rowSpan=3></TD></TR> 
  <TR> 
    <TD width="373" height=1> 
      <HR color=#682420 SIZE=1> 
    </TD></TR> 
  <TR><TD width="373" height="1"> 
<!=> 
<table width="250" align="center" height="1"> 
          <tr>     
            <td height="1" width="441">      
&nbsp;&nbsp;&nbsp;�������Ҫ�����¼��ַ�ʽ��ȡнˮ��лл��ף�����ÿ��ģ�лл~       
</td>      
          </tr>      
          <tr>      
            <td height="10" width="441">       
              <div align="center">      
                <a href="Money1.asp">����</a>&nbsp;&nbsp; <a href="Money2.asp">�ٸ�</a>               
              </div>               
              <div align="center">                
                ��               
              </div>               
              <div align="center">                
                <a href="Money3.asp">����</a>&nbsp;&nbsp; <a href="Money4.asp">����</a>               
              </div>               
              <div align="center">                
                ��               
              </div>               
              <div align="center">                
                <a href="sm.asp">               
                ˵��</a>              
              </div>              
</td>               
          </tr>              
<TR><TD align=middle width="441" height="1"> <SPAN lang=en>&copy;���齭���״�&nbsp;&nbsp;&nbsp;      
    </SPAN> </TD></TR>             
</table>             
<!=></TD></TR></form>             
<TR><TD width="11" height=1><IMG src="T3.gif" border=0></TD>             
    <TD width="373" background=M3.gif height=1>��</TD>     
    <TD width="224" height=1><IMG src="T4.gif"      
  border=0></TD></TR></TBODY></TABLE></FORM></BODY></HTML>