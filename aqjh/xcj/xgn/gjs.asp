<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("aqjh_jm")=session("aqjh_jm")+1
if session("aqjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
%>
<HTML><HEAD><TITLE>���鹭����ְ������ף��ҿ��ģ�^-^</TITLE>
<LINK href="lyy.css" rel=stylesheet></HEAD>
<BODY background=BG.gif oncontextmenu=self.event.returnValue=false>
<form method="POST" action="closeok.asp">
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=1 align=center border=0 height="1">
<TBODY><TR>
<TD width="11" height=15><IMG src="T1.gif" border=0></TD>
<TD width="373" background=M1.gif height=15>��</TD>
<TD width="244" height=15><IMG src="T2.gif" border=0></TD></TR>
  <TR><TD width="11" background=M2.gif rowSpan=3 height="1"></TD>
    <TD width="373" height=23><P align=center><SPAN lang=en>&copy;���齭��������ְ��˵����</P></SPAN></TD> 
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
&nbsp;&nbsp;<a href="fnyj.asp">��ŭһ��</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="ksbg.asp">���ٲ���</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<a href="ygtx.asp">�������</a>     
</td>    
          </tr>    
          <tr>    
            <td height="1" width="109">     
<img border="0" src="hs1.gif">     
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
<p align="left">&nbsp;&nbsp;&nbsp;   
�����������˵�һ��ְҵ��ӵ��ǿ��Ĺ����������ٵĲ�������������ս���з��Ӻܴ�����ã������ĵľ�������ԴҲ���ƣ�ϣ������ְҵ��Ϊ��Ҵ��������������~лл~</p>   
</td>     
          </tr>     
          <tr>     
            <td height="4" width="415" colspan="3">      
<SPAN lang=en>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &copy;���齭���״�&nbsp;&nbsp;&nbsp;    
<a href="sm.asp">������ҳ</a></SPAN>     
</td>      
          </tr>      
</table>       
<!=></TD></TR></form>       
<TR><TD width="11" height=1><IMG src="T3.gif" border=0></TD>       
    <TD width="373" background=M3.gif height=1>��</TD>       
    <TD width="244" height=1><IMG src="T4.gif"        
  border=0></TD></TR></TBODY></TABLE></FORM></BODY></HTML>