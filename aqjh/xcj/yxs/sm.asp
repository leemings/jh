<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("aqjh_jm")=session("aqjh_jm")+1
if session("aqjh_jm")>30 then Response.Redirect "../chat/readonly/bomb.htm"
%>
<HTML><HEAD><TITLE>���ֽ�����н��ȡ��ף��ҿ��ģ�^-^</TITLE>
<LINK href="lyy.css" rel=stylesheet></HEAD>
<BODY background=BG.gif oncontextmenu=self.event.returnValue=false>
<form method="POST" action="closeok.asp">
<TABLE id=AutoNumber1 style="BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=135 align=center border=0 height="1">
<TBODY><TR>
<TD width="11" height=15><IMG src="T1.gif" border=0></TD>
<TD width="373" background=M1.gif height=15>��</TD>
<TD width="224" height=15><IMG src="T2.gif" border=0></TD></TR>
  <TR><TD width="11" background=M2.gif rowSpan=3 height="1"></TD>
    <TD width="373" height=23><P align=center><SPAN lang=en>&copy;���ֻ�Ա��нˮ��ȡ˵��</P></SPAN></TD> 
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
&nbsp;&nbsp;&nbsp;˵������н�ж�����ƣ�Ŀ�Ķ��������
<p><font color="#0000FF">��һ���𿨻��ҳ���4000���첻��</font></p>
<p><font color="#0000FF">�ڶ���ֻ��ÿ���µ�28�Ųſ����죬28�Ž��첻����нˮ</font></p>
<p><font color="#0000FF">������1����Ա��20���𿨣�2����3����Ա��50����4�����ѻ�Ա��150���</font></p>
<p><font color="#0000FF">���ģ������ȡ�Ǹ��ݽ����˼��ȼ�������ת��������ģ�������*4+�ȼ�+ת��*4 
���������»��ֳ���20��ʱ��һ��</font></p>
<p><font color="#0000FF">���壺������û�ﵽ15���ӵĲ�����</font></p>
</td>      
          </tr>      
<TR><TD align=middle width="441" height="1"> <SPAN lang=en>&copy;���ֽ����״�&nbsp;&nbsp;&nbsp;<a href="yxs.asp">������ҳ</a>     
    </SPAN> </TD></TR>            
</table>            
<!=></TD></TR></form>            
<TR><TD width="11" height=1><IMG src="T3.gif" border=0></TD>            
    <TD width="373" background=M3.gif height=1>��</TD>     
    <TD width="224" height=1><IMG src="T4.gif"      
  border=0></TD></TR></TBODY></TABLE></FORM></BODY></HTML>