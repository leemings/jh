<%
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�㲻�ܽ��в��������д˲���������������ң�');window.close();</script>"
	Response.End
end if
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>-��Ա������ҡ���-</title>
<link href="play.css" tppabs="play.css" rel="stylesheet">
</head>
<body topmargin="5" bgcolor="#FFFFFF" leftmargin="5" background="../../bg.gif">
<table border="0" cellspacing="0" cellpadding="0" width="600">
<tr>
<td></td>   
</tr>   
</table>   
���ֽ������л�Ա������ <font color="#FF0000">��������ҡ���Ա�𿨶�Ҫ��ʧһЩ�����򿴲����Ķ��������λ��ʿ����ѡ��</font><br>      <TABLE align=left border=1 borderColor=#cccccc borderColorDark=#ffffff          
borderColorLight=#000000 cellPadding=0 cellSpacing=0          
style="FONT-SIZE: 9pt" width="600" bordercolordard="#ffffff">          
<TBODY>          
<TR>          
<TD width="22%">          
<DIV align=center>����</DIV>         
</TD>         
<TD width="18%">         
<DIV align=center>��Ҫָ��</DIV>         
</TD>         
<TD width="20%">         
<DIV align=center>��������</DIV>         
</TD>         
<TD width="20%">         
<div align="center">����ָ��</div>         
</TD>         
<TD width="20%">         
<DIV align=center>�Ƿ�����</DIV>         
</TD>         
</TR>         
<TR>         
<TD width="22%">         
<DIV align=center><img border="0" src="../../chat/img/menoy.gif" width="50" height="40">  
  ��� </DIV>        
<DIV align=center></DIV>        
<DIV align=center></DIV>        
</TD>        
<TD width="18%">        
<DIV align=center>������10���<br>       
������10���</DIV>      
<DIV align=center>����-80��</DIV>     
<DIV align=center>��ľ-50</DIV>     
<DIV align=center>����-20</DIV>     
</TD>     
<TD width="20%">     
<DIV align=center>��ң�2</DIV>     
<DIV align=center></DIV>     
<DIV align=center></DIV>     
<DIV align=center></DIV>     
<DIV align=center></DIV>     
<DIV align=center></DIV>     
<DIV align=center></DIV>     
</TD>     
<TD width="20%">     
<div align="center">���飫0<br>     
������0</div>     
<div align="center">��Ա�ȼ�=1</div>     
<div align="center">�ȼ�����25</div>     
</TD>     
<TD width="20%">     
<DIV align=center><A     
href="dg1.asp"><img border="0" src="../../chat/img/pic15.gif" width="60" height="73">����</A> </DIV>     
</TD>     
<tr>   
<TD width="22%">        
<DIV align=center><img border="0" src="../../chat/img/menoy.gif" width="50" height="40"> 
  ��� </DIV>       
<DIV align=center></DIV>       
<DIV align=center></DIV>       
</TD>       
<TD width="18%">       
<DIV align=center>������20���<br>
  ������20���</DIV>    
<DIV align=center>����-100��</DIV>    
<DIV align=center>��ľ��-80</DIV>    
<DIV align=center>����-30</DIV>    
</TD>   
<TD width="20%">   
<DIV align=center>��ң�5</DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="20%">   
<div align="center">���飫0<br>   
������0</div>   
<div align="center">��Ա�ȼ�=2</div>   
<div align="center">�ȼ�С��400</div>   
</TD>   
<TD width="20%">   
<DIV align=center><A   
href="dg3.asp"><img border="0" src="../../chat/img/pic15.gif" width="60" height="73">����</A> </DIV>   
</TD>   
</tr> 
<tr> 
<TD width="22%">   
<DIV align=center><IMG src="jk.gif"> ��Ա�� </DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="18%">   
<DIV align=center>������50���</DIV>   
<DIV align=center>������10���</DIV>   
        <DIV align=center>����-500��</DIV>   
        <DIV align=center>��ľ��ˮ-100</DIV>   
        <DIV align=center>����-80</DIV>   
</TD>   
<TD width="20%">   
<DIV align=center>��Ա�𿨣�1</DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center>���+1</DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="20%">   
<div align="center">���飫0<br>   
������0</div>   
<div align="center">��Ա�ȼ�=2</div>   
<div align="center">�ȼ�С��400</div>   
</TD>   
<TD width="20%">   
<DIV align=center><A   
href="dg2.asp"><img border="0" src="../../chat/pic/dz23.gif" width="50" height="58">����</A> </DIV>   
</TD>   
</tr> 
<TR>   
<TD width="22%">   
<DIV align=center><IMG src="jk.gif"> ��Ա�� </DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="18%">   
<DIV align=center>������20���<br>
          �Ṧ��20���</DIV>   
<DIV align=center>����-150</DIV>   
<DIV align=center>����-200</DIV>   
<DIV align=center>����-2000��</DIV>   
</TD>   
<TD width="20%">   
<DIV align=center>��Ա�𿨣�4</DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center>���+5</DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="20%">   
<div align="center">���飫0</div>   
<div align="center">����+100<br>   
  ���£�2000</div>   
<div align="center">��Ա�ȼ�=4</div>   
<div align="center">��4��Ϊ���ѻ�Ա��</div>   
</TD>   
<TD width="20%">   
<DIV align=center><A   
href="dg4.asp"><img border="0" src="../../chat/pic/dz23.gif" width="50" height="58">����</A> </DIV>   
</TD></TBODY></TABLE></BODY></HTML>