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
<title>-�����书����-</title>
<link href="play.css" tppabs="play.css" rel="stylesheet">
</head>
<body topmargin="5" bgcolor="#000000" leftmargin="5">
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<font size="3">�����״����ɾ�ѧ����ʥ��</font></font></p>
      <TABLE align=left border=1 borderColor=#cccccc borderColorDark=#ffffff          
borderColorLight=#000000 cellPadding=0 cellSpacing=0          
style="FONT-SIZE: 9pt" width="600" bordercolordard="#ffffff">          
<TBODY>          
<TR>          
<TD width="22%">          
<DIV align=center><font color="#FF0000">����</font></DIV>         
</TD>         
<TD width="20%">         
<DIV align=center><font color="#FF0000">�Ƿ�����</font></DIV>         
</TD>         
</TR>         
<TR>         
<TD width="22%">         
<DIV align=center>  
  <font color="#FF0000">  
  ��һ�׶�</font> </DIV>        
<DIV align=center>  
  <font color="#FF0000">�������������ɵ�һ���ķ���������</font> </DIV>        
<DIV align=center>  
  <font color="#FF0000">���ɹ��׶�100�����ϣ����������������������󷽿�������Ŭ���������ɳ�Ϊ�����еĸ���</font> </DIV>        
</TD>        
<TD width="20%">     
<DIV align=center><a href="dg1.asp"><img border="0" src="1-2.jpg"><font color="#0000FF">����</font></a> </DIV>     
</TD>     
<tr>   
<TD width="22%">        
<DIV align=center><font color="#FF0000">�ڶ��׶�</font></DIV>       
<DIV align=center><font color="#FF0000">�������������ɵڶ����ķ���������</font></DIV>       
<DIV align=center><font color="#FF0000">���ɹ��׶�500�����ϣ��������������ߣ���һ���Ļ۸��������죬ͬʱ��Ҫ������������������С�㣩30��</font></DIV>       
</TD>       
<TD width="20%">   
<DIV align=center><A     
href="dg1.asp"><img border="0" src="1-1.jpg"></A><a href="dg2.asp"><font color="#0000FF">����</font></a> </DIV>   
</TD>   
</tr> 
<tr> 
<TD width="22%">   
<DIV align=center>�� </DIV>   
<DIV align=center><font color="#FF0000">�����׶�</font> </DIV>   
<DIV align=center><font color="#FF0000">�������������ɵ������ķ���������</font> </DIV>   
<DIV align=center><font color="#FF0000">���ɹ��׶�2000���������������ߣ����ǣ��۸���Խ������Խ�죬ͬʱ��Ҫ������������������С�㣩100��</font> </DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="20%">   
<DIV align=center><A     
href="dg1.asp"><img border="0" src="1-4.jpg"></A><a href="dg3.asp"><font color="#0000FF">����</font></a> </DIV>   
</TD>   
</tr> 
<TR>   
<TD width="22%">   
<DIV align=center>�� </DIV>   
<DIV align=center><font color="#FF0000">Ͷ�ݸ���</font> </DIV>   
<DIV align=center><font color="#FF0000">���޽׶η֣�����������֮������������</font> </DIV>   
<DIV align=center><font color="#FF0000">��ֹ��������ʿ��������</font> </DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="20%">   
<DIV align=center><A     
href="dg1.asp"><img border="0" src="lg.gif"></A><a href="dg4.asp"><font color="#0000FF">����</font></a> </DIV>   
</TD></TBODY></TABLE></BODY></HTML>