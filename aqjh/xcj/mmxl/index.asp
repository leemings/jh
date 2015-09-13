<%
aqjh_name=Session("aqjh_name")
if aqjh_name="" then Response.Redirect "error.asp?id=440"
if Instr(LCase(Application("aqjh_useronlinename"&session("nowinroom")))," "&LCase(aqjh_name)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能进行操作，进行此操作必须进入聊天室！');window.close();</script>"
	Response.End
end if
%>
<html><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>-门派武功修炼-</title>
<link href="play.css" tppabs="play.css" rel="stylesheet">
</head>
<body topmargin="5" bgcolor="#000000" leftmargin="5">
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#FF0000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
<font size="3">快乐首创门派绝学修炼圣地</font></font></p>
      <TABLE align=left border=1 borderColor=#cccccc borderColorDark=#ffffff          
borderColorLight=#000000 cellPadding=0 cellSpacing=0          
style="FONT-SIZE: 9pt" width="600" bordercolordard="#ffffff">          
<TBODY>          
<TR>          
<TD width="22%">          
<DIV align=center><font color="#FF0000">类型</font></DIV>         
</TD>         
<TD width="20%">         
<DIV align=center><font color="#FF0000">是否修练</font></DIV>         
</TD>         
</TR>         
<TR>         
<TD width="22%">         
<DIV align=center>  
  <font color="#FF0000">  
  第一阶段</font> </DIV>        
<DIV align=center>  
  <font color="#FF0000">（江湖各大门派第一层心法修炼处）</font> </DIV>        
<DIV align=center>  
  <font color="#FF0000">门派贡献度100万以上，符合体力内力等修炼需求方可修炼，努力点修炼可成为高手中的高手</font> </DIV>        
</TD>        
<TD width="20%">     
<DIV align=center><a href="dg1.asp"><img border="0" src="1-2.jpg"><font color="#0000FF">修练</font></a> </DIV>     
</TD>     
<tr>   
<TD width="22%">        
<DIV align=center><font color="#FF0000">第二阶段</font></DIV>       
<DIV align=center><font color="#FF0000">（江湖各大门派第二层心法修炼处）</font></DIV>       
<DIV align=center><font color="#FF0000">门派贡献度500万以上，体力内力充沛者，有一定的慧根修炼方快，同时需要闯荡江湖天数（工资小点）30点</font></DIV>       
</TD>       
<TD width="20%">   
<DIV align=center><A     
href="dg1.asp"><img border="0" src="1-1.jpg"></A><a href="dg2.asp"><font color="#0000FF">修练</font></a> </DIV>   
</TD>   
</tr> 
<tr> 
<TD width="22%">   
<DIV align=center>　 </DIV>   
<DIV align=center><font color="#FF0000">第三阶段</font> </DIV>   
<DIV align=center><font color="#FF0000">（江湖各大门派第三层心法修炼处）</font> </DIV>   
<DIV align=center><font color="#FF0000">门派贡献度2000万，体力内力充沛者，根骨（慧根）越高修炼越快，同时需要闯荡江湖天数（工资小点）100点</font> </DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="20%">   
<DIV align=center><A     
href="dg1.asp"><img border="0" src="1-4.jpg"></A><a href="dg3.asp"><font color="#0000FF">修练</font></a> </DIV>   
</TD>   
</tr> 
<TR>   
<TD width="22%">   
<DIV align=center>　 </DIV>   
<DIV align=center><font color="#FF0000">投拜高人</font> </DIV>   
<DIV align=center><font color="#FF0000">（无阶段分，乃无门无派之侠客求炼处）</font> </DIV>   
<DIV align=center><font color="#FF0000">禁止各大派人士进入修炼</font> </DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="20%">   
<DIV align=center><A     
href="dg1.asp"><img border="0" src="lg.gif"></A><a href="dg4.asp"><font color="#0000FF">修练</font></a> </DIV>   
</TD></TBODY></TABLE></BODY></HTML>