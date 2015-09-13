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
<title>-会员修炼金币、金卡-</title>
<link href="play.css" tppabs="play.css" rel="stylesheet">
</head>
<body topmargin="5" bgcolor="#FFFFFF" leftmargin="5" background="../../bg.gif">
<table border="0" cellspacing="0" cellpadding="0" width="600">
<tr>
<td></td>   
</tr>   
</table>   
快乐江湖现有会员修炼： <font color="#FF0000">（修练金币、会员金卡都要损失一些看见或看不见的东西，请各位侠士慎重选择）</font><br>      <TABLE align=left border=1 borderColor=#cccccc borderColorDark=#ffffff          
borderColorLight=#000000 cellPadding=0 cellSpacing=0          
style="FONT-SIZE: 9pt" width="600" bordercolordard="#ffffff">          
<TBODY>          
<TR>          
<TD width="22%">          
<DIV align=center>类型</DIV>         
</TD>         
<TD width="18%">         
<DIV align=center>需要指数</DIV>         
</TD>         
<TD width="20%">         
<DIV align=center>增加数量</DIV>         
</TD>         
<TD width="20%">         
<div align="center">其它指数</div>         
</TD>         
<TD width="20%">         
<DIV align=center>是否修练</DIV>         
</TD>         
</TR>         
<TR>         
<TD width="22%">         
<DIV align=center><img border="0" src="../../chat/img/menoy.gif" width="50" height="40">  
  金币 </DIV>        
<DIV align=center></DIV>        
<DIV align=center></DIV>        
</TD>        
<TD width="18%">        
<DIV align=center>体力－10万点<br>       
内力－10万点</DIV>      
<DIV align=center>银两-80万</DIV>     
<DIV align=center>金木-50</DIV>     
<DIV align=center>智力-20</DIV>     
</TD>     
<TD width="20%">     
<DIV align=center>金币＋2</DIV>     
<DIV align=center></DIV>     
<DIV align=center></DIV>     
<DIV align=center></DIV>     
<DIV align=center></DIV>     
<DIV align=center></DIV>     
<DIV align=center></DIV>     
</TD>     
<TD width="20%">     
<div align="center">经验＋0<br>     
防御＋0</div>     
<div align="center">会员等级=1</div>     
<div align="center">等级大于25</div>     
</TD>     
<TD width="20%">     
<DIV align=center><A     
href="dg1.asp"><img border="0" src="../../chat/img/pic15.gif" width="60" height="73">修练</A> </DIV>     
</TD>     
<tr>   
<TD width="22%">        
<DIV align=center><img border="0" src="../../chat/img/menoy.gif" width="50" height="40"> 
  金币 </DIV>       
<DIV align=center></DIV>       
<DIV align=center></DIV>       
</TD>       
<TD width="18%">       
<DIV align=center>体力－20万点<br>
  内力－20万点</DIV>    
<DIV align=center>银两-100万</DIV>    
<DIV align=center>金木土-80</DIV>    
<DIV align=center>智力-30</DIV>    
</TD>   
<TD width="20%">   
<DIV align=center>金币＋5</DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="20%">   
<div align="center">经验＋0<br>   
防御＋0</div>   
<div align="center">会员等级=2</div>   
<div align="center">等级小于400</div>   
</TD>   
<TD width="20%">   
<DIV align=center><A   
href="dg3.asp"><img border="0" src="../../chat/img/pic15.gif" width="60" height="73">修练</A> </DIV>   
</TD>   
</tr> 
<tr> 
<TD width="22%">   
<DIV align=center><IMG src="jk.gif"> 会员金卡 </DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="18%">   
<DIV align=center>体力－50万点</DIV>   
<DIV align=center>法力－10万点</DIV>   
        <DIV align=center>银两-500万</DIV>   
        <DIV align=center>金木土水-100</DIV>   
        <DIV align=center>智力-80</DIV>   
</TD>   
<TD width="20%">   
<DIV align=center>会员金卡＋1</DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center>金币+1</DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="20%">   
<div align="center">经验＋0<br>   
防御＋0</div>   
<div align="center">会员等级=2</div>   
<div align="center">等级小于400</div>   
</TD>   
<TD width="20%">   
<DIV align=center><A   
href="dg2.asp"><img border="0" src="../../chat/pic/dz23.gif" width="50" height="58">修练</A> </DIV>   
</TD>   
</tr> 
<TR>   
<TD width="22%">   
<DIV align=center><IMG src="jk.gif"> 会员金卡 </DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="18%">   
<DIV align=center>法力－20万点<br>
          轻功－20万点</DIV>   
<DIV align=center>五行-150</DIV>   
<DIV align=center>智力-200</DIV>   
<DIV align=center>银两-2000万</DIV>   
</TD>   
<TD width="20%">   
<DIV align=center>会员金卡＋4</DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center>金币+5</DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
<DIV align=center></DIV>   
</TD>   
<TD width="20%">   
<div align="center">经验＋0</div>   
<div align="center">防御+100<br>   
  道德＋2000</div>   
<div align="center">会员等级=4</div>   
<div align="center">（4级为付费会员）</div>   
</TD>   
<TD width="20%">   
<DIV align=center><A   
href="dg4.asp"><img border="0" src="../../chat/pic/dz23.gif" width="50" height="58">修练</A> </DIV>   
</TD></TBODY></TABLE></BODY></HTML>