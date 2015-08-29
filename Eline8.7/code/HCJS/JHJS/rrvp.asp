<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
if sjjh_name="" then Response.Redirect "../../error.asp?id=440"
http=Request.ServerVariables("HTTP_REFERER") 
'if InStr(http,"hcjs/jhjs")=0 then 
'Response.Write "<script language=javascript>{alert('对不起，程序拒绝您的操作！！！\n     按确定返回！');window.close();}</script>" 
'Response.End 
'end if
vhid=trim(Request.querystring("id"))
if not IsNumeric(vhid) then
	Response.Write "<script language=JavaScript>{alert('滚吧，你想做什么？想捣乱吗？！');}</script>"
	Response.End 
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
rs.open "SELECT a,g,h,i,j FROM vh where id="&vhid,conn
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('提示：你想做什么？想捣乱吗？！');window.close();}</script>"
	Response.End
end if
vhname=rs("a")
vhpic=rs("h")
vhnj=rs("i")
nickname=rs("g")
vhj=rs("j")
rs.close
rs.open "select c,h,j,m from b where a='"&vhname&"'",conn
randomize
vhrnd=int(rnd()*10)/100+0.28
vhtt=int(rs("m")/vhrnd+sqr(rs("h"))*vhrnd)
vhc=rs("c")
vhh=rs("h")
vhm=rs("m")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<HTML><HEAD><TITLE><%=Application("sjjh_chatroomname")%>--交通工具旧车市场</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<BODY bgColor=#85C2E0 text=#ffffff topMargin=5>
<DIV align=center>
<p align="center">[ <font color="#FF3300" size=4><%=Application("sjjh_chatroomname")%>--交通工具旧车市场</font> ]</p> 
<table border="1" cellspacing="0" cellpadding="0" width="90%" align="center"> 
<tr><td width="180" align="center" valign="middle" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';" onmouseover="this.bgColor='#996632';" rowspan="2"><img src='images/<%=vhpic%>'><br>[<%=vhname%>]</td> 
<td  align="center" valign="middle" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#996632';">  
  <b><font color=blue>本旧车市场属大中华旧车市场，价钱公道！童叟不欺！</font></b><br><br> 
  你的<%if nickname<>"" then%>[<%=nickname%>]<%else%>[<%=vhname%>]<%end if%>出售可得银两：<%=int(vhh*0.3)%>两；金币：<%=int(vhm*0.6+0.5)%>个<br> 
  一旦出售了，将不办理购回！（请考虑清楚啊！）<br> 
  <form  method="post" action="rrvpok.asp"> 
  <input type="hidden" name="type" value="1"> 
  <input type="hidden" name="id" value="<%=vhid%>"> 
  再见吧！我的宝贝？！<input type="submit" name="submit" value="确定"> 
  </form> 
  </td></tr> 
</table> 
<br> 
<font color="#FF00FF" size=4><a href="myvh.asp" target="_self" title="装备交通工具及定制进入退出聊天室公告">点击这儿返回[装备座驾]页面</a></font><br><br></div> 
<DIV align=center> 
<br><br> 
<font color="#ff00ff" >『E线江湖』</font></div> 
</body> 
</html> 
 
 
 
 
