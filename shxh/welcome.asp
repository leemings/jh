<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.open Application("BA_jxqy_connstr")
conn.execute "update 系统设置 set 属性值='"&Application("Ba_jxqy_visitor")&"' where 属性='访问人数'"
conn.close
set conn=nothing
%>

<head>
<link rel="stylesheet" href="chatroom/css.css">
<base target="about:blank">
</head>
<body oncontextmenu=self.event.returnValue=false background="<%=bgimage%>" bgcolor="<%=bgcolor%>">
<table width='100%' >
<tr>
<td>
	<br>
	<p>
	&nbsp;&nbsp;亲爱的侠客/侠女，你是否还在记得当初年轻时闯荡江湖那份豪迈，那份侠骨柔情，十几年匆匆而过，曾经的文字江湖，繁华不再，但你是否还会时常忆起？
	</p>
	<p>
	&nbsp;&nbsp;现快乐江湖重启，重现经典。江湖刚开，人气不多，正急切需要侠客您的光临，现邀请侠客归来吧……
	</p>
	<p>
	&nbsp;&nbsp;本人（站长）是曾经执迷于文字江湖的一名玩家，总想着回来玩玩神话虚幻版本的江湖，但是寻寻觅觅，总难以寻得一稳定的江湖，可以让我这份挚爱的感情得到延续，所以才有了今天的快乐江湖
	</p>
	<p>
	&nbsp;&nbsp;因为开这个江湖的目的，是为了能够更开心的与各位志同道合的侠客侠女把酒言欢，所以本江湖承诺稳定，公平公正，长期开放，不设收费，不求谋利，只求与大家一起打造一个真正快乐的江湖……
	</p>
	<p>
	&nbsp;&nbsp;现在还有部分功能兼容不了浏览器，会持续修正中，包括手机，也陆续修正成可以用uc浏览器访问
	</p>
	<br>
	<p>有任何意见可以找站长反馈，站长<font color="#FF0000"><b>qq865240608</b></font></p>
	<p>感谢大家的支持，我们会尽最大的努力把‘快乐江湖’办好！</p>

	<p>另外本江湖招收副站长，官府管理员，有意者联系站长。拉人有奖励哦</p>
    <p>----<a href="#" target="_new">快乐江湖</a>----
</td><td width=131>
<img src="images/STAMP.gif" width="131" height="300">
</td></tr></table>
  
<div align="center"><br>
  程序修改：<a href="http://www.mxcz.org/" target="_blank"><font color="#990000"> 
  <% response.write Application("Ba_jxqy_copyrightassert")%>
  </font></a><font color="#990000"></font>&nbsp;&nbsp; 授权给： <font color="#990000">  
  <% response.write Application("Ba_jxqy_userright")%> 
  &nbsp;&nbsp; </font>软件序列号： <font color="#990000">  
  <% response.write Application("Ba_jxqy_seriesnumber")%> 
  </font><br> 
  访问统计：<font color="#993300">   
          <%=Application("Ba_jxqy_visitor")%>   
          </font> 会员功能：<font color="#993300">   
          <%if Application("Ba_jxqy_fellow")=true then%> 
          开放 
          <%else%> 
          关闭 
          <%end if%> 
          </font>   
</div> 
</body> 
