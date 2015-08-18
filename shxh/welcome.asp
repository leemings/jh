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
    <p>时间飞逝，日月穿梭，曾火爆的文字江湖，现已凋零，繁华不再。</p>
	<p>如今十几载匆匆而过，每每忆起当年纵横江湖的日子，甚是怀念，寻寻觅觅，却难寻得一合意江湖。</p>

	<p>现重新开启神话虚幻版江湖，本站承诺江湖永久经营，不会随意更换版本，对用户公平公正，只愿大家玩得开心。</p>
	<p>目前江湖初开，也许还存在许多错误和漏洞，希望各位朋友发现漏洞和错误能及时提出，我们好给予及时的修改。</p>
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
