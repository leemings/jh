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
    应我原江湖玩家的要求‘世纪联网’江湖程序正式更换为‘神话虚幻’江湖版本，以后也不会在更换。感谢大家的支持，我们会尽最大的努力把‘世纪联网’办好！
    <p>特别感谢K6论坛的 <font color="#000000"><b>coheike</b></font>&nbsp;&nbsp; 
    ken 金毛 紫枫 等朋友给我帮助！</p>
    <p>
    本程序是‘世纪联网’的修改版，时间匆忙 
    肯定还存在许多错误和漏洞，希望各位朋友发现漏洞和错误能及时提出，我们好给予及时的修改，关于江湖的所有问题，请到‘江湖论坛 
    <a href="http://bbs.tl99.com">http://bbs.tl99.com</a> ’上提出。<p>‘世纪联网’代理：至上思维工作室&nbsp; 
    和 山东吉通公司 江湖服务器支持ASP CGI PHP ACCESS数据库 一个虚拟目录 
    FTP权限 50人在线聊天&nbsp; 一年800元人民币 100人聊天 一年1200元 200人聊天 一年2000元人民币    江湖永久演示地址： 
    <font color="#FF0000">  http://xajh1.tl99.com</font>   在本站租用空间的都免费提供‘神话虚幻’江湖程序一套！ 租用1200元空间和2000元空间的另外赠送http://www.******.com 的域名一个，有需要江湖空间的用户请联系站长：新跃QQ：13575780 
    或者到我的江湖论坛发帖<a href="http://bbs.tl99.com" target="_blank">http://bbs.tl99.com</a>
    <p>下面是详细的<font color="#FF0000">汇款地址</font>：<font color="#FF0000">用户名：刘新跃 
    邮编：112000 地址：中国建设银行 辽宁省铁岭市支行 
    友谊储蓄所 帐号：0661110100100073711</font>
    <p>----<a href="http://www.tl99.com" target="_new">世纪联网</a>----
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
