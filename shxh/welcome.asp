<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
set conn=server.CreateObject("adodb.connection")
conn.open Application("BA_jxqy_connstr")
conn.execute "update ϵͳ���� set ����ֵ='"&Application("Ba_jxqy_visitor")&"' where ����='��������'"
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
    <p>ʱ����ţ����´������𱬵����ֽ��������ѵ��㣬�������١�</p>
	<p>���ʮ���شҴҶ�����ÿÿ�������ݺὭ�������ӣ����ǻ��ѰѰ���٣�ȴ��Ѱ��һ���⽭����</p>

	<p>�����¿�������ð潭������վ��ŵ�������þ�Ӫ��������������汾�����û���ƽ������ֻԸ�����ÿ��ġ�</p>
	<p>Ŀǰ����������Ҳ�������������©����ϣ����λ���ѷ���©���ʹ����ܼ�ʱ��������Ǻø��輰ʱ���޸ġ�</p>
	<p>���κ����������վ��������վ��<font color="#FF0000"><b>qq865240608</b></font></p>
	<p>��л��ҵ�֧�֣����ǻᾡ����Ŭ���ѡ��佭������ã�</p>
    <p>----<a href="#" target="_new">�佭��</a>----
</td><td width=131>
<img src="images/STAMP.gif" width="131" height="300">
</td></tr></table>
  
<div align="center"><br>
  �����޸ģ�<a href="http://www.mxcz.org/" target="_blank"><font color="#990000"> 
  <% response.write Application("Ba_jxqy_copyrightassert")%>
  </font></a><font color="#990000"></font>&nbsp;&nbsp; ��Ȩ���� <font color="#990000">  
  <% response.write Application("Ba_jxqy_userright")%> 
  &nbsp;&nbsp; </font>������кţ� <font color="#990000">  
  <% response.write Application("Ba_jxqy_seriesnumber")%> 
  </font><br> 
  ����ͳ�ƣ�<font color="#993300">   
          <%=Application("Ba_jxqy_visitor")%>   
          </font> ��Ա���ܣ�<font color="#993300">   
          <%if Application("Ba_jxqy_fellow")=true then%> 
          ���� 
          <%else%> 
          �ر� 
          <%end if%> 
          </font>   
</div> 
</body> 
