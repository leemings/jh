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
	<br>
	<p>
	&nbsp;&nbsp;�װ�������/��Ů�����Ƿ��ڼǵõ�������ʱ���������Ƿݺ������Ƿ��������飬ʮ����ҴҶ��������������ֽ������������٣������Ƿ񻹻�ʱ������
	</p>
	<p>
	&nbsp;&nbsp;�ֿ��ֽ������������־��䡣�����տ����������࣬��������Ҫ�������Ĺ��٣����������͹����ɡ���
	</p>
	<p>
	&nbsp;&nbsp;���ˣ�վ����������ִ�������ֽ�����һ����ң������Ż�����������ð汾�Ľ���������ѰѰ���٣�������Ѱ��һ�ȶ��Ľ����������������ֿ���ĸ���õ����������Բ����˽���Ŀ��ֽ���
	</p>
	<p>
	&nbsp;&nbsp;��Ϊ�����������Ŀ�ģ���Ϊ���ܹ������ĵ����λ־ͬ���ϵ�������Ů�Ѿ��Ի������Ա�������ŵ�ȶ�����ƽ���������ڿ��ţ������շѣ�����ı����ֻ������һ�����һ���������ֵĽ�������
	</p>
	<p>
	&nbsp;&nbsp;���ڻ��в��ֹ��ܼ��ݲ��������������������У������ֻ���Ҳ½�������ɿ�����uc���������
	</p>
	<br>
	<p>���κ����������վ��������վ��<font color="#FF0000"><b>qq865240608</b></font></p>
	<p>��л��ҵ�֧�֣����ǻᾡ����Ŭ���ѡ����ֽ�������ã�</p>

	<p>���Ȿ�������ո�վ�����ٸ�����Ա����������ϵվ���������н���Ŷ</p>
    <p>----<a href="#" target="_new">���ֽ���</a>----
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
