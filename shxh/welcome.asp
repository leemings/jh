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
    Ӧ��ԭ������ҵ�Ҫ����������������������ʽ����Ϊ������á������汾���Ժ�Ҳ�����ڸ�������л��ҵ�֧�֣����ǻᾡ����Ŭ���ѡ�������������ã�
    <p>�ر��лK6��̳�� <font color="#000000"><b>coheike</b></font>&nbsp;&nbsp; 
    ken ��ë �Ϸ� �����Ѹ��Ұ�����</p>
    <p>
    �������ǡ��������������޸İ棬ʱ���æ 
    �϶��������������©����ϣ����λ���ѷ���©���ʹ����ܼ�ʱ��������Ǻø��輰ʱ���޸ģ����ڽ������������⣬�뵽��������̳ 
    <a href="http://bbs.tl99.com">http://bbs.tl99.com</a> ���������<p>��������������������˼ά������&nbsp; 
    �� ɽ����ͨ��˾ ����������֧��ASP CGI PHP ACCESS���ݿ� һ������Ŀ¼ 
    FTPȨ�� 50����������&nbsp; һ��800Ԫ����� 100������ һ��1200Ԫ 200������ һ��2000Ԫ�����    ����������ʾ��ַ�� 
    <font color="#FF0000">  http://xajh1.tl99.com</font>   �ڱ�վ���ÿռ�Ķ�����ṩ������á���������һ�ף� ����1200Ԫ�ռ��2000Ԫ�ռ����������http://www.******.com ������һ��������Ҫ�����ռ���û�����ϵվ������ԾQQ��13575780 
    ���ߵ��ҵĽ�����̳����<a href="http://bbs.tl99.com" target="_blank">http://bbs.tl99.com</a>
    <p>��������ϸ��<font color="#FF0000">����ַ</font>��<font color="#FF0000">�û���������Ծ 
    �ʱࣺ112000 ��ַ���й��������� ����ʡ������֧�� 
    ���괢���� �ʺţ�0661110100100073711</font>
    <p>----<a href="http://www.tl99.com" target="_new">��������</a>----
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
