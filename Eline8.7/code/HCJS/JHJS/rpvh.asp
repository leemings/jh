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
'Response.Write "<script language=javascript>{alert('�Բ��𣬳���ܾ����Ĳ���������\n     ��ȷ�����أ�');window.close();}</script>" 
'Response.End 
'end if
vhid=trim(Request.querystring("id"))
if not IsNumeric(vhid) then
	Response.Write "<script language=JavaScript>{alert('���ɣ�������ʲô���뵷���𣿣�');}</script>"
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
	Response.Write "<script language=JavaScript>{alert('��ʾ��������ʲô���뵷���𣿣�');window.close();}</script>"
	Response.End
end if
vhname=rs("a")
vhpic=rs("h")
vhnj=rs("i")
nickname=rs("g")
vhj=rs("j")
rs.close
rs.open "select c,h,j,m from b where a='"&vhname&"'",conn
if vhnj*2>rs("j") and not vhj then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('��ʾ���������ڵ�״̬���ã�����Ҫά�ޣ�');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
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
<HTML><HEAD><TITLE><%=Application("sjjh_chatroomname")%>--��ͨ����ά�޳�</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<BODY bgColor=#85C2E0 text=#ffffff topMargin=5>
<DIV align=center>
<p align="center">[ <font color="#FF3300" size=4><%=Application("sjjh_chatroomname")%>--��ͨ����ά�޳�</font> ]</p>
<font color="#FF00FF" size=4><a href="myvh.asp" target="_self"  title="װ����ͨ���߼����ƽ����˳������ҹ���">����������[װ������]ҳ��</a></font><br><br>
<table border="1" cellspacing="0" cellpadding="0" width="90%" align="center">
<tr><td width="180" align="center" valign="middle" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';" onmouseover="this.bgColor='#996632';" rowspan="2"><img src='images/<%=vhpic%>'><br>[<%=vhname%>]</td>
<td  align="center" valign="middle" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#996632';"> 
  <b><font color=blue>��ά�޳����ٷ�ά�޳���������֤��ͯ�Ų��ۣ�</font></b><br><br>
  ���<%if nickname<>"" then%>[<%=nickname%>]<%else%>[<%=vhname%>]<%end if%>ά����Ҫ����������<%=int(vhh*0.6+0.5)%>������ң�<%=int(vhm*0.6+0.5)%>��<br>
  ��֤�ܹ��ָ����ų������ϣ����Ǵﲻ�������˿����û����Ͷ�߹��أ���<br>
  <form  method="post" action="rpvhok.asp">
  <input type="hidden" name="type" value="1">
  <input type="hidden" name="id" value="<%=vhid%>">
  ��ô������ѡ�������ά�ްɣ���<input type="submit" name="submit" value="ȷ��">
  </form>
  </td></tr>
<tr><td  align="center" valign="middle" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#996632';"> 
 <% if vhc<>"������" and vhc<>"ս��" then%>
  <b><font color=blue>��˽��ά�޳��������ã��۸񹫵�����ӭ��ˣ�</font></b><br><br>
  ���<%if nickname<>"" then%>[<%=nickname%>]<%else%>[<%=vhname%>]<%end if%>ά����Ҫ����������<%=int(vhh*vhrnd+0.5)%>������ң�<%=int(vhm*vhrnd+0.5)%>��<br>
   һ���ܹ��ָ������������ϣ�������֤��������������20%���ʻᱻ�޻�����<br>
  <form  method="post" action="rpvhok.asp">
  <input type="hidden" name="type" value="2">
  <input type="hidden" name="tt" value="<%=vhtt%>">
  <input type="hidden" name="vhrnd" value="<%=vhrnd*100%>">
  <input type="hidden" name="id" value="<%=vhid%>">
  ��ô������ѡ�������ά�ްɣ���<input type="submit" name="submit" value="ȷ��">
  </form>
  <%end if%>
  </td></tr>
</table>
<br>
<font color="#FF00FF" size=4><a href="myvh.asp" target="_self" title="װ����ͨ���߼����ƽ����˳������ҹ���">����������[װ������]ҳ��</a></font><br><br></div>
<DIV align=center>
<br><br>
<a href='http://www.hc3w.net' target=_blank><font color="#ff00ff" size=4>���ͽ���</font></a><font color="#ff00ff" size=4>--��Ȩ���� ����--����</font></div>
</body>
</html>