<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
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
conn.open Application("aqjh_usermdb")
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
<HTML><HEAD><TITLE><%=Application("aqjh_chatroomname")%>--��ͨ���߾ɳ��г�</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<BODY bgColor=#85C2E0 text=#ffffff topMargin=5>
<DIV align=center>
<p align="center">[ <font color="#FF3300" size=4><%=Application("aqjh_chatroomname")%>--��ͨ���߾ɳ��г�</font> ]</p> 
<table border="1" cellspacing="0" cellpadding="0" width="90%" align="center"> 
<tr><td width="180" align="center" valign="middle" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';" onmouseover="this.bgColor='#996632';" rowspan="2"><img src='images/<%=vhpic%>'><br>[<%=vhname%>]</td> 
<td  align="center" valign="middle" bgcolor="#85C2E0" onmouseout="this.bgColor='#85C2E0';"onmouseover="this.bgColor='#996632';">  
  <b><font color=blue>���ɳ��г������л��ɳ��г�����Ǯ������ͯ�Ų��ۣ�</font></b><br><br> 
  ���<%if nickname<>"" then%>[<%=nickname%>]<%else%>[<%=vhname%>]<%end if%>���ۿɵ�������<%=int(vhh*0.3)%>������ң�<%=int(vhm*0.6+0.5)%>��<br> 
  һ�������ˣ����������أ����뿼�����������<br> 
  <form  method="post" action="rrvpok.asp"> 
  <input type="hidden" name="type" value="1"> 
  <input type="hidden" name="id" value="<%=vhid%>"> 
  �ټ��ɣ��ҵı�������<input type="submit" name="submit" value="ȷ��"> 
  </form> 
  </td></tr> 
</table> 
<br> 
<font color="#FF00FF" size=4><a href="myvh.asp" target="_self" title="װ����ͨ���߼����ƽ����˳������ҹ���">����������[װ������]ҳ��</a></font><br><br></div> 
<DIV align=center> 
<br><br> 
<font color="#ff00ff" >�����齭����</font></div> 
</body> 
</html> 