<%@ LANGUAGE=VBScript codepage ="936" %>
<html><title>npc����</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666 oncontextmenu=window.event.returnValue=false ondragstart=window.event.returnValue=false>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
npcid=request.querystring("npcid")
npcid=cint(npcid)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
sqlstr="SELECT * FROM npc where id="&npcid
rs.Open sqlstr,conn
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('��Ǹ����Ҫ���ҵ�npc�����Ҳ�������鿴�Ƿ���ȷ��');history.back();</script>"
end if
%>
<form method=post action=npc_update.asp>
<input type=hidden size=4 name=id value="<%=npcid%>">
<br>�޸�˵����<br>1��[nͼƬ]������Ŀ¼ chat �ļ��С���ʽ���� npc/gui2.gif ���������ڵģ���<br>2��[n��Ʒ]��������npc����������õ�����Ʒ����ʽ���밴�գ����硰 �׷�|1;�Ż���¶��|1;����|1; ���������ڵģ���<br>�޸���ӣ�ע�⡰;��Ϊ��ǣ��ַ�֮���޿ո�<br>3������޸ġ�������Ʒ��ʱ������Ʒ�������ǽ������еģ����򣬡�������Ʒ���޷�ʹ�á�<br>4��[n��Ʒ����]���ȱ�����Ʒ�����͡�<br>w1ΪҩƷ�࣬w2Ϊ��ҩ�࣬w4Ϊ�����࣬w5Ϊ��Ƭ�࣬w6Ϊ��Ʒ�࣬w7Ϊ�ʻ��࣬w8Ϊ��ҩ�࣬w9Ϊħ���࣬w10����Ϊ�ࡣ<br>z1Ϊͷ����z2Ϊװ�Σ�z3Ϊ���ף�z4Ϊ���֣�z5Ϊ���ֺͶ��죬z6Ϊ˫�š�<br>5��������Ʒ����Ʒ���ͣ����������Ӧ�ģ����򣬴���npc�󣬵õ��ġ���Ʒ���ڡ�������Ʒ���о��޷���ʾ��<br>6��[n������]��ͼƬ��&lt;img src=npc/jiu2.gif width=50&gt;������Ŀ¼ chat �ļ��С���ʽ���� &lt;img src=npc/jiu2.gif width=50&gt; ���������ڵģ���<br><br><table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center"><tr><td colspan='6'><div align='center'>����ID:<%=npcid%>(<font color=blue size=-1>ע�⣺��������������Ӧ�����ݱ�����ͬ��������������пյ����ݴ���!----ԭ����û�еĿ��Դ�ո�!----</font>)</div></td></tr>
<%
Response.Write "<tr>"
for i=1 to rs.Fields.Count-1
strtype=rs.fields(i).Type
strname=rs.Fields(i).Name
if strname="����" then seename=rs.Fields(i).value
if rs.fields(i).Type=202 then strtype="�ַ�"
if rs.fields(i).Type=3 then strtype="����"
if rs.fields(i).Type=7 then strtype="����"
if rs.fields(i).Type=11 then strtype="�߼�"
if strname="grade" then strname="����ȼ�"
if strname="mvalue" then strname="�»���"
if strname="allvalue" then strname="�ܻ���"
if strname="times" then strname="��¼����"
if strname="regip" then strname="ע��ip"
if strname="lastip" then strname="���ip"
if strname="regtime" then strname="ע��ʱ��"
if strname="lasttime" then strname="���ʱ��"
	Response.Write "<td><div align='right'><font size=-1>"&strname&"(<font color=blue >"&strtype&"</font>)��<input type=text size='20' style='FONT-SIZE: 9pt;' name="&rs.Fields(i).Name&" value='"&rs.Fields(i).value&"'></font></div></td>"
	if i/3=int(i/3) then Response.Write "</tr><tr>"
	next
Response.Write "</tr>"
Response.Write "<tr><td colspan='6'><div align='center'><input type=submit value=���� name=submit> <input type=submit value=���� name=submit> <input type=submit value=ɾ�� name=submit> <input type=reset value=����></div></td></tr>"
Response.Write "</table></form>"
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>