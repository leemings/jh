<!--#include file=conn.asp-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="z_visual_const.asp" -->
<%
dim action
action=request.querystring("action")
%>
<html>
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	call main()
end if

sub main()
dim rs1,rs2,curType1,curType2,curseqno1,curseqno2%>
<table cellpadding="2" cellspacing="1" border="0" width="95%" class="tableBorder" align=center>
<tr>
<th align=left colspan=7 height=23>�����������</th>
</tr>
<tr>
<td width=20% class=forumrow>ע������</td>
<td width=80% class=forumrow colspan=6>���޹��߱�����ʹ��Ʒ�������û����ɼ���<br>
  �ڵ���Ʒ��������Ӧ�����ϲ�����<br>
  �۵��������Ʒ�����ǰ�Ѿ�����ϵͳ���Զ�������<br>
  �������޸ĵ������뾡���򵥣�<br>
  ��<font color="#FF0000">����ϸ�Ķ���ز�����˵��</font><br>
	<form action="?" method=post><input type=hidden name=action value=recalc><input type=Submit value="ˢ��������Ʒ" name="Submit"></form></td>
</tr>
<form action="?action=visualSearch" method=post>
<tr>
<td width=20% class=forumrow>��������</td>
<td width=80% class=forumrow colspan=6>
<select size=1 name="visualSearch" onchange="javascript:submit()">
<option value="0">��ѡ���ѯ����</option>
<option value="8" <%if request("visualSearch")=8 then%>selected<%end if%>>�г�������Ʒ</option>
<option value="1" <%if request("visualSearch")=1 then%>selected<%end if%>>�г����б�����Ʒ</option>
<option value="2" <%if request("visualSearch")=2 then%>selected<%end if%>>�г�����ר������Ա����Ʒ</option>
<option value="3" <%if request("visualSearch")=3 then%>selected<%end if%>>�г�����ר��������������Ʒ</option>
<option value="4" <%if request("visualSearch")=4 then%>selected<%end if%>>�г�����ר����������Ʒ</option>
<option value="5" <%if request("visualSearch")=5 then%>selected<%end if%>>�г�����ר��VIP��Ա����Ʒ</option>
<option value="6" <%if request("visualSearch")=6 then%>selected<%end if%>>�г����н����ϼܵ���Ʒ</option>
<option value="7" <%if request("visualSearch")=7 then%>selected<%end if%>>�г����п��Ϊ�����Ʒ</option>
</select>
</td>
</tr>
</form>
<%if request("action")="" or request("visualSearch")="0" then%>
<form action="?action=visualSearch" method=post>
<tr>
<th align=left colspan=7 height=23>�߼���ѯ</th>
</tr>
<tr>
<td width=20% class=forumrow>ע������</td>
<td width=80% class=forumrow colspan=6>�ڼ�¼�ܶ���������������Խ���ѯԽ�����뾡�����ٲ�ѯ�����������ʾ��¼��Ҳ����ѡ�����</td>
</tr>
<tr>
<td width=20% class=forumrow>�����ʾ��¼��</td>
<td width=80% class=forumrow colspan=6><input size=45 name="searchMax" type=textbox value=100></td>
</tr>
<tr>
<td width=20% class=forumrow>��Ʒ���</td>
<td width=80% class=forumrow colspan=6><input size=45 name="visualid" type=textbox></td>
</tr>
<tr>
<td width=20% class=forumrow>��Ʒ����</td>
<td width=80% class=forumrow colspan=6><input size=45 name="visualname" type=textbox>&nbsp;<input type=checkbox name="visualnamechk" value="yes" checked>��Ʒ��������ƥ��</td>
</tr>
<tr>
<td width=20% class=forumrow>��Ʒ���</td>
<td width=80% class=forumrow colspan=6>
<select size=1 name="visualtype">
<option value=0>����</option>
<%
set rs1=server.createobject("ADODB.recordset")
set rs2=server.createobject("ADODB.recordset")
rs1.open "select * from visual_type order by seqno",conn,1,1
do while not rs1.eof
	curseqno1=rs1("seqno")
	curtype1=rs1("typename")
	rs2.open "select * from visual_subtype where typeseq="&curseqno1&" order by seqno",conn,1,1
	rs1.movenext
	if not rs1.eof then
		response.write "<option value=00"&curseqno1&">��"&curtype1&"</option>"
	else
		response.write "<option value=00"&curseqno1&">��"&curtype1&"</option>"
	end if
	do while not rs2.eof
		curseqno2=rs2("seqno")
		curtype2=rs2("subtypename")
		rs2.movenext
		if not rs2.eof then
			if not rs1.eof then
				response.write "<option value="&curseqno2&"0"&curseqno1&">��&nbsp;&nbsp;��"&curtype2&"</option>"
			else
				response.write "<option value="&curseqno2&"0"&curseqno1&">��&nbsp;&nbsp;��"&curtype2&"</option>"
			end if
		else
			if not rs1.eof then
				response.write "<option value="&curseqno2&"0"&curseqno1&">��&nbsp;&nbsp;��"&curtype2&"</option>"
			else
				response.write "<option value="&curseqno2&"0"&curseqno1&">��&nbsp;&nbsp;��"&curtype2&"</option>"
			end if
		end if
	loop
	rs2.close
loop
rs1.close
set rs2=nothing
set rs1=nothing%>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>�����Ա�</td>
<td width=80% class=forumrow colspan=6>
<select size=1 name="visualsex">
<option value=9>����</option>
<option value=0>Ů</option>
<option value=1>��</option>
<option value=2>����</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>�޹���</td>
<td width=80% class=forumrow colspan=6>
<select size=1 name="visualflag">
<option value=9>����</option>
<option value=0>������Ʒ</option>
<option value=1>���л�Ա</option>
<option value=2>��������</option>
<option value=3>��������</option>
<option value=4>����Ա</option>
</select>
</td>
</tr>
<tr>
<td width=20% class=forumrow>�Ƿ�VIPר��</td>
<td width=80% class=forumrow colspan=6><input type=radio name="visualvip" value="2" checked>&nbsp;����&nbsp;&nbsp;<input type=radio name="visualvip" value="0">&nbsp;��&nbsp;&nbsp;<input type=radio name="visualvip" value="1">&nbsp;��</td>
</tr>
<tr>
<th align=left colspan=7 height=23>�����ѯ&nbsp;��ע�⣺ <����> �� <����> ��Ĭ�ϰ��� <����>������������ʹ�ô����� ��</th>
</tr>
<tr>
<td width=20% class=forumrow>���</td>
<td width=80% class=forumrow colspan=6><input type=radio value=more name="visualprice1" checked>&nbsp;����&nbsp;&nbsp;<input type=radio value=less name="visualprice1">&nbsp;����&nbsp;&nbsp;<input size=5 name="visualprice1T" type=textbox> Ԫ</td>
</tr>
<tr>
<td width=20% class=forumrow>VIP��</td>
<td width=80% class=forumrow colspan=6><input type=radio value=more name="visualprice2" checked>&nbsp;����&nbsp;&nbsp;<input type=radio value=less name="visualprice2">&nbsp;����&nbsp;&nbsp;<input size=5 name="visualprice2T" type=textbox> Ԫ</td>
</tr>
<tr>
<td width=20% class=forumrow>��Ч��</td>
<td width=80% class=forumrow colspan=6><input type=radio value=more name="visualperiod" checked>&nbsp;����&nbsp;&nbsp;<input type=radio value=less name="visualperiod">&nbsp;����&nbsp;&nbsp;<input size=5 name="visualperiodT" type=textbox> ��</td>
</tr>
<tr>
<td width=20% class=forumrow>���</td>
<td width=80% class=forumrow colspan=6><input type=radio value=more name="visualquantity" checked>&nbsp;����&nbsp;&nbsp;<input type=radio value=less name="visualquantity">&nbsp;����&nbsp;&nbsp;<input size=5 name="visualquantityT" type=textbox> ��</td>
</tr>
<tr>
<td width=20% class=forumrow>����</td>
<td width=80% class=forumrow colspan=6><input type=radio value=more name="visualtotalsales" checked>&nbsp;����&nbsp;&nbsp;<input type=radio value=less name="visualtotalsales">&nbsp;����&nbsp;&nbsp;<input size=5 name="visualtotalsalesT" type=textbox> ��</td>
</tr>
<tr>
<td width=100% class=forumrow align=center colspan=7><input name="submit" type=submit value="   ��  ��   "></td>
</tr>
<input type=hidden value="9" name="visualSearch">
</form>
<%elseif request("action")="recalc" then
	set rs=server.createobject("ADODB.Recordset")
	sql="select itemid,count(itemid) from visual_myitems group by itemid"
	rs.open sql,conn,1,1
	do while not rs.eof
		conn.execute("update visual_items set totalsales="&rs(1)&" where id="&rs(0))
		rs.movenext
	loop
	set rs=nothing
	response.write "<tr><td colspan=7 class=forumrow>�����ɹ���</td></tr>"
elseif request("action")="visualSearch" then%>
<tr>
<th colspan=7 align=left height=23>�������</th>
</tr>
<%dim currentpage,page_count,Pcount
dim totalrec,endpage
currentPage=request("page")
if currentpage="" or not isInteger(currentpage) then
	currentpage=1
else
	currentpage=clng(currentpage)
	if err then
		currentpage=1
		err.clear
	end if
end if
Set rs= Server.CreateObject("ADODB.Recordset")
select case cint(request("visualSearch"))
case 1
	sqlstr=" flag=0"
case 2
	sqlstr=" flag=4"
case 3
	sqlstr=" flag=3"
case 4
	sqlstr=" flag=2"
case 5
	sqlstr=" vip"
case 6
	sqlstr=" datediff('d',dateandtime,now())=0"
case 7
	sqlstr=" quantity=0"
case 8
	sqlstr=" true"
case 9
	sqlstr=""
	if request("visualid")<>"" and isInteger(request("visualid")) then
		sqlstr=" id="&request("visualid")
	end if
	if request("visualname")<>"" then
		if request("visualnamechk")="yes" then
			if sqlstr="" then
				sqlstr=" name='"&request("visualname")&"'"
			else
				sqlstr=sqlstr & " and name='"&request("visualname")&"'"
			end if
		else
			if sqlstr="" then
				sqlstr=" name like '%"&request("visualname")&"%'"
			else
				sqlstr=sqlstr & " and name like '%"&request("visualname")&"%'"
			end if
		end if
	end if
	if cint(request("visualtype"))>0 then
		if request("visualtype")\100="0" then
			if sqlstr="" then
				sqlstr=" type mod 100="&(request("visualtype") mod 100)&""
			else
				sqlstr=sqlstr & " and type mod 100="&(request("visualtype") mod 100)&""
			end if
		else
			if sqlstr="" then
				sqlstr=" type="&request("visualtype")&""
			else
				sqlstr=sqlstr & " and type="&request("visualtype")&""
			end if
		end if
	end if
	if request("visualsex")<>"9" then
		if sqlstr="" then
			sqlstr=" sex="&request("visualsex")
		else
			sqlstr=sqlstr & " and sex="&request("visualsex")
		end if
	end if
	if request("visualflag")<>"9" then
		if sqlstr="" then
			sqlstr=" flag="&request("visualflag")
		else
			sqlstr=sqlstr & " and flag="&request("visualflag")
		end if
	end if
	if request("visualvip")="0" or request("visualvip")="1" then
		if sqlstr="" then
			if request("visualvip")="0" then
				sqlstr=" not vip"
			else
				sqlstr=" vip"
			end if
		else
			if request("visualvip")="0" then
				sqlstr=sqlstr & " and not vip"
			else
				sqlstr=sqlstr & " and vip"
			end if
		end if
	end if
	dim Tsqlstr
	if request("visualprice1T")<>"" then
		if request("visualprice1")="more" then
			Tsqlstr=" price1 >= "&request("visualprice1T")&""
		else
			Tsqlstr=" price1 <= "&request("visualprice1T")&""
		end if 	
		if sqlstr="" then 
			sqlstr=Tsqlstr
		else
			sqlstr=sqlstr & "and" & Tsqlstr
		end if 
	end if
	if request("visualprice2T")<>"" then
		if request("visualprice2")="more" then
			Tsqlstr=" price2 >= "&request("visualprice2T")&""
		else
			Tsqlstr=" price2 <= "&request("visualprice2T")&""
		end if 	
		if sqlstr="" then 
			sqlstr=Tsqlstr
		else
			sqlstr=sqlstr & " and " & Tsqlstr
		end if 
	end if
	if request("visualperiodT")<>"" then
		if request("visualperiod")="more" then
			Tsqlstr=" period >= "&request("visualperiodT")&""
		else
			Tsqlstr=" period <= "&request("visualperiodT")&""
		end if 	
		if sqlstr="" then 
			sqlstr=Tsqlstr
		else
			sqlstr=sqlstr & " and " & Tsqlstr
		end if 
	end if
	if request("visualquantityT")<>"" then
		if request("visualquantity")="more" then
			Tsqlstr=" quantity >= "&request("visualquantityT")&""
		else
			Tsqlstr=" quantity <= "&request("visualquantityT")&""
		end if 	
		if sqlstr="" then 
			sqlstr=Tsqlstr
		else
			sqlstr=sqlstr & " and " & Tsqlstr
		end if 
	end if
	if request("visualtotalsalesT")<>"" then
		if request("visualtotalsales")="more" then
			Tsqlstr=" totalsales >= "&request("visualtotalsalesT")&""
		else
			Tsqlstr=" totalsales <= "&request("visualtotalsalesT")&""
		end if 	
		if sqlstr="" then 
			sqlstr=Tsqlstr
		else
			sqlstr=sqlstr & " and " & Tsqlstr
		end if 
	end if
	if sqlstr="" then
		response.write "<tr><td colspan=7 class=forumrow>��ָ������������</td></tr>"
		response.end
	end if
case else
	response.write "<tr><td colspan=7 class=forumrow>����Ĳ�����</td></tr>"
	response.end
end select
sql="select id,name,sex,type,flag,vip,totalsales from visual_items where "&sqlstr&" and type<>0 order by dateandtime desc"
rs.open sql,conn,1,1
if rs.eof or rs.bof then
	response.write "<tr><td colspan=7 class=forumrow>û���ҵ���ؼ�¼��</td></tr>"
else%>
<FORM METHOD=POST ACTION="?action=tovisual">
<tr align=center>
<td class=forumHeaderBackgroundAlternate><B>��Ʒ���</B></td>
<td class=forumHeaderBackgroundAlternate><B>��Ʒ����</B></td>
<td class=forumHeaderBackgroundAlternate><B>��Ʒ���</B></td>
<td class=forumHeaderBackgroundAlternate><B>�����Ա�</B></td>
<td class=forumHeaderBackgroundAlternate><B>�޹���</B></td>
<td class=forumHeaderBackgroundAlternate><B>�Ƿ�VIPר��</B></td>
<td class=forumHeaderBackgroundAlternate><B>����</B></td>
</tr>
<%rs.PageSize = Cint(Forum_Setting(11))
rs.AbsolutePage=currentpage
page_count=0
totalrec=rs.recordcount
while (not rs.eof) and (not page_count = Cint(Forum_Setting(11)))%>
	<tr>
	<td class=forumrow width=2% align=center><%=rs("id")%></td>
	<td class=forumrow><a href="?action=modify&visualid=<%=rs("id")%>"><%=rs("name")%></a>&nbsp;(<%=rs("totalsales")%>)</td>
	<td class=forumrow width=20% align=center><%
		set rs2=Server.CreateObject("ADODB.Recordset")
		rs2.open "select top 1 typename from visual_type where seqno="&(rs("type") mod 100),conn,1,1
		if not rs2.eof and not rs2.bof then
			response.write rs2("typename")
		else
			response.write "δ֪"
		end if
		rs2.close
		response.write "��"
		rs2.open "select top 1 subtypename from visual_subtype where seqno="&(rs("type")\100)&" and typeseq="&(rs("type") mod 100),conn,1,1
		if not rs2.eof and not rs2.bof then
			response.write rs2("subtypename")
		else
			response.write "δ֪"
		end if
		rs2.close
		set rs2=nothing
	%></td>
	<td class=forumrow width=5% align=center><%
		select case rs("sex")
		case 0
			response.write "Ů"
		case 1
			response.write "��"
		case else
			response.write "����"
		end select
	%></td>
	<td class=forumrow width=10% align=center><%
		select case rs("flag")
		case 0
			response.write "������Ʒ"
		case 1 
			response.write "���л�Ա"
		case 2
			response.write "��������"
		case 3
			response.write "��������"
		case 4
			response.write "����Ա"
		end select
	%></td>
	<td class=forumrow width=10% align=center><%
		if rs("vip") then
			response.write "��"
		else
			response.write "��"
		end if
	%></td>
	<td class=forumrow align=center width=5%><input type="checkbox" name="visualid" value="<%=rs("id")%>"></td>
	</tr>
	<%page_count = page_count + 1
	rs.movenext
wend
Pcount=rs.PageCount
%>
<tr><td colspan=7 class=forumrow align=center>��ҳ��
<%call DispPageNum(currentpage,PCount,"""?page=","&visualSearch="&request("visualSearch")&"&visualid="&request("visualid")&"&visualname="&request("visualname")&"&visualtype="&request("visualtype")&"&visualsex="&request("visualsex")&"&visualflag="&request("visualflag")&"&visualvip="&request("visualvip")&"&action="&request("action")&"&visualprice1="&request("visualprice1")&"&visualprice1T="&request("visualprice1T")&"&visualprice2="&request("visualprice2")&"&visualprice2T="&request("visualprice2T")&"&visualperiod="&request("visualperiod")&"&visualperiodT="&request("visualperiodT")&"&visualquantity="&request("cisualquantity")&"&visualquantityT="&request("visualquantityT")&"&visualtotalsales="&request("visualtotalsales")&"&visualtotalsalesT="&request("visualtotalsalesT")&"&searchmax="&request("searchmax")&"""")%>
</td></tr>
<tr>
<td colspan=5 class=forumrow align=center><B>��ѡ������Ҫ���еĲ���</B></td>
<td colspan=2 class=forumrow align=center rowspan=9><input type=checkbox value="on" name="chkall" onclick="CheckAll(this.form)">&nbsp;��ǰҳȫ��<br><input type=checkbox value="on" name="chkpageall" onclick="CheckPageAll(this.form)">&nbsp;����ҳȫ��</td>
</tr>
<tr>
<td colspan=2 class=forumrow align=right>�޸����&nbsp;<input type="radio" name="visualaction" value=1 checked>&nbsp;</td>
<td colspan=3 class=forumrow align=left>&nbsp;<select size=1 name="selvisualtype"><%
set rs1=server.createobject("ADODB.recordset")
set rs2=server.createobject("ADODB.recordset")
rs1.open "select * from visual_type order by seqno",conn,1,1
do while not rs1.eof
	curseqno1=rs1("seqno")
	curtype1=rs1("typename")
	rs2.open "select * from visual_subtype where typeseq="&curseqno1&" order by seqno",conn,1,1
	rs1.movenext
	if not rs1.eof then
		response.write "<option value=00"&curseqno1&">��"&curtype1&"</option>"
	else
		response.write "<option value=00"&curseqno1&">��"&curtype1&"</option>"
	end if
	do while not rs2.eof
		curseqno2=rs2("seqno")
		curtype2=rs2("subtypename")
		rs2.movenext
		if not rs2.eof then
			if not rs1.eof then
				response.write "<option value="&curseqno2&"0"&curseqno1&">��&nbsp;&nbsp;��"&curtype2&"</option>"
			else
				response.write "<option value="&curseqno2&"0"&curseqno1&">��&nbsp;&nbsp;��"&curtype2&"</option>"
			end if
		else
			if not rs1.eof then
				response.write "<option value="&curseqno2&"0"&curseqno1&">��&nbsp;&nbsp;��"&curtype2&"</option>"
			else
				response.write "<option value="&curseqno2&"0"&curseqno1&">��&nbsp;&nbsp;��"&curtype2&"</option>"
			end if
		end if
	loop
	rs2.close
loop
rs1.close
set rs2=nothing
set rs1=nothing
%></select></td>
</tr>
<tr>
<td colspan=2 class=forumrow align=right>�޸��Ա�&nbsp;<input type="radio" name="visualaction" value=2>&nbsp;</td>
<td colspan=3 class=forumrow align=left>&nbsp;<select size=1 name="selvisualsex">
<option value=0>Ů</option>
<option value=1>��</option>
<option value=2>����</option>
</select></td>
</tr>
<tr>
<td colspan=2 class=forumrow align=right>�޸��޹���&nbsp;<input type="radio" name="visualaction" value=3>&nbsp;</td>
<td colspan=3 class=forumrow align=left>&nbsp;<select size=1 name="selvisualflag">
<option value=0>������Ʒ</option>
<option value=1>���л�Ա</option>
<option value=2>��������</option>
<option value=3>��������</option>
<option value=4>����Ա</option>
</select></td>
</tr>
<tr>
<td colspan=2 class=forumrow align=right>�޸�VIPר��&nbsp;<input type="radio" name="visualaction" value=4>&nbsp;</td>
<td colspan=3 class=forumrow align=left>&nbsp;<select size=1 name="selvisualvip">
<option value=0>��</option>
<option value=1>��</option>
</select></td>
</tr>
<tr>
<td colspan=2 class=forumrow align=right>�޸ı��&nbsp;<input type="radio" name="visualaction" value=5>&nbsp;</td>
<td colspan=3 class=forumrow align=left>&nbsp;<select size=1 name="selvisualprice1">
<option value=0>�޸�Ϊ(Ԫ)</option>
<option value=1>����(%)</option>
<option value=2>����(%)</option>
<option value=3>����(Ԫ)</option>
<option value=4>����(Ԫ)</option>
</select>&nbsp;<input type=textboxbox name="selvisualprice1T" size=15></td>
</tr>
<tr>
<td colspan=2 class=forumrow align=right>�޸�VIP��&nbsp;<input type="radio" name="visualaction" value=6>&nbsp;</td>
<td colspan=3 class=forumrow align=left>&nbsp;<select size=1 name="selvisualprice2">
<option value=0>�޸�Ϊ(Ԫ)</option>
<option value=1>����(%)</option>
<option value=2>����(%)</option>
<option value=3>����(Ԫ)</option>
<option value=4>����(Ԫ)</option>
</select>&nbsp;<input type=textboxbox name="selvisualprice2T" size=15></td>
</tr>
<tr>
<td colspan=2 class=forumrow align=right>�޸���Ч��&nbsp;<input type="radio" name="visualaction" value=7>&nbsp;</td>
<td colspan=3 class=forumrow align=left>&nbsp;<select size=1 name="selvisualperiod">
<option value=0>�޸�Ϊ(��)</option>
<option value=1>����(%)</option>
<option value=2>����(%)</option>
<option value=3>����(��)</option>
<option value=4>����(��)</option>
</select>&nbsp;<input type=textboxbox name="selvisualperiodT" size=15></td>
</tr>
<tr>
<td colspan=2 class=forumrow align=right>�޸Ŀ��&nbsp;<input type="radio" name="visualaction" value=8>&nbsp;</td>
<td colspan=3 class=forumrow align=left>&nbsp;<select size=1 name="selvisualquantity">
<option value=0>�޸�Ϊ(��)</option>
<option value=1>����(%)</option>
<option value=2>����(%)</option>
<option value=3>����(��)</option>
<option value=4>����(��)</option>
</select>&nbsp;<input type=textboxbox name="selvisualquantityT" size=15></td>
</tr>
<tr><td colspan=7 class=forumrow align=center><input type=hidden name=sqlstr value="<%=sqlstr%>">
<input type=submit name=submit value="ִ��ѡ���Ĳ���"  onclick="{if(confirm('ȷ��ִ��ѡ��Ĳ�����?')){return true;}return false;}">
</td></tr>
</FORM>
<%end if
rs.close
set rs=nothing	
elseif request("action")="tovisual" then
	response.write "<tr><th colspan=7 height=23 align=left>ִ�н��</th></tr>"
	if request("visualaction")="" then
		response.write "<tr><td colspan=7 class=forumrow>��ָ����ز�����</td></tr>"
		founderr=true
	end if
	if (request("visualid")="" and request("chkpageall")<>"on") or (request("sqlstr")="" and request("chkpageall")="on") then
		response.write "<tr><td colspan=7 class=forumrow>��ѡ�������Ʒ��</td></tr>"
		founderr=true
	end if
	if not founderr then
		if request("visualaction")=1 then
			if request("chkpageall")="on" then
				conn.execute("update visual_items set type="&replace(request("selvisualtype"),"'","")&" where "&request("sqlstr")&" and type<>0")
			else
				conn.execute("update visual_items set type="&replace(request("selvisualtype"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
			end if
			response.write "<tr><td colspan=7 class=forumrow>�����ɹ���</td></tr>"
		elseif request("visualaction")=2 then
			if request("chkpageall")="on" then
				conn.execute("update visual_items set sex="&replace(request("selvisualsex"),"'","")&" where "&request("sqlstr")&" and type<>0")
			else
				conn.execute("update visual_items set sex="&replace(request("selvisualsex"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
			end if
			response.write "<tr><td colspan=7 class=forumrow>�����ɹ���</td></tr>"
		elseif request("visualaction")=3 then
			if request("chkpageall")="on" then
				conn.execute("update visual_items set flag="&replace(request("selvisualflag"),"'","")&" where "&request("sqlstr")&" and type<>0")
			else
				conn.execute("update visual_items set flag="&replace(request("selvisualflag"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
			end if
			response.write "<tr><td colspan=7 class=forumrow>�����ɹ���</td></tr>"
		elseif request("visualaction")=4 then
			if request("selvisualvip")="1" then
				if request("chkpageall")="on" then
					conn.execute("update visual_items set vip=true where "&request("sqlstr")&" and type<>0")
				else
					conn.execute("update visual_items set vip=true where id in ("&replace(request("visualid"),"'","")&")")
				end if
			else
				if request("chkpageall")="on" then
					conn.execute("update visual_items set vip=false where "&request("sqlstr")&" and type<>0")
				else
					conn.execute("update visual_items set vip=false where id in ("&replace(request("visualid"),"'","")&")")
				end if
			end if
			response.write "<tr><td colspan=7 class=forumrow>�����ɹ���</td></tr>"
		elseif request("visualaction")=5 then
			if request("selvisualprice1T")="" or not isnumeric(request("selvisualprice1T")) then
				response.write "<tr><td colspan=7 class=forumrow>����ı�ۡ�</td></tr>"
			else
				if request("chkpageall")="on" then
					select case request("selvisualprice1")
					case 0
						conn.execute("update visual_items set price1="&replace(request("selvisualprice1T"),"'","")&" where "&request("sqlstr")&" and type<>0")
					case 1
						conn.execute("update visual_items set price1=price1*(1+"&replace(request("selvisualprice1T"),"'","")&"/100) where "&request("sqlstr")&" and type<>0")
					case 2
						if csng(request("selvisualprice1T"))<=100 then
							conn.execute("update visual_items set price1=price1*(1-"&replace(request("selvisualprice1T"),"'","")&"/100) where "&request("sqlstr")&" and type<>0")
						else
							conn.execute("update visual_items set price1=0 where "&request("sqlstr")&" and type<>0")
						end if
					case 3
						conn.execute("update visual_items set price1=price1+"&replace(request("selvisualprice1T"),"'","")&" where "&request("sqlstr")&" and type<>0")
					case 4
						conn.execute("update visual_items set price1=price1-"&replace(request("selvisualprice1T"),"'","")&" where "&request("sqlstr")&" and type<>0")
						conn.execute("update visual_items set price1=0 where "&request("sqlstr")&" and type<>0 and price1<0")
					end select
				else
					select case request("selvisualprice1")
					case 0
						conn.execute("update visual_items set price1="&replace(request("selvisualprice1T"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
					case 1
						conn.execute("update visual_items set price1=price1*(1+"&replace(request("selvisualprice1T"),"'","")&"/100) where id in ("&replace(request("visualid"),"'","")&")")
					case 2
						if csng(request("selvisualprice1T"))<=100 then
							conn.execute("update visual_items set price1=price1*(1-"&replace(request("selvisualprice1T"),"'","")&"/100) where id in ("&replace(request("visualid"),"'","")&")")
						else
							conn.execute("update visual_items set price1=0 where id in ("&replace(request("visualid"),"'","")&")")
						end if
					case 3
						conn.execute("update visual_items set price1=price1+"&replace(request("selvisualprice1T"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
					case 4
						conn.execute("update visual_items set price1=price1-"&replace(request("selvisualprice1T"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
						conn.execute("update visual_items set price1=0 where id in ("&replace(request("visualid"),"'","")&") and price1<0")
					end select
				end if
				response.write "<tr><td colspan=7 class=forumrow>�����ɹ���</td></tr>"
			end if
		elseif request("visualaction")=6 then
			if request("selvisualprice2T")="" or not isnumeric(request("selvisualprice2T")) then
				response.write "<tr><td colspan=7 class=forumrow>�����VIP�ۡ�</td></tr>"
			else
				if request("chkpageall")="on" then
					select case request("selvisualprice2")
					case 0
						conn.execute("update visual_items set price2="&replace(request("selvisualprice2T"),"'","")&" where "&request("sqlstr")&" and type<>0")
					case 1
						conn.execute("update visual_items set price2=price2*(1+"&replace(request("selvisualprice2T"),"'","")&"/100) where "&request("sqlstr")&" and type<>0")
					case 2
						if csng(request("selvisualprice2T"))<=100 then
							conn.execute("update visual_items set price2=price2*(1-"&replace(request("selvisualprice2T"),"'","")&"/100) where "&request("sqlstr")&" and type<>0")
						else
							conn.execute("update visual_items set price2=0 where "&request("sqlstr")&" and type<>0")
						end if
					case 3
						conn.execute("update visual_items set price2=price2+"&replace(request("selvisualprice2T"),"'","")&" where "&request("sqlstr")&" and type<>0")
					case 4
						conn.execute("update visual_items set price2=price2-"&replace(request("selvisualprice2T"),"'","")&" where "&request("sqlstr")&" and type<>0")
						conn.execute("update visual_items set price2=0 where "&request("sqlstr")&" and type<>0 and price2<0")
					end select
				else
					select case request("selvisualprice2")
					case 0
						conn.execute("update visual_items set price2="&replace(request("selvisualprice2T"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
					case 1
						conn.execute("update visual_items set price2=price2*(1+"&replace(request("selvisualprice2T"),"'","")&"/100) where id in ("&replace(request("visualid"),"'","")&")")
					case 2
						if csng(request("selvisualprice2T"))<=100 then
							conn.execute("update visual_items set price2=price2*(1-"&replace(request("selvisualprice2T"),"'","")&"/100) where id in ("&replace(request("visualid"),"'","")&")")
						else
							conn.execute("update visual_items set price2=0 where id in ("&replace(request("visualid"),"'","")&")")
						end if
					case 3
						conn.execute("update visual_items set price2=price2+"&replace(request("selvisualprice2T"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
					case 4
						conn.execute("update visual_items set price2=price2-"&replace(request("selvisualprice2T"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
						conn.execute("update visual_items set price2=0 where id in ("&replace(request("visualid"),"'","")&") and price2<0")
					end select
				end if
				response.write "<tr><td colspan=7 class=forumrow>�����ɹ���</td></tr>"
			end if
		elseif request("visualaction")=7 then
			if request("selvisualperiodT")="" or not isnumeric(request("selvisualperiodT")) then
				response.write "<tr><td colspan=7 class=forumrow>�������Ч�ڡ�</td></tr>"
			else
				if request("chkpageall")="on" then
					select case request("selvisualperiod")
					case 0
						conn.execute("update visual_items set period="&replace(request("selvisualperiodT"),"'","")&" where "&request("sqlstr")&" and type<>0")
					case 1
						conn.execute("update visual_items set period=period*(1+"&replace(request("selvisualperiodT"),"'","")&"/100) where "&request("sqlstr")&" and type<>0")
					case 2
						if csng(request("selvisualperiodT"))<=100 then
							conn.execute("update visual_items set period=period*(1-"&replace(request("selvisualperiodT"),"'","")&"/100) where "&request("sqlstr")&" and type<>0")
						else
							conn.execute("update visual_items set period=0 where "&request("sqlstr")&" and type<>0")
						end if
					case 3
						conn.execute("update visual_items set period=period+"&replace(request("selvisualperiodT"),"'","")&" where "&request("sqlstr")&" and type<>0")
					case 4
						conn.execute("update visual_items set period=period-"&replace(request("selvisualperiodT"),"'","")&" where "&request("sqlstr")&" and type<>0")
						conn.execute("update visual_items set period=0 where "&request("sqlstr")&" and type<>0 and period<0")
					end select
				else
					select case request("selvisualperiod")
					case 0
						conn.execute("update visual_items set period="&replace(request("selvisualperiodT"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
					case 1
						conn.execute("update visual_items set period=period*(1+"&replace(request("selvisualperiodT"),"'","")&"/100) where id in ("&replace(request("visualid"),"'","")&")")
					case 2
						if csng(request("selvisualperiodT"))<=100 then
							conn.execute("update visual_items set period=period*(1-"&replace(request("selvisualperiodT"),"'","")&"/100) where id in ("&replace(request("visualid"),"'","")&")")
						else
							conn.execute("update visual_items set period=0 where id in ("&replace(request("visualid"),"'","")&")")
						end if
					case 3
						conn.execute("update visual_items set period=period+"&replace(request("selvisualperiodT"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
					case 4
						conn.execute("update visual_items set period=period-"&replace(request("selvisualperiodT"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
						conn.execute("update visual_items set period=0 where id in ("&replace(request("visualid"),"'","")&") and period<0")
					end select
				end if
				response.write "<tr><td colspan=7 class=forumrow>�����ɹ���</td></tr>"
			end if
		elseif request("visualaction")=8 then
			if request("selvisualquantityT")="" or not isnumeric(request("selvisualquantityT")) then
				response.write "<tr><td colspan=7 class=forumrow>����Ŀ�档</td></tr>"
			else
				if request("chkpageall")="on" then
					select case request("selvisualquantity")
					case 0
						conn.execute("update visual_items set quantity="&replace(request("selvisualquantityT"),"'","")&" where "&request("sqlstr")&" and type<>0")
					case 1
						conn.execute("update visual_items set quantity=quantity*(1+"&replace(request("selvisualquantityT"),"'","")&"/100) where "&request("sqlstr")&" and type<>0")
					case 2
						if csng(request("selvisualquantityT"))<=100 then
							conn.execute("update visual_items set quantity=quantity*(1-"&replace(request("selvisualquantityT"),"'","")&"/100) where "&request("sqlstr")&" and type<>0")
						else
							conn.execute("update visual_items set quantity=0 where "&request("sqlstr")&" and type<>0")
						end if
					case 3
						conn.execute("update visual_items set quantity=quantity+"&replace(request("selvisualquantityT"),"'","")&" where "&request("sqlstr")&" and type<>0")
					case 4
						conn.execute("update visual_items set quantity=quantity-"&replace(request("selvisualquantityT"),"'","")&" where "&request("sqlstr")&" and type<>0")
						conn.execute("update visual_items set quantity=0 where "&request("sqlstr")&" and type<>0 and quantity<0")
					end select
				else
					select case request("selvisualquantity")
					case 0
						conn.execute("update visual_items set quantity="&replace(request("selvisualquantityT"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
					case 1
						conn.execute("update visual_items set quantity=quantity*(1+"&replace(request("selvisualquantityT"),"'","")&"/100) where id in ("&replace(request("visualid"),"'","")&")")
					case 2
						if csng(request("selvisualquantityT"))<=100 then
							conn.execute("update visual_items set quantity=quantity*(1-"&replace(request("selvisualquantityT"),"'","")&"/100) where id in ("&replace(request("visualid"),"'","")&")")
						else
							conn.execute("update visual_items set quantity=0 where id in ("&replace(request("visualid"),"'","")&")")
						end if
					case 3
						conn.execute("update visual_items set quantity=quantity+"&replace(request("selvisualquantityT"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
					case 4
						conn.execute("update visual_items set quantity=quantity-"&replace(request("selvisualquantityT"),"'","")&" where id in ("&replace(request("visualid"),"'","")&")")
						conn.execute("update visual_items set quantity=0 where id in ("&replace(request("visualid"),"'","")&") and quantity<0")
					end select
				end if
				response.write "<tr><td colspan=7 class=forumrow>�����ɹ���</td></tr>"
			end if
		else
			response.write "<tr><td colspan=7 class=forumrow>����Ĳ�����</td></tr>"
		end if
	end if
elseif request("action")="modify" then
	dim visualname
	response.write "<tr><th colspan=7 height=23 align=left>��Ʒ���ϲ���</th></tr>"
	if not isnumeric(request("visualid")) then
		response.write "<tr><td colspan=7 class=forumrow>�������Ʒ������</td></tr>"
		founderr=true
	end if
	if not founderr then
		Set rs= Server.CreateObject("ADODB.Recordset")
		sql="select * from visual_items where id="&request("visualid")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			response.write "<tr><td colspan=7 class=forumrow>û���ҵ������Ʒ��</td></tr>"
			founderr=true
		else%>
			<tr valign=middle>
				<td class=forumrow colspan=7>
					<div align="center">
      			<table width="400" border="0" cellspacing="1" cellpadding="0" class="tableborder">
      				<form action="?action=savevisualinfo&itemid=<%=rs("id")%>" method=post>
			        <tr valign=middle height=25> 
			          <td rowspan="5" height="84" width="100" class=forumrow align=center><img src="<%
			          	if localpic=1 then
			          		response.write PicPath&"0/"&right("0000"&rs("id"),4)&".GIF"
			          	else
			          		response.write PicPath&right("00000000"&rs("id"),8)&"/00/00/"
			          	end if
	          		%>" width="84" height="84" border="0"></td>
	          		<td class=forumrow>&nbsp; <input type=textboxbox name=itemname size=16 value=<%=rs("name")%>>(<select name=itemsex><%
          				if rs("sex")=1 then
          					response.write "<option value=1 selected>��</option>"
          				else
          					response.write "<option value=1>��</option>"
          				end if
          				if rs("sex")=0 then
          					response.write "<option value=0 selected>Ů</option>"
          				else
          					response.write "<option value=0>Ů</option>"
          				end if
									if rs("sex")=2 then
          					response.write "<option value=2 selected>����</option>"
          				else
          					response.write "<option value=2>����</option>"
          				end if
          			%></select>)&nbsp;</td>
	        		</tr>
	        		<tr valign=middle height=25> 
	          		<td class=forumrow>&nbsp;&nbsp;�ꡡ�ۣ�<input type=textbox name=itemprice1 size=15 value=<%=rs("price1")%> Ԫ</td>
	        		</tr>
	        		<tr valign=middle height=25> 
	          		<td class=forumrow>&nbsp;&nbsp;VIP �ۣ�<input type=textbox name=itemprice2 size=15 value=<%=rs("price2")%> Ԫ</td>
	        		</tr>
	        		<tr valign=middle height=25> 
	          		<td class=forumrow>&nbsp;&nbsp;��Ч�ڣ�<input type=textbox name=itemperiod size=15 value=<%=rs("period")%> ��</td>
	        		</tr>
	        		<tr valign=middle height=25> 
	          		<td class=forumrow>&nbsp;&nbsp;�⡡�棺<input type=textbox name=itemquantity size=15 value=<%=rs("quantity")%> ��</td>
	        		</tr>
	        		<tr valign=middle height=25> 
	          		<td class=forumrow align=center><%
		          		if rs("vip") then
		          			response.write "<input type=checkbox name=itemvip value=yes checked> VIP��Ʒ"
		          		else
		          			response.write "<input type=checkbox name=itemvip value=yes> VIP��Ʒ"
		          		end if
	          		%></td>
	          		<td class=forumrow>&nbsp;&nbsp;�޹��ߣ�<select name=itemflag><%
		          		if rs("flag")=0 then
		          			response.write "<option value=0 selected>������Ʒ</option>"
		          		else
		          			response.write "<option value=0>������Ʒ</option>"
		          		end if
		          		if rs("flag")=1 then
		          			response.write "<option value=1 selected>���л�Ա</option>"
		          		else
		          			response.write "<option value=1>���л�Ա</option>"
		          		end if
									if rs("flag")=2 then
		          			response.write "<option value=2 selected>��������</option>"
		          		else
		          			response.write "<option value=2>��������</option>"
		          		end if
									if rs("flag")=3 then
		          			response.write "<option value=3 selected>��������</option>"
		          		else
		          			response.write "<option value=3>��������</option>"
		          		end if
									if rs("flag")=4 then
		          			response.write "<option value=4 selected>����Ա</option>"
		          		else
		          			response.write "<option value=4>����Ա</option>"
		          		end if
		          	%></select></td>
		          </tr>
		          <tr valign=middle height=25>
		          	<td colspan=2 class=forumrow align=center><input type=Submit value="������" name="Submit"></td>
		          </tr>
	      			</form>
	      		</table>  
					</div>
				</td>
			</tr>
		<%end if
		rs.close
		set rs=nothing
	end if
elseif request("action")="savevisualinfo" then
	response.write "<tr><th colspan=7 height=23 align=left>������Ʒ����</th></tr>"
	if not isnumeric(request("itemid")) then
		response.write "<tr><td colspan=7 class=forumrow>�������Ʒ������</td></tr>"
		founderr=true
	end if
	if not founderr then
		Set rs= Server.CreateObject("ADODB.Recordset")
		sql="select * from visual_items where id="&request("itemid")
		rs.open sql,conn,1,3
		if rs.eof and rs.bof then
			response.write "<tr><td colspan=7 class=forumrow>û���ҵ������Ʒ��</td></tr>"
			founderr=true
		else
			rs("name")=request.form("itemname")
			rs("sex")=request.form("itemsex")
			rs("price1")=request.form("itemprice1")
			rs("price2")=request.form("itemprice2")
			rs("period")=request.form("itemperiod")
			rs("quantity")=request.form("itemquantity")
			if request.form("itemvip")="yes" then
				rs("vip")=true
			else
				rs("vip")=false
			end if
			rs("flag")=request.form("itemflag")
			rs.update
		end if
		rs.close
		set rs=nothing
	end if
	if founderr then
		response.write "<tr><td colspan=7 class=forumrow>����ʧ�ܡ�</td></tr>"
	else
		response.write "<tr><td colspan=7 class=forumrow>������Ʒ���ݳɹ���</td></tr>"
	end if
end if%>
</table>
<%end sub%>                      
<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name == 'visualid') e.checked = form.chkall.checked; 
   }
  }

function CheckPageAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if ((e.name.substring(0,3) == 'chk' || e.name == 'visualid') && e.name != 'chkpageall') e.disabled = form.chkpageall.checked; 
   }
  }
//-->
</script>
</html>