<!--#include file=conn.asp-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_visual_const.asp" -->
<%
dim CurItem,ItemName,rsvisual
dim FixType,FixDays,CurType
dim fixfee,CurDate,SendMsg,SucFlag,FixPeriod
Dim FixFee0,FixFee1,New0,New1

founderr=false
if not founduser or membername=""  then
	SendMsg=SendMsg+"<br>"+"<li>����û��<a href=login.asp>��¼��̳</a>������ʹ�ø���������ƹ��ܡ��������û��<a href=reg.asp>ע��</a>������<a href=reg.asp>ע��</a>��"
	founderr=true
end if
if isnull(request("itemid")) or request("itemid")="" then
	SendMsg=SendMsg+"<br>"+"<li>û��ָ����Ʒ��ţ�"
	founderr=true
else
	CurItem=CheckStr(request("itemid"))
	if not isnumeric(curitem) then
		SendMsg=SendMsg+"<br>"+"<li>�ύ�������д���"
		founderr=true
	end if
end if

if request("action")="fix" then
	if isnull(Request("fixtype")) or Request("fixtype")="" then
		SendMsg=SendMsg+"<br>"+"<li>���·�ʽ����Ϊ�գ�"
		founderr=true
	else
		FixType=CheckStr(Request("fixtype"))
		if not isInteger(FixType) then
			SendMsg=SendMsg+"<br>"+"<li>���·�ʽ������������"
			founderr=true
		end if
		if cint(FixType)<>0 then FixType=1 else FixType=0
	end if
	if FixType=1 then
		if isnull(Request("fixdays")) or Request("fixdays")="" then
			SendMsg=SendMsg+"<br>"+"<li>������������Ϊ�գ�"
			founderr=true
		else
			FixDays=CheckStr(Request("fixdays"))
			if not isInteger(FixDays) then
				SendMsg=SendMsg+"<br>"+"<li>��������������������"
				founderr=true
			end if
			if clng(FixDays)<=0 or clng(FixDays)>100 then
				SendMsg=SendMsg+"<br>"+"<li>���������������0��С��100��"
				founderr=true
			end if
			FixDays=clng(FixDays)
		end if
	end if
end if
if founderr then
	call fix_head()
	call fix_Msg()
else
	set rs=server.createobject("adodb.recordset")
	sql="select top 1 * from visual_items where id="&CurItem
	rs.open sql,conn,1,1
	ItemName=rs("name")
	CurType=rs("type") mod 100
	if CurType>=5 then CurType=4
	set rsvisual=conn.execute("select top 1 new0,new1,period from visual_config")
	new0=rsvisual(0)
	new1=rsvisual(1)
	FixPeriod=rsvisual(2)
	rsvisual.close
	if rs("period")>0 then FixPeriod=rs("Period")
	if FixPeriod<=0 then FixPeriod=30
	if v_myvip<>1 then
		if request("action")="fix" then
			if FixType=0 then
				fixfee=rs("price1")*new0
			else
				fixfee=int(rs("price1")*new1*FixDays/100/FixPeriod+0.99)
			end if
		end if
		fixfee1=rs("price1")/FixPeriod*new1/100
		fixfee0=rs("price1")*new0
	else
		if request("action")="fix" then
			if FixType=0 then
				fixfee=rs("price2")*new0
			else
				fixfee=int(rs("price2")*new1*FixDays/100/FixPeriod+0.99)
			end if
		end if
		fixfee1=rs("price2")/FixPeriod*new1/100
		fixfee0=rs("price2")*new0
	end if
	rs.close
	CurDate=year(now())&"-"&month(now())&"-"&day(now())
	set rs=server.createobject("adodb.recordset")
	sql="select * from visual_myitems where itemid="&CurItem&" and username='"&membername&"'"
	rs.open sql,conn,1,1
	if rs.bof or rs.eof then
		SendMsg=SendMsg+"<br>"+"<li>��û�������Ʒ�����ܷ��£�"
		founderr=true
	elseif rs("period")=0 then
		SendMsg=SendMsg+"<br>"+"<li>���������Ʒ�Ѿ�����Զ��Ч�ˣ������ٴη��£�"
		founderr=true
	end if
	rs.close
	set rs=nothing
	if founderr then
		call fix_head()
		call fix_Msg()
	else
		if request("action")="fix" then
			if mymoney<fixfee then
				sendmsg=sendmsg+"<br>"+"<li>���´���Ʒ��Ҫ�ֽ� "&fixfee&" Ԫ�������ֽ�("&mymoney&" Ԫ)���㣬�޷����£�"
				founderr=true
				call fix_head()
			else
				dim myvisualsplit,TempSplit,CurPeriod
				myvisualsplit=split(v_myvisual,"|")
				set rs=server.createobject("adodb.recordset")
				rs.open "select top 1 * from visual_myitems where username='"&membername&"' and itemid="&curItem,conn,1,3
				if fixtype=0 then
					rs("period")=0
				else
					if rs("period")-datediff("d",rs("fixdate"),now())>0 then
						rs("period")=rs("period")-datediff("d",rs("fixdate"),now())+fixdays
					else
						rs("period")=fixdays
					end if
				end if
				rs("fixDate")=cdate(curdate)
				if rs("type")=0 then
					rs("type")=CurType
				end if
				rs.update
				rs.close
				set rs=nothing
	  		conn.execute("update [user] set userwealth=userwealth-"&FixFee&" where username='"&membername&"'")
				SendMsg="<li>�Ѿ��ɹ��ط����ˡ�<font color=blue>"&itemName&"</font>���������� <font color="&forum_body(8)&">"&FixFee&"</font> Ԫ��"
				call fix_head()
				response.write "<script>"
				response.write "window.opener.document.execCommand('Refresh');"
				response.write "</script>"
			end if
			call fix_Msg()
		else
			call fix_head()%>
			<form action="?" method=post name=fix>
			<input type=hidden name="action" value="fix">
			<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
				<tr> 
					<th colspan=3>����������Ʒ -- <%=ItemName%></th>
				</tr>
				<tr> 
					<td class=tablebody1 valign=middle><b>���·�ʽ��</b></td>
					<td class=tablebody1 valign=middle><input type=radio name="fixtype" value="1" checked>&nbsp;������Ч�� <input type=text name="fixdays" value=10 size=10> ��</td>
         	<td rowspan=3 class=tablebody1 valign=middle baseName=images/img_visual/blank.gif width="84" height="84" border="0" itemgender="" itemno=<%=CurItem%> layerno="" localpic=<%=LocalPic%> ImagePath="<%=PicPath%>" style="behavior:url(inc/z_show2.htc)"></td>
				</tr>
				<tr> 
					<td class=tablebody1 valign=middle>&nbsp;</td>
					<td class=tablebody1 valign=middle><input type=radio name="fixtype" value="0">&nbsp;����Ϊ��Զ��Ч</td>
				</tr>
				<tr> 
					<td class=tablebody1 colspan=2><b>˵��</b>��<br><font color=<%=forum_body(8)%>>�� ����һ����Ч�ڵķ���Ϊ��<%=formatnumber(fixfee1,2,-1)%> Ԫ<br>�� ���¸���ƷΪ��Զ��Ч�ķ���Ϊ��<%=fixfee0%> Ԫ</font><BR>�� ����ӵ�е��ֽ�Ϊ��<%=mymoney%> Ԫ</td>
				</tr>
				<tr> 
					<td class=tablebody2 valign=middle colspan=3 align=center><input type=hidden name=itemid value=<%=CurItem%>><input type=Submit value="ȷ��" name=Submit>&nbsp;<input type="button" name="close" value="�ر�" onclick="window.close()"></td>
				</tr>
			</table>
			</form>
		<%end if
	end if
end if
call footer()

sub fix_head()%>
	<html><head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=Forum_info(0)%>--����������Ʒ</title>
	<!--#include file=inc/forum_css.asp-->
	</head>
	<body <%=Forum_body(11)%>>
<%end sub

sub fix_msg()%>
	<br>
	<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
		<tr align=center>
			<th width="100%">����������Ʒ��Ϣ</th>
		</tr>
		<tr>
			<td width="100%" class=tablebody1><b>���������</b><br><%
				if founderr then
					response.write "<font color="&forum_body(8)&">"
					response.write sendmsg
					response.write "</font>"
				else
					response.write sendmsg
				end if
			%></td>
		</tr>
		<tr align=center>
			<td width="100%" class=tablebody2><a href="javascript:window.close()">[�رմ���]</a><%if request("action")="fix" and founderr then%>&nbsp;&nbsp;<a href="javascript:history.go(-1)">[������ҳ]</a><%end if%></td>
		</tr>
	</table>
<%end sub%>