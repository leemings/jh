<!--#include file=conn.asp-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_visual_const.asp" -->
<%
dim CurItem,ItemName,rsvisual
dim Salefee,SalePrice,CurDate,SendMsg,SucFlag
dim MinPrice,CurType

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

set rs=server.createobject("adodb.recordset")
sql="select top 1 * from visual_items where id="&CurItem
rs.open sql,conn,1,1
ItemName=rs("name")
CurType=rs("type") mod 100
if CurType>=5 then CurType=4
salefee=conn.execute("select top 1 sale from visual_config")(0)
minprice=int(rs("price1")*conn.execute("select top 1 salemin from visual_config")(0)/100+0.99)
rs.close
if request("action")="sale" then
	if isnull(Request("SalePrice")) or Request("SalePrice")="" then
		SendMsg=SendMsg+"<br>"+"<li>���ۼ۸񲻿�Ϊ�գ�"
		founderr=true
	else
		SalePrice=CheckStr(Request("SalePrice"))
		if not isInteger(SalePrice) then
			SendMsg=SendMsg+"<br>"+"<li>���ۼ۸������������"
			founderr=true
		end if
		if clng(SalePrice)<=MinPrice then
			SendMsg=SendMsg+"<br>"+"<li>���ۼ۸�������"&MinPrice&"��"
			founderr=true
		end if
		SalePrice=clng(SalePrice)
	end if
end if
if founderr then
	call sale_head()
	call sale_Msg()
else
	CurDate=year(now())&"-"&month(now())&"-"&day(now())
	set rs=server.createobject("adodb.recordset")
	sql="select * from visual_myitems where itemid="&CurItem&" and username='"&membername&"'"
	rs.open sql,conn,1,1
	if rs.bof or rs.eof then
		if request("StopSale")=1 then
			SendMsg=SendMsg+"<br>"+"<li>��û�������Ʒ�����Ʒ�ѱ��˹��򣬲���ͣ�ۣ�"
		else
			SendMsg=SendMsg+"<br>"+"<li>��û�������Ʒ�����ܳ��ۣ�"
		end if
		founderr=true
	else
		if request("StopSale")=1 then
			if rs("type")<>5 then
				SendMsg=SendMsg+"<br>"+"<li>����δ���������Ʒ������ͣ�ۣ�"
				founderr=true
			end if
		else
			if rs("type")=5 then
				SendMsg=SendMsg+"<br>"+"<li>���Ѿ������������Ʒ�������ٴγ��ۣ�"
				founderr=true
			end if
		end if
	end if
	rs.close
	set rs=nothing
	if request("action")="sale" then
		if int(saleprice*salefee/100+0.99)>mymoney then
			SendMsg=SendMsg+"<br>"+"<li>�����ֽ�("&mymoney&" Ԫ)���㣬�޷�֧�������г��������� "&int(saleprice*salefee/100+0.99)&" Ԫ������ʧ�ܣ�"
			founderr=true
		end if
	end if
	if founderr then
		call sale_head()
		call sale_Msg()
	else
		if request("stopsale")=1 then
			set rs=server.createobject("adodb.recordset")
			rs.open "select top 1 * from visual_myitems where username='"&membername&"' and itemid="&curItem,conn,1,3
			salefee=int(rs("salePrice")*salefee/100+0.99)
			rs("salePrice")=0
			rs("type")=CurType
			rs.update
			rs.close
			set rs=nothing
			conn.execute("delete from visual_events where ItemID="&CurItem&" and (UserName='' or isnull(username)) and FromUser='"&membername&"' and Type=2")
			conn.execute("update [user] set userwealth=userwealth+"&salefee&" where username='"&membername&"'")
			SendMsg="<li>�Ѿ�ȡ���ˡ�<font color=blue>"&itemName&"</font>���ĳ��ۣ�"
			call sale_head()
			response.write "<script>"
			response.write "window.opener.document.execCommand('Refresh');"
			response.write "</script>"
			call sale_Msg()
		else
			if request("action")="sale" then
				set rs=server.createobject("adodb.recordset")
				rs.open "select top 1 * from visual_myitems where username='"&membername&"' and itemid="&curItem,conn,1,3
				rs("salePrice")=saleprice
				rs("type")=5
				rs.update
				rs.close
				set rs=nothing
				salefee=int(saleprice*salefee/100+0.99)
				conn.execute("insert into visual_events (ItemID,UserName,FromUser,DateAndTime,Price,Type) values ("&CurItem&",'','"&membername&"','"&curdate&"',"&saleprice&",2)")
				conn.execute("update [user] set userwealth=userwealth-"&salefee&" where username='"&membername&"'")
				SendMsg="<li>�Ѿ��ɹ��س����ˡ�<font color=blue>"&itemName&"</font>�����ȴ����˵Ĺ���"
				call sale_head()
				response.write "<script>"
				response.write "window.opener.document.execCommand('Refresh');"
				response.write "</script>"
				call sale_Msg()
			else
				call sale_head()%>
				<form action="?" method=post name=fix>
				<input type=hidden name="action" value="sale">
				<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
					<tr> 
						<th colspan=3>����������Ʒ -- <%=ItemName%></th>
					</tr>
					<tr> 
						<td class=tablebody1 valign=middle><b>���ۼ۸�</b></td>
						<td class=tablebody1 valign=middle><input type=text name="saleprice" size=10> Ԫ</td>
	         	<td rowspan=2 class=tablebody1 valign=middle baseName=images/img_visual/blank.gif width="84" height="84" border="0" itemgender="" itemno=<%=CurItem%> layerno="" localpic=<%=LocalPic%> ImagePath="<%=PicPath%>" style="behavior:url(inc/z_show2.htc)"></td>
					</tr>
					<tr> 
						<td class=tablebody1 colspan=2><b>˵��</b>��<br>�� ������Ʒ����ͼ۸�Ϊ <%=minPrice%> Ԫ<br><font color=<%=forum_body(8)%>>�� ������Ʒÿ����ȡ�ɽ��۸�� <%=SaleFee%> %</font></td>
					</tr>
					<tr> 
						<td class=tablebody2 valign=middle colspan=3 align=center><input type=hidden name=itemid value=<%=CurItem%>><input type=Submit value="ȷ��" name=Submit>&nbsp;<input type="button" name="close" value="�ر�" onclick="window.close()"></td>
					</tr>
				</table>
				</form>
			<%end if
		end if
	end if
end if
call footer()

sub sale_head()%>
	<html><head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=Forum_info(0)%>--����������Ʒ</title>
	<!--#include file=inc/forum_css.asp-->
	</head>
	<body <%=Forum_body(11)%>>
<%end sub

sub sale_msg()%>
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
			<td width="100%" class=tablebody2><a href="javascript:window.close()">[�رմ���]</a><%if request("action")="sale" and founderr then%>&nbsp;&nbsp;<a href="javascript:history.go(-1)">[������ҳ]</a><%end if%></td>
		</tr>
	</table>
<%end sub%>