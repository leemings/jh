<!--#include file=conn.asp-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_visual_const.asp" -->
<%
dim CurItems,ItemName,rsvisual
dim SucFlag,CurItem,CartBag,SendMsg,CurType,SalePrice

founderr=false
if not founduser or membername=""  then
	sendmsg=sendmsg+"<br>"+"<li>����û��<a href=login.asp>��¼��̳</a>������ʹ�ø���������ƹ��ܡ��������û��<a href=reg.asp>ע��</a>������<a href=reg.asp>ע��</a>��"
	founderr=true
end if
CurItem=CheckStr(request("itemid"))
if isnull(curitem) then
	sendmsg=sendmsg+"<br>"+"<li>û��ָ����Ʒ��ţ�"
	founderr=true
else
	if not isnumeric(curitem) then
		sendmsg=sendmsg+"<br>"+"<li>�ύ�������д���"
		founderr=true
	end if
end if

if founderr then
	call send_head()
	call send_msg()
else
	set rsvisual=server.createobject("ADODB.Recordset")
	sql="select top 1 itemid from visual_myitems where username='"&membername&"' and itemid="&CurItem
	rsvisual.open sql,conn,1,1
	if not (rsvisual.eof or rsvisual.bof) then
		sendmsg=sendmsg+"<br>"+"<li>�Ѿ�ӵ�и���Ʒ�������ٴι���"
		founderr=true
	else
		if request("fromuser")<>"" then
			rsvisual.close
			sql="select m.username,m.saleprice,v.name,v.type from visual_myitems m inner join visual_items v on m.itemid=v.id where m.itemid="&CurItem&" and m.username='"&checkstr(trim(request("fromuser")))&"'"
			rsvisual.open sql,conn,1,1
			ItemName=rsvisual("name")
			CurType=rsvisual("type") mod 100
			if CurType>=5 then CurType=4
			saleprice=rsvisual("salePrice")
			if saleprice>mymoney then
				sendmsg=sendmsg+"<br>"+"<li>�������Ʒ��Ҫ�ֽ� "&saleprice&" Ԫ�������ֽ�("&mymoney&" Ԫ)���㣬�޷�����"
				founderr=true
			end if
		else
			rsvisual.close
			sql="select * from visual_items where id="&CurItem
			rsvisual.open sql,conn,1,1
			ItemName=rsvisual("name")
			if rsvisual("vip") and v_myvip<>1 then
				sendmsg=sendmsg+"<br>"+"<li>������VIP�����ܹ���VIPר����Ʒ��"
				founderr=true
			end if
			if rsvisual("quantity")<=0 then
				sendmsg=sendmsg+"<br>"+"<li>����Ʒ��治�㣬���޷�����"
				founderr=true
			end if
			if rsvisual("flag")=0 then
				sendmsg=sendmsg+"<br>"+"<li>������Ʒ�����޷�����"
				founderr=true
			end if
			if rsvisual("flag")=1 and membername="" then
				sendmsg=sendmsg+"<br>"+"<li>ֻ�л�Ա���ܹ������Ʒ�����޷�����"
				founderr=true
			end if
			if rsvisual("flag")=2 and not master and not boardmaster and not superboardmaster then
				sendmsg=sendmsg+"<br>"+"<li>ֻ�а���������Ա���ܹ������Ʒ�����޷�����"
				founderr=true
			end if
			if rsvisual("flag")=3 and not master and not superboardmaster then
				sendmsg=sendmsg+"<br>"+"<li>ֻ�г��������͹���Ա���ܹ������Ʒ�����޷�����"
				founderr=true
			end if
			if rsvisual("flag")=4 and not master then
				sendmsg=sendmsg+"<br>"+"<li>ֻ�й���Ա���ܹ������Ʒ�����޷�����"
				founderr=true
			end if
		end if
		if founderr then
			call send_head()
			call send_msg()
		else	
			if request("fromuser")<>"" then
				conn.execute("update visual_myitems set type="&curtype&",price=saleprice,saleprice=0,fromuser=username,username='"&membername&"' where itemid="&curitem&" and type=5 and username='"&checkstr(trim(request("fromuser")))&"'")
				conn.execute("update visual_events set username='"&membername&"' where itemid="&curitem&" and type=2 and fromuser='"&checkstr(trim(request("fromuser")))&"' and (username='' or isnull(username))")
				conn.execute("insert into visual_events (ItemID,UserName,FromUser,DateAndTime,Price,Type) values ("&CurItem&",'"&membername&"','"&checkstr(trim(request("fromuser")))&"','"&year(now())&"-"&month(now())&"-"&day(now())&"',"&saleprice&",0)")
				conn.execute("update [user] set userwealth=userwealth-"&saleprice&" where username='"&membername&"'")
				conn.execute("update [user] set userwealth=userwealth+"&saleprice&" where username='"&checkstr(trim(request("fromuser")))&"'")
				set rs=server.createobject("adodb.recordset")
				sql="select * from [message]"      
				rs.open sql,conn,1,3      
				rs.addnew      
				rs("sender")=forum_info(0)
				rs("incept")=trim(request("fromuser"))
				rs("title")="�����۵���Ʒ���˹����ˣ�"
				rs("content")="���ڶ����г����۵���Ʒ��[color=blue]"&itemName&"[/color]���Ѿ���"&membername&"�����ˣ�"
				rs("flag")=0 
				rs("issend")=1
				rs.update
				rs.close
				set rs=nothing
				sendmsg="<li>�Ѿ��ɹ��ع����˶�����Ʒ <font color=blue>"&ItemName&"</font>��"
				call send_head()
				response.write "<script>"
				response.write "window.opener.document.execCommand('Refresh');"
				response.write "</script>"
				call send_msg()
			else
				CartBag=request.cookies("mybuy_"&userid)
				if isnull(CartBag) then CartBag=""
				if instr(1,"|"&CartBag&"|","|"&CurItem&"|")<=0 then
					if ubound(split(CartBag,"|"))<CartLength-1 then
						if CartBag<>"" then CartBag=CartBag&"|"
						CartBag=CartBag&CurItem
						response.cookies("mybuy_"&userid)=CartBag
						sendmsg="<li>�ѽ���Ʒ <font color=blue>"&ItemName&"</font> �ŵ����Ĺ�����У�"
					else
						sendmsg=sendmsg+"<br>"+"<li>���Ĺ�����������޷�������Ʒ���룡"
						founderr=true
					end if
				else
					sendmsg=sendmsg+"<br>"+"<li>����������д���Ʒ�������ٴι���"
					founderr=true
				end if
				call send_head()
				call send_msg()
			end if
		end if
	end if
	rsvisual.close
	set rsvisual=nothing
end if
call footer()

sub send_head()%>
	<html><head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=Forum_info(0)%>--����������Ʒ</title>
	<!--#include file=inc/forum_css.asp-->
	</head>
	<body <%=Forum_body(11)%>">
<%end sub

sub send_msg()%>
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
			<td width="100%" class=tablebody2><a href="javascript:window.close()">[�رմ���]</a></td>
		</tr>
	</table>
<%end sub%>