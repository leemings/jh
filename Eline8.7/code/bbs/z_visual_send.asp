<!--#include file=conn.asp-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_visual_const.asp" -->
<%
dim CurItem,ItemName,rsvisual,UserName,MsgBody
dim SucFlag,CartBag,SendMsg
dim SendSplit,CurType
dim itemsplit,itemcount
dim sendfee,CurDate,CurPrice,SalePrice

SucFlag=False
if request("myself")=1 then
	itemcount=0
else
	if request("fromuser")<>"" then
		itemcount=0
	else
		CartBag=request.cookies("mybuy_"&userid)
		if isnull(CartBag) then CartBag=""
		if CartBag="" then
			itemcount=0
		else
			itemsplit=split(CartBag,"|")
			itemcount=ubound(itemsplit)+1
		end if
	end if
end if
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

if request("action")="send" then
	if isnull(Request("username")) or Request("username")="" then
		SendMsg=SendMsg+"<br>"+"<li>�����߲���Ϊ�գ�"
		founderr=true
	else
		UserName=CheckStr(Request("username"))
		if lcase(trim(username))=lcase(trim(membername)) then
			SendMsg=SendMsg+"<br>"+"<li>���ܸ��Լ�������Ʒ��"
			founderr=true
		end if
		if conn.execute("select top 1 userid from [user] where username='"&username&"'").eof then
			SendMsg=SendMsg+"<br>"+"<li>�����߲����ڣ�"
			founderr=true
		end if
	end if
	if isnull(Request("msgbody")) or Request("msgbody")="" then
		SendMsg=SendMsg+"<br>"+"<li>��ͬ���Բ���Ϊ�գ�"
		founderr=true
	else
		msgbody=CheckStr(Request("msgbody"))
	end if
elseif itemcount>0 then
	if isnull(Request("username")) or Request("username")="" then
		UserName=""
	else
		UserName=CheckStr(Request("username"))
	end if
	MsgBody=""
	for i=0 to itemcount-1
		SendSplit=split(itemsplit(i),"$")
		if ubound(SendSplit)=2 then
			if SendSplit(0)=CurItem and lcase(trim(SendSplit(1)))=lcase(trim(UserName)) then
				MsgBody=SendSplit(2)
				SucFlag=true
				exit for
			end if
		end if
	next
end if
if founderr then
	call send_head()
	call Send_Msg()
else
	set rsvisual=server.createobject("ADODB.Recordset")
	if request("action")="send" then
		sql="select top 1 itemid from visual_myitems where username='"&username&"' and itemid="&CurItem
		rsvisual.open sql,conn,1,1
		if not (rsvisual.eof or rsvisual.bof) then
			SendMsg=SendMsg+"<br>"+"<li>�û� <font color=blue>"&username&"</font> �Ѿ�ӵ�и���Ʒ���������ͣ�"
			founderr=true
		end if
		rsvisual.close
	end if
	sql="select * from visual_items where id="&CurItem
	rsvisual.open sql,conn,1,1
	ItemName=rsvisual("name")
	CurType=rsvisual("type") mod 100
	if CurType>=5 then CurType=4
	sendfee=int(rsvisual("price1")*conn.execute("select top 1 send from visual_config")(0)/100+0.99)
	if v_myvip<>1 then
		CurPrice=rsvisual("price1")
	else
		CurPrice=rsvisual("price2")
	end if
	if request("myself")=1 then
		set rs=server.createobject("adodb.recordset")
		sql="select * from visual_myitems where itemid="&CurItem&" and username='"&membername&"'"
		rs.open sql,conn,1,1
		if rs.eof then
			SendMsg=SendMsg+"<br>"+"<li>��û�������Ʒ���������ͣ�"
			founderr=true
		end if
		rs.close
		set rs=nothing
	else
		if request("fromuser")<>"" then
			set rs=server.createobject("adodb.recordset")
			sql="select saleprice from visual_myitems where itemid="&CurItem&" and username='"&checkstr(trim(request("fromuser")))&"'"
			rs.open sql,conn,1,1
			saleprice=rs("salePrice")
			if saleprice+sendfee>mymoney then
				sendmsg=sendmsg+"<br>"+"<li>�������ʹ���Ʒ��Ҫ�ֽ� "&(saleprice+sendfee)&" Ԫ�������ֽ�("&mymoney&" Ԫ)���㣬�޷����ͣ�"
				founderr=true
			end if
			rs.close
			set rs=nothing
		else
			if rsvisual("vip") and v_myvip<>1 then
				SendMsg=SendMsg+"<br>"+"<li>������VIP�����ܹ���VIPר����Ʒ��"
				founderr=true
			end if
			if rsvisual("quantity")<=0 then
				SendMsg=SendMsg+"<br>"+"<li>����Ʒ��治�㣬���޷�����"
				founderr=true
			end if
			if rsvisual("flag")=0 then
				SendMsg=SendMsg+"<br>"+"<li>������Ʒ�����޷�����"
				founderr=true
			end if
			if rsvisual("flag")=1 and membername="" then
				SendMsg=SendMsg+"<br>"+"<li>ֻ�л�Ա���ܹ������Ʒ�����޷�����"
				founderr=true
			end if
			if rsvisual("flag")=2 and not master and not boardmaster and not superboardmaster then
				SendMsg=SendMsg+"<br>"+"<li>ֻ�а���������Ա���ܹ������Ʒ�����޷�����"
				founderr=true
			end if
			if rsvisual("flag")=3 and not master and not superboardmaster then
				SendMsg=SendMsg+"<br>"+"<li>ֻ�г��������͹���Ա���ܹ������Ʒ�����޷�����"
				founderr=true
			end if
			if rsvisual("flag")=4 and not master then
				SendMsg=SendMsg+"<br>"+"<li>ֻ�й���Ա���ܹ������Ʒ�����޷�����"
				founderr=true
			end if
		end if
	end if
	rsvisual.close
	set rsvisual=nothing
	if founderr then
		call send_head()
		call Send_Msg()
	else
		if request("action")="send" then
			if request("myself")=1 then
				CurDate=year(now())&"-"&month(now())&"-"&day(now())
				conn.execute("update visual_myitems set username='"&username&"',FromUser='"&membername&"',Type="&CurType&" where username='"&membername&"' and itemid="&curItem)
				conn.execute("insert into visual_events (ItemID,UserName,FromUser,DateAndTime,Price,Type) values ("&CurItem&",'"&UserName&"','"&membername&"','"&curdate&"',"&curprice&",1)")
				set rs=server.createobject("adodb.recordset")
				sql="select * from [message]"      
				rs.open sql,conn,1,3      
				rs.addnew      
				rs("sender")=membername      
				rs("incept")=UserName
				rs("title")="����һ�������"&ItemName&"����"
				rs("content")=MsgBody&chr(10)&chr(10)&"��Ʒ�Ѿ��ŵ����Ĵ�����У���Ц�ɣ�"&chr(10)&"�뵽���µ�ַ�鿴�����"&chr(10)&"[align=center][URL=z_visual.asp?shopid=100][color=blue][b][u]�鿴�����[/u][/b][/color][/URL][/align]"
				rs("flag")=0 
				rs("issend")=1
				rs.update
				rs.close
				set rs=nothing
	  		conn.execute("update [user] set userwealth=userwealth-"&SendFee&" where username='"&membername&"'")
				SendMsg="<li>�Ѿ��ɹ��ظ� <font color=blue>"&username&"</font> �����ˡ�<font color=blue>"&itemName&"</font>����"
				call send_head()
				response.write "<script>"
				response.write "window.opener.document.execCommand('Refresh');"
				response.write "</script>"
				call Send_Msg()
			else
				if request("fromuser")<>"" then
					CurDate=year(now())&"-"&month(now())&"-"&day(now())
					conn.execute("update visual_myitems set username='"&username&"',FromUser='"&membername&"',Type="&CurType&",price=saleprice,saleprice=0 where username='"&checkstr(trim(request("fromuser")))&"' and type=5 and itemid="&curItem)
					conn.execute("update visual_events set username='"&membername&"' where itemid="&curitem&" and type=2 and fromuser='"&checkstr(trim(request("fromuser")))&"' and (username='' or isnull(username))")
					conn.execute("insert into visual_events (ItemID,UserName,FromUser,DateAndTime,Price,Type) values ("&CurItem&",'"&UserName&"','"&membername&"','"&curdate&"',"&saleprice&",1)")
					set rs=server.createobject("adodb.recordset")
					sql="select * from [message]"      
					rs.open sql,conn,1,3      
					rs.addnew      
					rs("sender")=membername      
					rs("incept")=UserName
					rs("title")="����һ�������"&ItemName&"����"
					rs("content")=MsgBody&chr(10)&chr(10)&"��Ʒ�Ѿ��ŵ����Ĵ�����У���Ц�ɣ�"&chr(10)&"�뵽���µ�ַ�鿴�����"&chr(10)&"[align=center][URL=z_visual.asp?shopid=100][color=blue][b][u]�鿴�����[/u][/b][/color][/URL][/align]"
					rs("flag")=0 
					rs("issend")=1
					rs.update
					rs.close
					set rs=nothing
		  		conn.execute("update [user] set userwealth=userwealth-"&(saleprice+SendFee)&" where username='"&membername&"'")
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
					SendMsg="<li>�Ѿ��ɹ��ظ� <font color=blue>"&username&"</font> �����˶�����Ʒ��<font color=blue>"&itemName&"</font>����"
					call send_head()
					response.write "<script>"
					response.write "window.opener.document.execCommand('Refresh');"
					response.write "</script>"
					call Send_Msg()
				else
					CartBag=""
					if CurItem>0 then
						for i=0 to itemcount-1
							if CartBag<>"" then CartBag=CartBag&"|"
							SendSplit=split(itemsplit(i),"$")
							if ubound(SendSplit)=2 then
								if SendSplit(0)=CurItem and lcase(trim(SendSplit(1)))=lcase(trim(UserName)) then
									SendSplit(2)=MsgBody
									SucFlag=true
								end if
								CartBag=CartBag&SendSplit(0)&"$"&SendSplit(1)&"$"&SendSplit(2)
							else
								CartBag=CartBag&ItemSplit(i)
							end if
						next
					end if
					if SucFlag then
						response.cookies("mybuy_"&userid)=CartBag
						SendMsg="<li>�Ѿ��޸��˸� <font color=blue>"&username&"</font> ����ͬ���ԣ�"
					else
						if itemcount<CartLength then
							if CartBag<>"" then CartBag=CartBag&"|"
							CartBag=CartBag&CurItem&"$"&username&"$"&msgbody
							response.cookies("mybuy_"&userid)=CartBag
							SendMsg="<li>�ѽ����͸� <font color=blue>"&username&"</font> �������Ʒ�ŵ����Ĺ�����У�"
						else
							founderr=true
							SendMsg="<li>���Ĺ�����������޷������͸� <font color=blue>"&username&"</font> �������Ʒ���룡"
						end if
					end if
					call send_head()
					call Send_Msg()
				end if
			end if
		else
			call send_head()%>
			<form action="?" method=post name=messager>
			<input type=hidden name="action" value="send">
			<table cellpadding=3 cellspacing=1 align=center bgcolor=<%=Forum_Body(27)%> width=97%>
				<tr> 
					<th colspan=3>����������Ʒ -- ��ͬ���ԣ�������������Ϣ��</th>
				</tr>
				<tr> 
					<td class=tablebody1 valign=middle><b>�����ˣ�</b></td>
					<td class=tablebody1 valign=middle colspan=2><%if username<>"" then%><input type=hidden name="username" value="<%=username%>"><b><%=username%></b><%else%><input type=text name="username" size=45> <SELECT onchange=document.messager.username.value=this.options[this.selectedIndex].value;document.messager.username.focus()><OPTION selected value="">ѡ��</OPTION><%
						set rs=server.createobject("adodb.recordset")
						sql="select F_friend from Friend where F_username='"&membername&"' and F_type=0 order by F_addtime desc"
						rs.open sql,conn,1,1
						do while not rs.eof
							%><OPTION value="<%=rs(0)%>"><%=rs(0)%></OPTION><%
							rs.movenext
						loop
						rs.close
						set rs=nothing
					%></SELECT><%end if%></td>
				</tr>
				<tr> 
					<td class=tablebody1 valign=top width=15%><b>���ݣ�</b></td>
					<td class=tablebody1 valign=middle><textarea cols=50 rows=6 name="msgbody" title="Ctrl+Enter����"><%=msgbody%></textarea></td>
         	<td valign=middle class=tablebody1 baseName=images/img_visual/blank.gif width="84" height="84" border="0" itemgender="" itemno=<%=CurItem%> layerno="" localpic=<%=LocalPic%> ImagePath="<%=PicPath%>" style="behavior:url(inc/z_show2.htc)"></td>
				</tr>
				<tr> 
					<td class=tablebody1 colspan=3><b>˵��</b>��<br>�� ������ʹ��<b>Ctrl+Enter</b>����ݷ�������<br>�� �������<b><%=GroupSetting(34)%></b>���ַ�<br><font color=<%=forum_body(8)%>>�� ���͸���Ʒ��������Ϊ��<%=SendFee%> Ԫ<%
						if request("fromuser")<>"" then
							response.write "<br>�� �������Ʒ�ķ���Ϊ��"&saleprice&" Ԫ"
						end if
					%></font></td>
				</tr>
				<tr> 
					<td class=tablebody2 valign=middle colspan=3 align=center><input type=hidden name=itemid value=<%=curitem%>><input type=hidden name=myself value=<%=request("myself")%>><%if request("fromuser")<>"" then%><input type=hidden name=fromuser value=<%=request("fromuser")%>><%end if%><input type=Submit value="����" name=Submit>&nbsp;<input type="reset" name="Clear" value="���">&nbsp;<input type="button" name="close" value="�ر�" onclick="window.close()"><%if SucFlag then%>&nbsp;<a href=z_visual_cartbag.asp>[���ع����]</a><%end if%></td>
				</tr>
			</table>
			</form>
		<%end if
	end if
end if
call footer()

sub send_head()%>
	<html><head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=Forum_info(0)%>--����������Ʒ</title>
	<!--#include file=inc/forum_css.asp-->
	</head>
	<body <%=Forum_body(11)%>" onkeydown="if(event.keyCode==13 && event.ctrlKey && document.messager=='[object]')messager.submit()">
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
			<td width="100%" class=tablebody2><a href="javascript:window.close()">[�رմ���]</a>&nbsp;&nbsp;<%
				if SucFlag then
					%><a href="z_visual_cartbag.asp">[���ع����]</a><%
				'else
					%><!--<a href="javascript:history.go(-1)">[������ҳ]</a>--><%
				end if
			%></td>
		</tr>
	</table>
<%end sub%>