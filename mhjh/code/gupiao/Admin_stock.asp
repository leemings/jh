<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<html><head><title><%=Gupiao_Setting(5)%>-���п������</title>
<!--#include file="css.asp"-->
</head><body bgcolor="#ffffff" text="#000000" style="FONT-SIZE: 9pt" topmargin=5 leftmargin=0 oncontextmenu=self.event.returnValue=false>
<table  width="98%" border=0 cellspacing=1 cellpadding=0 align=center bgcolor="#0066CC">
	<TR>
		<TD BACKGROUND="Images/topbg.gif" height=9 colspan=3>
	</TD>
	</TR>
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif"><font size="2"><b>���й�������</b></font></td>
	</tr>
	<tr><td bgcolor="#E4E8EF"> 
<% 
call AdminHead()
if not master then
	errmess="<li>��ҳΪ����Աר�ã���û�й���ҳ��Ȩ�ޣ�"
	call endinfo(1)
else 
	select case request("action")
		case "search"
			call search()
		case "edit"
			call Edit()
		case "SaveEdit"		
			call SaveEdit()
		case "del"
			call del()
		case "newgupiao"
			call newgupiao()
		case "SaveNew"		
			call SaveNew()
		case "ChangStat"
			call ChangStat()					
		case else
			call main()
	end select
end if
%>
<br>
</td></tr>
<tr><td height=32 background="images/footer.gif" valign=middle></td></tr>
</table>
</body></html>
<%
CloseDatabase		'�ر����ݿ�
'=====================================
sub main()
%>
<table width="97%" border="0" cellspacing="1" cellpadding="3" align="center" bgcolor="#0066CC">
	<tr>
		<td valign=center align=middle height=23 background="Images/Header.gif" colspan="8">��Ʊ���� | <a href="?action=newgupiao" class="cblue">�¹�����</a></td>
	</tr>
	<tr> 
		<form name="form1" method="post" action="?action=search">
		<td colspan="8" height=20 bgcolor="#ffffff">���ٲ��ң�<input type="text"  name=gname >��<input type="checkbox"  name="usernamechk" value="yes" checked>������ȫƥ�䣠<input type="submit" name="Submit" value="��ѯ��Ʊ">
		<div align="right"><font color=red>[ע��] ɾ������ֱ��ɾ����Ӧ�Ĺ�Ʊ���ֹ��߻��Ե�ǰ�۸��׳����й�Ʊ</font></div>
		</td>
		</form>
	</tr>
	<tr align="center" > 
		<td align="center" background="Images/lan.gif"><font color=#ffffff>����</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>��Ʊ����</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>�ܹɷ�</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>�������̼�</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>Ŀǰ�ɽ���</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>�ɽ���</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>״̬</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>����</font></td>
	</tr>
<%        
	dim currentpage,page_count,Pcount
	dim totalrec,endpage
	currentPage=request("page")
	if currentpage="" or not isnumeric(currentpage) then
		currentpage=1
	else
		currentpage=clng(currentpage)
		if err then
			currentpage=1
			err.clear
		end if
	end if
	Set rs= Server.CreateObject("ADODB.Recordset")
	sql= "select * from ��Ʊ"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td colspan=8 bgcolor=#E4E8EF>��û�����й�Ʊ��</td></tr>"
	else
		rs.PageSize = 20 
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		do while (not rs.eof) and (not page_count = 20)	
%>
  <tr align="center"> 
    <td bgcolor="#FFFFFF" ><a href="exchange.asp?sid=<%=rs("sid")%>"> <%=rs("sid")%> </a></td>
    <td bgcolor="#FFFFFF"><a href="?action=edit&sid=<%=rs("sid")%>"><%=rs("��ҵ")%></a></td>
    <td bgcolor="#FFFFFF"><%=formatnumber(rs("�ܹɷ�"),0)%></td>
    <td bgcolor="#FFFFFF"><%=formatcurrency(rs("���̼۸�"),3)%></td>
    <td bgcolor="#FFFFFF"><%=formatcurrency(rs("��ǰ�۸�"),3)%></td>
	<td bgcolor="#FFFFFF"><%=rs("�ɽ���")%></td>
	<td bgcolor="#FFFFFF"><%if rs("״̬")="��" then%><font color=red><%end if%><%=rs("״̬")%></td>
    <td bgcolor="#FFFFFF"><%if rs("״̬")="��" then%><a href="?action=ChangStat&stat=closethis&sid=<%=rs("sid")%>">����</a><%else%><a href="?action=ChangStat&stat=openthis&sid=<%=rs("sid")%>" class=cred>����</a><%end if%> | <a href="?action=edit&sid=<%=rs("sid")%>">�༭</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('��ȷ��ɾ�� <%=rs("��ҵ")%> �����Ʊ��?')){return true;}return false;}" class=cred>ɾ��</a></td>
  </tr>
<%         
		page_count = page_count + 1		
		rs.movenext
	loop
	Pcount=rs.PageCount         
%>
	<tr><td colspan=8 background="images/title.gif">
		<table border="0" cellpadding="0" cellspacing="0" width="100%"><tr>
			<td align=left>����<font color=blue><%=totalrec%></font>֧��Ʊ��ÿҳ<font color=blue><%=rs.PageSize%></font>֧����<font color=red><%=currentpage%></font>ҳ/��<font color=blue><%=Pcount%></font>ҳ</td>
			<td align=right>��ҳ��
<%
			if currentpage > 4 then
				response.write "<a href=""?page=1"">[1]</a> ..."
			end if
			if Pcount>currentpage+3 then
			endpage=currentpage+3
			else
			endpage=Pcount
			end if
			for i=currentpage-3 to endpage
			if not i<1 then
				if i = clng(currentpage) then
				response.write " <font color=red>["&i&"]</font>"
				else
				response.write " <a href=""?page="&i&""">["&i&"]</a>"
				end if
			end if
			next
			if currentpage+3 < Pcount then 
				response.write "... <a href=""?page="&Pcount&""">["&Pcount&"]</a>"
			end if
				
		end if
		rs.close
		set rs=nothing	
%>			
			</td></tr>
		 </table>
	</td></tr>	  	
</table>
<% 
end sub
'-----------------------
sub ChangStat()
	on error resume next
	if request("stat")="closethis" then
		conn.execute("update ��Ʊ set ״̬='��' where sid="&request("sid"))
	elseif request("stat")="openthis" then
		conn.execute("update ��Ʊ set ״̬='��' where sid="&request("sid"))		
	end if
	response.redirect "?"
end sub
'-------------�༭��Ʊ---------------
sub Edit()
	dim Sid
	Sid=request("sid")
	if Sid="" then
		errmess="<li>����������ָ����ع�Ʊ"
		call endinfo(2)
		exit sub	
	end if
	set rs=conn.execute("select * from [��Ʊ] where sid="&Sid)
	if rs.eof and rs.bof then
		rs.close
		errmess="<li>û���ҵ��ù�Ʊ"
		call endinfo(2)
		exit sub
	end if
%>
<table  width="75%" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9 colspan="2"></TD>
	</TR>
	<tr>
		<td colspan="2" valign=center align=middle height=23 class=big background="images/header.gif">�༭ <font face="Courier New, Courier, mono" color="#FF0000"><%=sid%></font> �Ź�Ʊ��<%=rs("��ҵ")%></td>
	</tr>
	<form method=post  action="?action=SaveEdit">
	<tr>
		<td BGCOLOR="#E4E4EF" width="40%">����</td>
		<td bgcolor="#FFFFFF" width="60%">&nbsp;<%=rs("sid")%><input type=hidden name=sid value="<%=rs("sid")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">��Ӫ��</td>
		<td bgcolor="#FFFFFF">&nbsp;<%=rs("��Ӫ��")%></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">��ҵ</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpname value="<%=rs("��ҵ")%>"><input type=hidden name=gpoldname value="<%=rs("��ҵ")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">�ܹɷ�</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TotalStockNum value="<%=rs("�ܹɷ�")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">��ʼ������</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=inigpnum value="<%=rs("IniTradeNum")%>"></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">ʣ�ཻ����</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpnum value="<%=rs("������")%>"></td>
	</tr>			
	<tr>
		<td BGCOLOR="#E4E4EF">���̼۸�</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=openprice value="<%=rs("���̼۸�")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">��ǰ�۸�</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=nowprice value="<%=rs("��ǰ�۸�")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">��������</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpdate value="<%=rs("����")%>"></td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">״̬</td>
		<td bgcolor="#FFFFFF"><input type=radio name=state value="��" <%if rs("״̬")="��" then%> checked <%end if%>> ������<input type=radio name=state value="��" <%if rs("״̬")="��" then%> checked <%end if%>>��</td>
	</tr>

	<tr>
		<td BGCOLOR="#E4E4EF">��˾ͼƬ(150*190)<br>���ͼƬ·�������·������ǰĿ¼�ǹ�Ʊ�ĸ�Ŀ¼<br>ͼƬҲ���������ϵ�ͼƬ����</td> 
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=LtdImg size="39" value="<%=rs("LtdImg")%>"></td> 
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF" valign="top"><b>��˾��飺</b><br>���й�˾����<br>��˾��鲻�ܳ���100Byte</td>   
		<td bgcolor="#FFFFFF">&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL"><%=htmlencode(rs("Explain"))%></textarea></td>  
	</tr>

	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>ʣ��ɷ�</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text readonly name=remaingp value="<%=rs("ʣ��ɷ�")%>"></td> 
	</tr>			
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>���ճɽ���</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text readonly name=chengjiao value="<%=rs("�ɽ���")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>�����������</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text readonly name=buy value="<%=rs("�������")%>"></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><font color=gray>������������</font></td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text readonly name=sale value="<%=rs("��������")%>"></td>
	</tr>					

	<tr><td colspan="2" bgcolor="#FFFFFF" align="center"><input type=submit name=submit value=ִ���޸�></td>
	</form>
</table>
<%	
	rs.close
end sub
'-------------�����޸�---------------
sub SaveEdit()
	dim GpName,gpnum,OpenPrice,NowPrice,inigpnum
	dim TotalStockNum,gpdate,LtdImg,Explain,gpoldname
	GpName=trim(replace(Request.form("gpname"),"'","")) 
	gpoldname=trim(replace(Request.form("gpoldname"),"'","")) 
	gpnum=trim(Request.form("gpnum"))
	inigpnum=trim(Request.form("inigpnum"))
	
	OpenPrice=trim(Request.form("openprice"))
	NowPrice=trim(Request.form("nowprice"))
	
	TotalStockNum=trim(Request.form("TotalStockNum")) 
	gpdate=trim(Request.form("gpdate"))
	LtdImg=trim(Request.form("LtdImg"))
	Explain=trim(replace(Request.form("Explain"),"'",""))	
	
	if GpName="" then
		errmess="<li>���������й�Ʊ������"
		call endinfo(2)
		exit sub
	end if	
	errmess=""
	if not isnumeric(gpnum) then 
		errmess=errmess+"<li>������������������"
	end if
	if not isnumeric(inigpnum) then 
		errmess=errmess+"<li>��ʼ������������������"
	end if		
	if not isnumeric(OpenPrice) then 
		errmess=errmess+"<li>���̼۸������������"
	end if	
	if not isnumeric(NowPrice) then 
		errmess=errmess+"<li>��ǰ�۸������������"
	end if
	
	if not isnumeric(TotalStockNum) then 
		errmess=errmess+"<li>��Ʊ��������������"
	end if
	
	if gpoldname<>GpName then
		set rs=conn.execute("select sid from [��Ʊ] where [��ҵ]='"&GpName&"'")
		if not rs.eof then
			errmess="<li>����������й�˾�Ѿ����ڣ���������д���й�˾����"
		end if
	end if	
	
	if errmess<>"" then
		call endinfo(2)
		exit sub	
	end if	
	if isdate(gpdate) then 
		gpdate=" ,����='"&gpdate&"'"
	else
		gpdate=""
	end if
	if Explain="" then Explain="����"
			
	conn.execute("update [��Ʊ] set [��ҵ]='"&GpName&"',������="&gpnum&",IniTradeNum="&inigpnum&",���̼۸�="&OpenPrice&",��ǰ�۸�="&NowPrice&",״̬='"&Request.form("state")&"',�ܹɷ�="&TotalStockNum&",LtdImg='"&LtdImg&"',Explain='"&Explain&"' "&gpdate&" where sid="&request("sid"))
	if gpoldname<>GpName then	'�ı��Ʊ����ʱ request("sid")
		conn.execute("update [��] set ��ҵ='"&GpName&"' where sid="&request("sid")) 
		conn.execute("update [PersonalStock] set Stockname='"&GpName&"' where Stockname='"&gpoldname&"'")
	end if
	
	sucmess="<li>��Ʊ��Ϣ�Ѿ��޸����!"
	rUrl="?"
	call endinfo(2)
end sub
'---------------ɾ����Ʊ-------------
sub del()
	dim DelID,AddMoney,NowPrice,trs,GPname
	DelID=request("sid")
	if DelID="" then
		errmess="����������ָ��Ҫɾ���Ĺ�Ʊ"
		call endinfo(2)
		exit sub	
	end if
	set rs=conn.execute("select ��ǰ�۸�,��ҵ from ��Ʊ where sid="&DelID)
	if rs.eof and rs.bof then
		rs.close
		errmess="û���ҵ���Ҫɾ���Ĺ�Ʊ"
		call endinfo(2)
		exit sub
	else
		NowPrice=rs(0)
		GPname=rs(1)
	end if
	rs.close
	conn.execute("delete from [��Ʊ] where sid="&DelID)
	set rs=conn.execute("select * from [��] where sid="&DelID)
	do while not rs.eof
		AddMoney=rs("�ֹ���")*NowPrice
		conn.execute "update [�ͻ�] set �ʽ�=�ʽ�+"&AddMoney&",���ʽ�=���ʽ�-"&AddMoney&",��������=��������+1,������=������+1 where id="&rs("uid")&""
		rs.movenext
	loop
	rs.close
	conn.execute("delete from [��] where sid="&DelID)
	sucmess="<font color=#66CCFF>"&GPname&"</font><font color=#CCCCCC>���գ������˸ùɵĿͻ������Ե�ǰ�۸�ȫ���׳�</font>"
	conn.execute "insert into RndEvent(content,addtime) values('"&sucmess&"','"&now()&"' )"  
	call ResetCid()		'�������й�Ʊ��Cid
	sucmess="<li>"&GPname&" ��Ʊɾ���ɹ��������˸ù�Ʊ�����пͻ���ǿ���Ե�ǰ�۸��׳�"
	call endinfo(2)
end sub
'---------------������Ʊ-------------
sub search()
	dim GupiaoName
	GupiaoName=trim(replace(request("gname"),"'",""))
	if GupiaoName="" then
		errmess="<li>��������ҹؼ���"
		call endinfo(2)
		exit sub
	end if	
	if request("usernamechk")="yes" then
		sql="select * from ��Ʊ where ��ҵ = '"&GupiaoName&"'"  
	else	
		sql="select * from ��Ʊ where ��ҵ like '%"&GupiaoName&"%'"         
	end if	
	
%>
<table border="0" width="97%" bgcolor="#0066CC" cellspacing="1" cellpadding="3" align="center">
	<tr align="center" > 
		<td align="center" background="Images/lan.gif"><font color=#ffffff>����</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>��Ʊ����</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>�ܹɷ�</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>�������̼�</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>Ŀǰ�ɽ���</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>�ɽ���</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>״̬</font></td>
		<td align="center" background="Images/lan.gif"><font color=#ffffff>����</font></td>
	</tr>
<%		
	set rs=conn.execute(sql)         
	if rs.eof then
		response.Write "<tr><td bgcolor=""#FFFFFF"" colspan=8>û���ҵ��κι�Ʊ</td></tr>"
	else
		do while not rs.eof
%>
	<tr align="center"> 
		<td bgcolor="#FFFFFF" ><a href="exchange.asp?sid=<%=rs("sid")%>"> <%=rs("sid")%> </a></td>
		<td bgcolor="#FFFFFF"><a href="?action=edit&sid=<%=rs("sid")%>"><%=rs("��ҵ")%></a></td>
		<td bgcolor="#FFFFFF"><%=formatnumber(rs("�ܹɷ�"),0)%></td>
		<td bgcolor="#FFFFFF"><%=formatcurrency(rs("���̼۸�"),3)%></td>
		<td bgcolor="#FFFFFF"><%=formatcurrency(rs("��ǰ�۸�"),3)%></td>
		<td bgcolor="#FFFFFF"><%=rs("�ɽ���")%></td>
		<td bgcolor="#FFFFFF"><%if rs("״̬")="��" then%><font color=red><%end if%><%=rs("״̬")%></td>
		<td bgcolor="#FFFFFF"><a href="?action=edit&sid=<%=rs("sid")%>">�༭</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('��ȷ��ɾ�� <%=rs("��ҵ")%> �����Ʊ��?')){return true;}return false;}">ɾ��</a></td>
	</tr>	
  <%
			rs.movenext
		loop
		rs.close
	end if
%>
	<TR><TD BACKGROUND="Images/title.gif" height=21 align="center" colspan="8"><A href="?">[�����б�]</A></TD></TR>
</table>
<%	
end sub
'-----------------�¹�����---------------
sub newgupiao()
%>
<table  width="75%" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="#0066CC">
	<TR>
	<TD BACKGROUND="images/topbg.gif" height=9 colspan="2"></TD>
	</TR>
	<tr>
		<td colspan="2" valign=center align=middle height=23 class=big background="images/header.gif">�¹�����</td>
	</tr>
	<form method=post  action="?action=SaveNew">
	<tr>
		<td BGCOLOR="#E4E4EF" width="40%">ID</td>
		<td bgcolor="#FFFFFF" width="60%">&nbsp;�Զ�����</td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>��ҵ(��Ʊ��)</b>������Ҫ����20��</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpname value=""> <font color=red>*</font></td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>������</b><br>���ս�����</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=gpnum value="10000"> ��</td> 
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>�ܹɷ�</b><br>��Ʊ������</td> 
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=TotalStockNum value="10000000"> ��</td>
	</tr>				
	<tr>
		<td BGCOLOR="#E4E4EF">���̼۸�</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=openprice value="10.00"> Ԫ</td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">��ǰ�۸�</td>
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=nowprice value="10.00"> Ԫ</td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">״̬</td>
		<td bgcolor="#FFFFFF"><input type=radio name=state value="��" checked> ������<input type=radio name=state value="��">��</td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF"><b>��˾ͼƬ(150*190)</b><br>���ͼƬ·�������·������ǰĿ¼�ǹ�Ʊ�ĸ�Ŀ¼<br>ͼƬҲ���������ϵ�ͼƬ����</td> 
		<td bgcolor="#FFFFFF">&nbsp;<input type=text name=LtdImg size="39" value="Images/LtdImg/1.jpg"></td> 
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF" valign="top"><b>��˾��飺</b><br>���й�˾����<br>��˾��鲻�ܳ���100Byte</td>   
		<td bgcolor="#FFFFFF">&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL">����һ������ʲô��û�����£�</textarea></td>  
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">�ɽ���</td>
		<td bgcolor="#FFFFFF">&nbsp;0</td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">�ǵ�</td>
		<td bgcolor="#FFFFFF">&nbsp;0</td>
	</tr>	
	<tr>
		<td BGCOLOR="#E4E4EF">�������</td>
		<td bgcolor="#FFFFFF">&nbsp;0</td>
	</tr>
	<tr>
		<td BGCOLOR="#E4E4EF">��������</td>
		<td bgcolor="#FFFFFF">&nbsp;0</td>
	</tr>					

	<tr><td colspan="2" bgcolor="#FFFFFF" align="center"><input type=submit name=submit value=����>��<input type=reset name=submit value=����></td>
	</form>
</table>
<%
end sub
'-------------�����¹�Ʊ---------------
sub SaveNew()
	dim GpName,gpnum,OpenPrice,NowPrice
	dim TotalStockNum,LtdImg,Explain,gpoldname
	GpName=trim(replace(Request.form("gpname"),"'","")) 
	gpnum=trim(Request.form("gpnum"))
	OpenPrice=trim(Request.form("openprice"))
	NowPrice=trim(Request.form("nowprice"))
	
	if GpName="" then
		errmess="<li>���������й�Ʊ������"
		call endinfo(2)
		exit sub
	end if	
	TotalStockNum=trim(Request.form("TotalStockNum")) 
	LtdImg=trim(Request.form("LtdImg"))
	Explain=trim(replace(Request.form("Explain"),"'",""))	
	
	if GpName="" then
		errmess="<li>���������й�Ʊ������"
		call endinfo(2)
		exit sub
	end if	
	errmess=""
	if not isnumeric(gpnum) then 
		errmess=errmess+"<li>ʣ�ཻ����������������"
	end if
	if not isnumeric(OpenPrice) then 
		errmess=errmess+"<li>���̼۸������������"
	end if	
	if not isnumeric(NowPrice) then 
		errmess=errmess+"<li>��ǰ�۸������������"
	end if
	
	if not isnumeric(TotalStockNum) then 
		errmess=errmess+"<li>��Ʊ��������������"
	end if
	if errmess<>"" then
		call endinfo(2)
		exit sub	
	end if	
	if Explain="" then Explain="����"
	on error resume next 
	set rs=conn.execute("select sid from ��Ʊ where ��ҵ='"&GpName&"'")
	if not rs.eof then
		errmess="<li>������Ĺ�Ʊ�����Ѿ������ˣ���������д��Ʊ����"
		endinfo(2)
		exit sub
	else
		conn.execute("insert into ��Ʊ(��ҵ,������,���̼۸�,��ǰ�۸�,״̬,����,�ܹɷ�,ʣ��ɷ�,LtdImg,Explain,IniTradeNum,TongJi) values('"&GpName&"',"&gpnum&","&OpenPrice&","&NowPrice&",'"&Request.form("state")&"',date(),"&TotalStockNum&","&TotalStockNum&",'"&LtdImg&"','"&Explain&"',"&gpnum&",'0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0')")
	end if
	if err.number=0 then
		sucmess="<li>��ϲ���¹�Ʊ <font color=blue>"&Request.form("gpname")&"</font> �ɹ�����"
		rUrl="?"
		call ResetCid()
	else
		errmess="<li>���ִ��󣬾������£�"&Err.Description
	end if	
	call endinfo(2)
end sub
'---------------������Ʊ��Cid--------------------
sub ResetCid()
'˵����Cid ����������¼��ģ������Ǵ�1��ʼ����Ȼ����1��2��3��4��5........
	Dim Cid 
    set rs=conn.execute("select sid from [��Ʊ] order by sid")
	Cid=1
	do while not rs.eof
		conn.execute("update [��Ʊ] set Cid="&Cid&" where sid="&rs(0))
		Cid=Cid+1	
		rs.movenext
	loop
end sub
%>