<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%
stats="��Ʊ����"
call nav()
call head_var(0,0,GuPiao_Setting(5),"z_gp_gupiao.asp")
if not master then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	call AdminHead()%>
	<table class=tableborder1 cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
	<%select case request("action")
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
	end select%>
	</table>
<%end if
call footer()
'=====================================
sub main()%>
	<tr>
		<th valign=center align=middle height=25 colspan="8">��Ʊ���� | <a href="?action=newgupiao" class="cblue">�¹�����</a></th>
	</tr>
	<tr> 
		<form name="form1" method="post" action="?action=search">
		<td class=tablebody2 colspan="8" height=20>���ٲ��ң�<input type="text"  name=gname >��<input type="checkbox"  name="usernamechk" value="yes" checked>������ȫƥ�䣠<input type="submit" name="Submit" value="��ѯ��Ʊ">
		<div align="right"><font color=red>[ע��] ɾ������ֱ��ɾ����Ӧ�Ĺ�Ʊ���ֹ��߻��Ե�ǰ�۸��׳����й�Ʊ</font></div>
		</td>
		</form>
	</tr>
	<tr align="center" > 
		<th align="center">����</th>
		<th align="center">��Ʊ����</th>
		<th align="center">�ܹɷ�</th>
		<th align="center">�������̼�</th>
		<th align="center">Ŀǰ�ɽ���</th>
		<th align="center">�ɽ���</th>
		<th align="center">״̬</th>
		<th align="center">����</th>
	</tr>
	<%dim currentpage,page_count,Pcount
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
	sql= "select * from [GuPiao]"
	rs.open sql,gp_conn,1,1
	if rs.eof and rs.bof then
		response.write "<tr><td class=tablebody1 colspan=8>��û�����й�Ʊ��</td></tr>"
	else
		rs.PageSize = Gupiao_Setting(2) 
		rs.AbsolutePage=currentpage
		page_count=0
		totalrec=rs.recordcount
		do while (not rs.eof) and (not page_count = rs.PageSize)%>
		  <tr align="center"> 
		    <td class=tablebody1><a href="z_gp_Exchange.asp?sid=<%=rs("sid")%>"> <%=rs("sid")%> </a></td>
		    <td class=tablebody1><a href="?action=edit&sid=<%=rs("sid")%>"><%=rs("QiYe")%></a></td>
		    <td class=tablebody1 align=right><%=formatnumber(rs("JiaoYiLiang"),0)%>&nbsp;</td>
		    <td class=tablebody1 align=right><%=formatnumber(rs("KaiPanJiaGe"),2)%>&nbsp;</td>
		    <td class=tablebody1 align=right><%=formatnumber(rs("DangQianJiaGe"),2)%>&nbsp;</td>
				<td class=tablebody1 align=right><%=rs("ChengJiaoLiang")%>&nbsp;</td>
				<td class=tablebody1><%if rs("ZhuangTai")="��" then%><font color=red><%end if%><%=rs("ZhuangTai")%></td>
		    <td class=tablebody1><%if rs("ZhuangTai")="��" then%><a href="?action=ChangStat&stat=closethis&sid=<%=rs("sid")%>">����</a><%else%><a href="?action=ChangStat&stat=openthis&sid=<%=rs("sid")%>" class=cred>����</a><%end if%> | <a href="?action=edit&sid=<%=rs("sid")%>">�༭</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('��ȷ��ɾ�� <%=rs("QiYe")%> �����Ʊ��?')){return true;}return false;}" class=cred>ɾ��</a></td>
		  </tr>
			<%page_count = page_count + 1		
			rs.movenext
		loop
		Pcount=rs.PageCount%>
		<tr>
			<td class=tablebody2 colspan=2 align=left>���� <font color=blue><%=totalrec%></font> ֧��Ʊ��ÿҳ <font color=blue><%=rs.PageSize%></font> ֧���� <font color=red><%=currentpage%></font> ҳ / �� <font color=blue><%=Pcount%></font> ҳ</td>
			<td class=tablebody2 colspan=6 align=right>��ҳ��<%call disppagenum(currentpage,Pcount,"""?page=","""")%>&nbsp;&nbsp;&nbsp;&nbsp;</td>
		</tr>
	<%end if
	rs.close
	set rs=nothing	
end sub
'-----------------------
sub ChangStat()
	on error resume next
	if request("stat")="closethis" then
		gp_conn.execute("update [GuPiao] set ZhuangTai='��' where sid="&request("sid"))
	elseif request("stat")="openthis" then
		gp_conn.execute("update [GuPiao] set ZhuangTai='��' where sid="&request("sid"))		
	end if
	response.redirect "?"
end sub
'-------------�༭��Ʊ---------------
sub Edit()
	dim Sid,ExplainSplit
	Sid=request("sid")
	if Sid="" then
		errmess="<li>����������ָ����ع�Ʊ"
		call endinfo(2)
		exit sub	
	end if
	set rs=gp_conn.execute("select * from [GuPiao] where sid="&Sid)
	if rs.eof and rs.bof then
		rs.close
		errmess="<li>û���ҵ��ù�Ʊ"
		call endinfo(2)
		exit sub
	end if
	ExplainSplit=split(rs("Explain"),"|")
	if ubound(ExplainSplit)<1 then
		redim preserve ExplainSplit(1)
		ExplainSplit(1)=10000
	end if%>
	<tr>
		<th colspan="2" valign=center align=middle height=25>�༭ <font color="<%=forum_body(8)%>"><%=sid%></font> �Ź�Ʊ��<%=rs("QiYe")%></th>
	</tr>
	<form method=post  action="?action=SaveEdit">
	<tr height=25>
		<td class=tablebody2 width="40%"><b>����</b></td>
		<td class=tablebody1 width="60%">&nbsp;<%=rs("sid")%><input type=hidden name=sid value="<%=rs("sid")%>"></td>
	</tr>
	<tr height=25>
		<td class=tablebody2><b>��Ӫ��</b></td>
		<td class=tablebody1>&nbsp;<%=rs("JingYingZhe")%></td>
	</tr>	
	<tr>
		<td class=tablebody2><b>��ҵ</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=gpname value="<%=rs("QiYe")%>"><input type=hidden name=gpoldname value="<%=rs("QiYe")%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><b>�ܹɷ�</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=TotalStockNum value="<%=rs("ZongGuFen")%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><b>��ʼ������</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=inigpnum value="<%=ExplainSplit(1)%>"></td>
	</tr>	
	<tr>
		<td class=tablebody2><b>ʣ�ཻ����</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=gpnum value="<%=rs("JiaoYiLiang")%>"></td>
	</tr>			
	<tr>
		<td class=tablebody2><b>���̼۸�</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=openprice value="<%=round(rs("KaiPanJiaGe"),2)%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><b>��ǰ�۸�</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=nowprice value="<%=round(rs("DangQianJiaGe"),2)%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><b>��������</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=gpdate value="<%=rs("RiQi")%>"></td>
	</tr>	
	<tr>
		<td class=tablebody2><b>״̬</b></td>
		<td class=tablebody1><input type=radio name=state value="��" <%if rs("ZhuangTai")="��" then%> checked <%end if%>> ������<input type=radio name=state value="��" <%if rs("ZhuangTai")="��" then%> checked <%end if%>>��</td>
	</tr>
	<tr>
		<td class=tablebody2><b>��˾ͼƬ</b>(150*190)<br>���ͼƬ·�������·������ǰĿ¼�ǹ�Ʊ�ĸ�Ŀ¼<br>ͼƬҲ���������ϵ�ͼƬ����</td> 
		<td class=tablebody1>&nbsp;<input type=text name=LtdImg size="39" value="<%=rs("LtdImg")%>"></td> 
	</tr>	
	<tr>
		<td class=tablebody2 valign="top"><b>��˾���</b><br>���й�˾����<br>��˾��鲻�ܳ���100Byte</td>   
		<td class=tablebody1>&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL"><%=htmlencode(ExplainSplit(0))%></textarea></td>  
	</tr>

	<tr>
		<td class=tablebody2><font color=gray>ʣ��ɷ�</font></td>
		<td class=tablebody1>&nbsp;<input type=text readonly name=remaingp value="<%=rs("ShengYuGuFen")%>"></td> 
	</tr>			
	<tr>
		<td class=tablebody2><font color=gray>���ճɽ���</font></td>
		<td class=tablebody1>&nbsp;<input type=text readonly name=chengjiao value="<%=rs("ChengJiaoLiang")%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><font color=gray>�����������</font></td>
		<td class=tablebody1>&nbsp;<input type=text readonly name=buy value="<%=rs("MaiRuBiShu")%>"></td>
	</tr>
	<tr>
		<td class=tablebody2><font color=gray>������������</font></td>
		<td class=tablebody1>&nbsp;<input type=text readonly name=sale value="<%=rs("MaiChuBiShu")%>"></td>
	</tr>					
	<tr>
		<th colspan="2" align="center"><input type=submit name=submit value=ִ���޸�></th>
	</tr>
	</form>
<%rs.close
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
		set rs=gp_conn.execute("select sid from [GuPiao] where [QiYe]='"&GpName&"'")
		if not rs.eof then
			errmess="<li>����������й�˾�Ѿ����ڣ���������д���й�˾����"
		end if
	end if	
	
	if errmess<>"" then
		call endinfo(2)
		exit sub	
	end if	
	if isdate(gpdate) then 
		gpdate=" ,RiQi='"&gpdate&"'"
	else
		gpdate=""
	end if
	if Explain="" then Explain="����"
			
	gp_conn.execute("update [GuPiao] set [QiYe]='"&GpName&"',JiaoYiLiang="&gpnum&",KaiPanJiaGe="&OpenPrice&",DangQianJiaGe="&NowPrice&",ZhuangTai='"&Request.form("state")&"',ZongGuFen="&TotalStockNum&",LtdImg='"&LtdImg&"',Explain='"&Explain&"|"&inigpnum&"' "&gpdate&" where sid="&request("sid"))
	if gpoldname<>GpName then	'�ı��Ʊ����ʱ request("sid")
		gp_conn.execute("update [DaHu] set QiYe='"&GpName&"' where sid="&request("sid")) 
		gp_conn.execute("update [PersonalStock] set Stockname='"&GpName&"' where Stockname='"&gpoldname&"'")
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
	set rs=gp_conn.execute("select DangQianJiaGe,QiYe from [GuPiao] where sid="&DelID)
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
	gp_conn.execute("delete from [GuPiao] where sid="&DelID)
	set rs=gp_conn.execute("select * from [DaHu] where sid="&DelID)
	do while not rs.eof
		AddMoney=rs("ChiGuShu")*NowPrice
		gp_conn.execute "update [KeHu] set ZiJin=ZiJin+"&AddMoney&",ZongZiJin=ZongZiJin-"&AddMoney&",JinRiMaiChu=JinRiMaiChu+1,ZongMaiChu=ZongMaiChu+1 where id="&rs("uid")&""
		rs.movenext
	loop
	rs.close
	gp_conn.execute("delete from [DaHu] where sid="&DelID)
	sucmess="<font color=#66CCFF>"&GPname&"</font><font color=#CCCCCC>���գ������˸ùɵĿͻ������Ե�ǰ�۸�ȫ���׳�</font>"
	gp_conn.execute "insert into RndEvent(content,addtime) values('"&sucmess&"','"&now()&"' )"  
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
		sql="select * from [GuPiao] where QiYe = '"&GupiaoName&"'"  
	else	
		sql="select * from [GuPiao] where QiYe like '%"&GupiaoName&"%'"         
	end if%>
	<tr align="center" > 
		<th align="center">����</th>
		<th align="center">��Ʊ����</th>
		<th align="center">�ܹɷ�</th>
		<th align="center">�������̼�</th>
		<th align="center">Ŀǰ�ɽ���</th>
		<th align="center">�ɽ���</th>
		<th align="center">״̬</th>
		<th align="center">����</th>
	</tr>
	<%set rs=gp_conn.execute(sql)         
	if rs.eof then
		response.Write "<tr><td class=tablebody1 colspan=8>û���ҵ��κι�Ʊ</td></tr>"
	else
		do while not rs.eof%>
			<tr align="center"> 
				<td class=tablebody1><a href="z_gp_Exchange.asp?sid=<%=rs("sid")%>"> <%=rs("sid")%> </a></td>
				<td class=tablebody1><a href="?action=edit&sid=<%=rs("sid")%>"><%=rs("QiYe")%></a></td>
				<td class=tablebody1 align=right><%=formatnumber(rs("JiaoYiLiang"),0)%>&nbsp;</td>
				<td class=tablebody1 align=right><%=formatnumber(rs("KaiPanJiaGe"),2)%>&nbsp;</td>
				<td class=tablebody1 align=right><%=formatnumber(rs("DangQianJiaGe"),2)%>&nbsp;</td>
				<td class=tablebody1 align=right><%=rs("ChengJiaoLiang")%>&nbsp;</td>
				<td class=tablebody1><%if rs("ZhuangTai")="��" then%><font color=red><%end if%><%=rs("ZhuangTai")%></td>
				<td class=tablebody1><a href="?action=edit&sid=<%=rs("sid")%>">�༭</a> | <a href="?action=del&sid=<%=rs("sid")%>" onclick="javascript:{if(confirm('��ȷ��ɾ�� <%=rs("QiYe")%> �����Ʊ��?')){return true;}return false;}">ɾ��</a></td>
			</tr>	
  		<%rs.movenext
		loop
		rs.close
	end if%>
	<TR>
		<Td class=tablebody2 height=25 align="center" colspan="8"><A href="?">[�����б�]</A></Td>
	</TR>
<%end sub
'-----------------�¹�����---------------
sub newgupiao()%>
	<tr>
		<th colspan="2" valign=center align=middle height=25>�¹�����</th>
	</tr>
	<form method=post  action="?action=SaveNew">
	<tr>
		<td class=tablebody2 width="40%"><b>ID</b></td>
		<td class=tablebody1 width="60%">&nbsp;�Զ�����</td>
	</tr>
	<tr>
		<td class=tablebody2><b>��ҵ(��Ʊ��)</b>������Ҫ����20��</td>
		<td class=tablebody1>&nbsp;<input type=text name=gpname value=""> <font color=<%=forum_body(8)%>>*</font></td>
	</tr>
	<tr>
		<td class=tablebody2><b>��ʼ������</b><br>ÿ�ս�����</td>
		<td class=tablebody1>&nbsp;<input type=text name=gpnum value="10000"> ��</td> 
	</tr>
	<tr>
		<td class=tablebody2><b>�ܹɷ�</b><br>��Ʊ������</td> 
		<td class=tablebody1>&nbsp;<input type=text name=TotalStockNum value="10000000"> ��</td>
	</tr>				
	<tr>
		<td class=tablebody2><b>���̼۸�</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=openprice value="10.00"> Ԫ</td>
	</tr>
	<tr>
		<td class=tablebody2><b>��ǰ�۸�</b></td>
		<td class=tablebody1>&nbsp;<input type=text name=nowprice value="10.00"> Ԫ</td>
	</tr>
	<tr>
		<td class=tablebody2><b>״̬</b></td>
		<td class=tablebody1><input type=radio name=state value="��" checked> ������<input type=radio name=state value="��">��</td>
	</tr>
	<tr>
		<td class=tablebody2><b>��˾ͼƬ</b>(150*190)<br>���ͼƬ·�������·������ǰĿ¼�ǹ�Ʊ�ĸ�Ŀ¼<br>ͼƬҲ���������ϵ�ͼƬ����</td> 
		<td class=tablebody1>&nbsp;<input type=text name=LtdImg size="39" value="Images/LtdImg/1.jpg"></td> 
	</tr>	
	<tr>
		<td class=tablebody2 valign="top"><b>��˾���</b><br>���й�˾����<br>��˾��鲻�ܳ���100Byte</td>   
		<td class=tablebody1>&nbsp;<textarea cols="39" name="Explain" rows="5" wrap="PHYSICAL">����һ������ʲô��û�����£�</textarea></td>  
	</tr>	
	<tr>
		<td class=tablebody2><b>�ɽ���</b></td>
		<td class=tablebody1>&nbsp;0</td>
	</tr>	
	<tr>
		<td class=tablebody2><b>�ǵ�</b></td>
		<td class=tablebody1>&nbsp;0</td>
	</tr>	
	<tr>
		<td class=tablebody2><b>�������</b></td>
		<td class=tablebody1>&nbsp;0</td>
	</tr>
	<tr>
		<td class=tablebody2><b>��������</b></td>
		<td class=tablebody1>&nbsp;0</td>
	</tr>					
	<tr><th colspan="2" align="center"><input type=submit name=submit value=����>��<input type=reset name=submit value=����></th>
	</form>
<%end sub
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
		errmess=errmess+"<li>ÿ�ս�����������������"
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
	set rs=gp_conn.execute("select sid from [GuPiao] where QiYe='"&GpName&"'")
	if not rs.eof then
		errmess="<li>������Ĺ�Ʊ�����Ѿ������ˣ���������д��Ʊ����"
		endinfo(2)
		exit sub
	else
		gp_conn.execute("insert into [GuPiao](QiYe,JiaoYiLiang,KaiPanJiaGe,DangQianJiaGe,ZhuangTai,RiQi,ZongGuFen,LtdImg,Explain,TongJi) values('"&GpName&"',"&gpnum&","&OpenPrice&","&NowPrice&",'"&Request.form("state")&"',date(),"&TotalStockNum&",'"&LtdImg&"','"&Explain&"|"&gpnum&"','0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0')")
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
	set rs=gp_conn.execute("select sid from [GuPiao] order by sid")
	Cid=1
	do while not rs.eof
		gp_conn.execute("update [GuPiao] set Cid="&Cid&" where sid="&rs(0))
		Cid=Cid+1	
		rs.movenext
	loop
end sub
%>