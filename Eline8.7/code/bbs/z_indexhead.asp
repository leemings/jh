<table cellspacing=1 cellpadding=3 align=center class=tableBorder1 border=0>
	<tr>
		<%if founduser then%>
			<td align=center nowrap valign=middle class=tableBody1 rowspan=4 width=140>
				<%call ShowMyselfVisual(140,"tablebody1")%>
			</td>					
		<%else%>
			<td align=center nowrap valign=middle class=tableBody1 rowspan=4 width=140><a href="http://www.happyjh.com"><img src="images/welcome.gif" border="0" alt="һ�����绶ӭ����"></a></td>
		<%end if
		if not founduser then%>
			<td class=tableBody1 align=center valign=middle width=100% >��ӭ�����ι�,<a href=reg.asp class=cblue><font color=#ff0000>ע��</font></a>��<a href=login.asp class=cblue><font color=#ff0000>��¼</font></a>����ܲ鿴������Ϣ!</td><%
		else%>
			<td class=tableBody1 align=center valign=top width=100% >
				<TABLE border=0 width="98%" align=center>
					<TR height=20>
						<td><a href="JavaScript:openScript('messanger.asp?action=new',500,400)">������</a></td>
						<td width=10 align=center>��</td>
						<td><%
							if Cint(newincept())>Cint(0) then
								if Cint(forum_setting(10))=1 then
									%><script language=JavaScript>openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',640,400)</script><%
								end if
								%><bgsound src="<%=Forum_info(7)&Forum_statePic(8)%>" border=0><a href="usersms.asp?action=inbox">�ҵ��ռ���</a> (<a href="javascript:openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',500,400)"><font color="<%=Forum_body(8)%>"><%=newincept()%> ��</font></a>)<%
							else
								%><a href="usersms.asp?action=inbox">�ҵ��ռ���</a> (<font color=gray>0 ��</font>)<%
							end if
						%></td>
						<td width=10 align=center><font face=Wingdings color=blue>v</font></td>
						<td><font color="red"><%=myClass%></font></td>	
					</TR>
					<tr>
						<td height="1" bgcolor="<%=Forum_body(27)%>" colspan=5></td>
					</tr>
					<TR>
						<td><a href="topicwithme.asp?s=3">���������</a></td>
						<td width=10 align=center>��</td>
						<td><a href="topicwithme.asp?s=2">���ظ�������</a></td>
						<td width=10 align=center>��</td>
						<td><a href="topicwithme.asp?s=1">���������</a></td>
					</TR>
					<TR>
						<td>��������<b><%=myArticle%></b></td>
						<td width=10 align=center>��</td>
						<td>�ֽ�<b><%=mymoney%></b></td>
						<td width=10 align=center>��</td>
						<td>������<b><font color="red"><%=mypower%></b></font></td>
					</TR>
					<TR>
						<td>����ֵ��<b><%=myUserEP%></b></td>
						<td width=10 align=center>��</td>
						<td>������<b><%=myUserCP%></b></td>
						<td width=10 align=center>��</td>
						<td>���֣�<font color="red"><b><%=myUserScore%></b></font></td>
					</TR>
					<TR height=20>
						<td colspan=5><%call getRe()%><br>���ϴε�¼��ʱ���� <font color=#0000FF><b><%=LastLogin%></b></font></td>
					</TR>
				</TABLE>		
			</td>
		<%end if
		sql="select top 1 TopicNum,BbsNum,TodayNum,UserNum,lastUser,yesterdaynum,maxpostnum,maxpostdate,LastPost from config where active=1"
		set rs=conn.execute(sql)
		LastPostInfo=split(rs("LastPost"),"$")
		if not isdate(LastPostInfo(2)) then LastPostInfo(2)=Now()
		LastPostTime=LastPostInfo(2)
		if datediff("d",LastPostTime,Now())>0 and ubound(LastPostInfo)=7 then
			conn.execute("update config set yesterdaynum=todaynum,LastPost='"&checkStr(LastPostInfo(0))&"$"&checkStr(LastPostInfo(1))&"$"&Now()&"$"&checkStr(LastPostInfo(3))&"$"&checkStr(LastPostInfo(4))&"$"&checkStr(LastPostInfo(5))&"$"&checkStr(LastPostInfo(6))&"$"&checkStr(LastPostInfo(7))&"'")
			conn.execute("update config set todaynum=0")
		end if
		if rs(2)>rs(6) then conn.execute("update config set maxpostnum=todaynum,maxpostdate=Now()")%>
		<td class=tableBody1 align=center rowspan=4 valign=middle><%Call AllVisualPhoto()%></td>
	</tr>
	<tr>
		<td class=tableBody1 align=center valign=top height=1 width=100% ></td>
	</tr>
	<tr>
		<td class=tableBody1 align=center valign=top width=100% >
			<TABLE border=0 width="98%" align=center>			
				<TR>
					<TD>���շ�����<FONT COLOR="<%=Forum_body(8)%>"><B><%=rs(2)%></B></FONT></td>
					<TD>����������<b><%= rs(0) %></b></td>
				</TR>
				<TR>
					<TD>���շ�����<B><%=rs(5)%></B></td>
					<TD>����������<b><%=rs(1)%></B></td>
				</TR>
				<TR>
					<TD>��վ������<FONT COLOR="<%=Forum_body(8)%>"><B><%=datediff("d","2003-3-16",date())%></B></FONT></td>
					<TD>����շ�����<B><%=rs(6)%></B></td>
				</TR>
			</TABLE>
		</td>
	</tr>
	<tr>
		<td class=tableBody2 align="center" height="25" nowrap width=100% ><a href=queryresult.asp?stype=3>�鿴����</a><font face=Wingdings color=blue>v</font><a href=hotlist.asp?stype=1&searchdate=30>��������</a><font face=Wingdings color=blue>v</font><a href=toplist.asp?orders=1>��������</a><font face=Wingdings color=blue>v</font><a href=toplist.asp?orders=7>�����б�</a><font face=Wingdings color=blue>v</font><a href=cookies.asp?action=hasview>����Ѷ�</a></td>
	</tr>
</table>
<br>
<%Sub ScrollPhoto()
	dim rsPhoto,sjjh_name
        sjjh_name=Session("sjjh_name")
	set rsPhoto=server.createobject("Adodb.recordset")
	if PhotoListCount<=0 then
		sql="select p.URL from Visual_PhotoUser u inner join Visual_Photo p on u.photo_id=p.id where u.username='"&membername&"' and not isnull(p.url) and p.url<>'' order by p.id desc" 
	else
		sql="select top "&PhotoListCount&" p.URL from Visual_PhotoUser u inner join Visual_Photo p on u.photo_id=p.id where u.username='"&membername&"' and not isnull(p.url) and p.url<>'' order by p.id desc" 
	end if
	rsPhoto.open sql,conn,1,1
	if rsPhoto.eof or sjjh_name="" then
		response.write "<DIV style=""width:280""><font style=""color:#993366; FONT-SIZE:11pt"">����û������κ����࣡</font><br><br><a href=z_visual.asp?shopid=197><font style=""color:#ff0000"">����E����̳����ݡ�</font></a><br>���Ÿ���д��������Ѻ�Ӱ����ɣ�</font></DIV>"
	else
		response.write "<script language=""JavaScript"" src=""z_dhtmllib.js""></script>"
		response.write "<script language=""JavaScript"" src=""z_scroller.js""></script>"
		response.write "<script language=""JavaScript"">"
		response.write "var myScroller1 = new Scroller(0, 0, 282, 228, 0, 0);"
		response.write "myScroller1.setSpeed(1000);"		' ͼƬ�������ٶ�(����/��)
		response.write "myScroller1.setPause(5000);"	' ͼƬͣ����ʱ��(����)
		do while not rsPhoto.eof
			response.write "myScroller1.addItem('<div with=280 align=center><img src="""&rsPhoto(0)&""" style=""border:1px solid "&forum_body(27)&"""></div>');"&vbNewLine
			rsPhoto.movenext
		loop
		response.write "function runscroll() {"
		response.write "var layer;"
		response.write "var mikex, mikey;"
		response.write "layer = getLayer(""placeholder"");"
		response.write "mikex = getPageLeft(layer);"
		response.write "mikey = getPageTop(layer);"
		response.write "myScroller1.create();"
		response.write "myScroller1.hide();"
		response.write "myScroller1.moveTo(mikex, mikey);"
		response.write "myScroller1.show();"
		response.write "}"
		response.write "window.onload=runscroll"
		response.write "</script>"
		response.write "<div id=""placeholder"" aling=center style=""position:relative; width:282px; height:228px;""> </div>"
	end if
	rsPhoto.close
	set rsPhoto=nothing
end sub

function AllVisual()
	dim rsPhoto,sjjh_name
        sjjh_name=Session("sjjh_name")
	set rsPhoto=server.createobject("Adodb.recordset")
	sql="select URL from Visual_Photo where fromuser='"&membername&"' and not isnull(url) and url<>''" 
	rsPhoto.open sql,conn,1,1
	if rsPhoto.eof or sjjh_name="" then
		AllVisual="<DIV style=""width:280""><font style=""color:#993366; FONT-SIZE:11pt"">����û������κ����࣡</font><br><br><a href=z_visual.asp?shopid=197><font style=""color:#ff0000"">����E����̳����ݡ�</font></a><br>���Ÿ���д��������Ѻ�Ӱ����ɣ�</font></DIV>"
	else
		dim RndNum
		randomize
		RndNum=int(rsPhoto.RecordCount*rnd)
		if RndNum>0 then
			rsPhoto.Move RndNum
		end if
		AllVisual="<DIV style=""width:280""><img src="""&rsPhoto(0)&""" style=""border:1px solid "&forum_body(27)&"""></DIV>"
	end if
	rsPhoto.close
	set rsPhoto=nothing
end function

sub AllVisualPhoto()
	dim rsPhoto,sjjh_name
        sjjh_name=Session("sjjh_name")
	set rsPhoto=server.createobject("Adodb.recordset")
	sql="select * from Visual_Photo where fromuser='"&membername&"' and finished" 
	rsPhoto.open sql,conn,1,1
	if rsPhoto.eof or sjjh_name="" then
		response.write "<DIV style=""width:280""><font style=""color:#993366; FONT-SIZE:11pt"">����û������κ����࣡</font><br><br><a href=z_visual.asp?shopid=197><font style=""color:#ff0000"">����E����̳����ݡ�</font></a><br>���Ÿ���д��������Ѻ�Ӱ����ɣ�</font></DIV>"
	else
		dim RndNum
		randomize
		RndNum=int(rsPhoto.RecordCount*rnd)
		if RndNum>0 then
			rsPhoto.Move RndNum
		end if
		dim rsPhotoUser
		set rsPhotoUser=server.createobject("ADODB.recordset")
		dim usercount,uservisualsplit(),usernamesplit(),PhotoDate
		dim PhotoNameLeft,PhotoNameTop,PhotoNameFont,PhotoNameSize,PhotoNameBold,PhotoNameItalic,PhotoNameColor,PhotoNameDirection
		dim DateLeft,DateTop,DateFont,DateSize,DateBold,DateItalic,DateColor,DateDirection
		dim NameLeft(),NameTop(),NameFont(),NameSize(),NameBold(),NameItalic(),NameColor(),NameDirection()
		dim OuterLeft(),OuterTop(),OuterWidth(),OuterHeight()
		dim InnerLeft(),InnerTop(),InnerWidth(),InnerHeight()
		dim LayerNo(),Direction()
		sql="select * from visual_photouser where photo_id="&rsPhoto("ID")
		rsPhotoUser.open sql,conn,1,1
		usercount=0
		do while not rsPhotoUser.eof
			redim preserve uservisualsplit(usercount)
			uservisualsplit(usercount)=rsPhotoUser("uservisual")
			redim preserve usernamesplit(usercount)
			usernamesplit(usercount)=rsPhotoUser("username")
			redim preserve LayerNo(UserCount)
			redim preserve OuterLeft(UserCount)
			redim preserve OuterTop(UserCount)
			redim preserve OuterWidth(UserCount)
			redim preserve OuterHeight(UserCount)
			redim preserve InnerLeft(UserCount)
			redim preserve InnerTop(UserCount)
			redim preserve InnerWidth(UserCount)
			redim preserve InnerHeight(UserCount)
			redim preserve Direction(UserCount)
			redim preserve NameLeft(UserCount)
			redim preserve NameTop(UserCount)
			redim preserve NameFont(UserCount)
			redim preserve NameSize(UserCount)
			redim preserve NameBold(UserCount)
			redim preserve NameItalic(UserCount)
			redim preserve NameColor(UserCount)
			redim preserve NameDirection(UserCount)
			LayerNo(UserCount)=rsPhotoUser("LayerNo")
			OuterLeft(UserCount)=rsPhotoUser("OuterLeft")
			OuterTop(UserCount)=rsPhotoUser("OuterTop")
			OuterWidth(UserCount)=rsPhotoUser("OuterWidth")
			OuterHeight(UserCount)=rsPhotoUser("OuterHeight")
			InnerLeft(UserCount)=rsPhotoUser("InnerLeft")
			InnerTop(UserCount)=rsPhotoUser("InnerTop")
			InnerWidth(UserCount)=rsPhotoUser("InnerWidth")
			InnerHeight(UserCount)=rsPhotoUser("InnerHeight")
			Direction(UserCount)=rsPhotoUser("Direction")
			NameLeft(UserCount)=rsPhotoUser("NameLeft")
			NameTop(UserCount)=rsPhotoUser("NameTop")
			NameFont(UserCount)=rsPhotoUser("NameFont")
			NameSize(UserCount)=rsPhotoUser("NameSize")
			NameBold(UserCount)=rsPhotoUser("NameBold")
			NameItalic(UserCount)=rsPhotoUser("NameItalic")
			NameColor(UserCount)=rsPhotoUser("NameColor")
			NameDirection(UserCount)=rsPhotoUser("NameDirection")
			usercount=usercount+1
			rsPhotoUser.movenext
		loop
		rsPhotoUser.close
		sql="select * from visual_Accouterment where photo_id="&rsPhoto("ID")
		rsPhotoUser.open sql,conn,1,1
		redim preserve uservisualsplit(usercount+5)
		redim preserve OuterLeft(UserCount+5)
		redim preserve OuterTop(UserCount+5)
		redim preserve OuterWidth(UserCount+5)
		redim preserve OuterHeight(UserCount+5)
		redim preserve InnerLeft(UserCount+5)
		redim preserve InnerTop(UserCount+5)
		redim preserve InnerWidth(UserCount+5)
		redim preserve InnerHeight(UserCount+5)
		redim preserve Direction(UserCount+5)
		for i=0 to 5
			uservisualsplit(usercount+i)=""
			OuterLeft(UserCount+i)=0
			OuterTop(UserCount+i)=0
			OuterWidth(UserCount+i)=140
			OuterHeight(UserCount+i)=226
			InnerLeft(UserCount+i)=0
			InnerTop(UserCount+i)=0
			InnerWidth(UserCount+i)=140
			InnerHeight(UserCount+i)=226
			Direction(UserCount+i)=0
		next
		do while not rsPhotoUser.eof
			i=rsPhotoUser("SeqNo")
			uservisualsplit(usercount+i)=rsPhotoUser("ItemPicPath")
			OuterLeft(UserCount+i)=rsPhotoUser("OuterLeft")
			OuterTop(UserCount+i)=rsPhotoUser("OuterTop")
			OuterWidth(UserCount+i)=rsPhotoUser("OuterWidth")
			OuterHeight(UserCount+i)=rsPhotoUser("OuterHeight")
			InnerLeft(UserCount+i)=rsPhotoUser("InnerLeft")
			InnerTop(UserCount+i)=rsPhotoUser("InnerTop")
			InnerWidth(UserCount+i)=rsPhotoUser("InnerWidth")
			InnerHeight(UserCount+i)=rsPhotoUser("InnerHeight")
			Direction(UserCount+i)=rsPhotoUser("Direction")
			rsPhotoUser.movenext
		loop
		rsPhotoUser.close
		set rsPhotoUser=nothing
		response.write "<DIV style=""width:280"">"
		call ShowVisualPhoto(rsPhoto("Width"),rsPhoto("Height"),rsPhoto("Background"),rsPhoto("BackBody"),rsPhoto("ForeBody"),rsPhoto("Foreground"),rsPhoto("Name"),rsPhoto("PhotoNameLeft"),rsPhoto("PhotoNameTop"),rsPhoto("PhotoNameFont"),rsPhoto("PhotoNameSize"),rsPhoto("PhotoNameBold"),rsPhoto("PhotoNameItalic"),rsPhoto("PhotoNameColor"),rsPhoto("PhotoNameDirection"),rsPhoto("AddDate"),rsPhoto("DateLeft"),rsPhoto("DateTop"),rsPhoto("DateFont"),rsPhoto("DateSize"),rsPhoto("DateBold"),rsPhoto("DateItalic"),rsPhoto("DateColor"),rsPhoto("DateDirection"),UserCount,UserVisualSplit,LayerNo,OuterLeft,OuterTop,OuterWidth,OuterHeight,InnerLeft,InnerTop,InnerWidth,InnerHeight,Direction,UserNameSplit,NameLeft,NameTop,NameFont,NameSize,NameBold,NameItalic,NameColor,NameDirection,False,rsPhoto("isUpload"),(rsPhoto("Child")=1))
		response.write "</DIV>"
	end if
	rsPhoto.close
	set rsPhoto=nothing
end sub%>