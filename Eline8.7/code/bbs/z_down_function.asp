<%dim rsdown,sqldown,flagdown,rsp,rss,downpath,upmod
dim daytop,weektop,alltop,listnum,editnum,hotnum,eventnum,newnum,gudinnum,vipshow,sonum,paymod,retime,showset
sqldown="select flag from [admin] where username='"&membername&"'"
set rsdown=conndown.execute(sqldown)
if not(rsdown.eof and rsdown.bof) then
	flagdown=rsdown("flag")
else
	flagdown=0
end if
set rsp=conndown.execute("select downpath from [notdown]")
downpath=rsp("downpath")
if right(downpath,1)<>"/" then
	downpath=downpath&"/"
end if
rsp.close
set rsp=nothing
set rss=conndown.execute("select * from [showpage]")
daytop=rss("daytop")
weektop=rss("weektop")
alltop=rss("alltop")
listnum=rss("listnum")
editnum=rss("editnum")
eventnum=rss("eventnum")
hotnum=rss("hotnum")
newnum=rss("newnum")
gudinnum=rss("gudinnum")
vipshow=rss("vipshow")
sonum=rss("sonum")
paymod=rss("paymod")
retime=rss("retime")
showset=split(rss("showset"),",")
rss.close

sub down_nav()%>
	<table class=tableborder1 cellspacing=1 width="<%=Forum_body(12)%>" align=center>
		<tr>
			<th valign=middle height=25 width="100%">|<%
				dim rs1
				set rs1=server.createobject("adodb.recordset")
				sql="select class,classid from Dclass"
				rs1.open sql,conndown,1,3
				do while not rs1.eof
					%>&nbsp;&nbsp;<a href="z_down_index.asp?classid=<%=rs1("classid")%>" onMouseOver='ShowMenu(classlist<%=rs1("classid")%>,120)'><font color="ffffff"><%=rs1("class")%></font></a>&nbsp;&nbsp;|<%
					rs1.movenext
				loop
				rs1.close
				set rs1=nothing
			%></th>
		</tr>
	</table>
	<br>
<%end sub

'-------------�ϴ���ʽ--------------------------
upmod=0   '0:����������ϴ���1:lyfupload����ϴ�
'-----------------------------------------------

function downshow(sx)
	if sx="" or sx=1 or sx=0 or sx=" " then
		response.write "<font color=red>��������</font>"
	elseif sx=2 then
		response.write "<font color=red>��Ա����</font>"
	elseif sx=3 then
		response.write "<font color=red>VIP����</font>"
	elseif sx=4 then
		response.write "<font color=red>��Լ����</font>"
	elseif sx=5 then
		response.write "<font color=red>��������</font>"
	end if
end function

function soft(sx)
	if sx=0 then
		response.write "<font color=red>��ͨ���</font>"
	elseif sx=1 then
		response.write "<font color=red>RealPlay����</font>"
	elseif sx=2 then
		response.write "<font color=red>MediaPlay����</font>"
	elseif sx=3 then
		response.write "<font color=red>FLASH����</font>"
	end if
end function

function downtime(addt,downt)
	dim dt
	if downt=0 then
		downtime="����������"
	else
		dt=downt-(date()-addt)
		if date()-addt<=downt then
			downtime="��Ч��Ϊ"&downt&"�죻��ʣ "&dt&" ��"
		else
			downtime="��Ч��Ϊ"&downt&"�죻�ѹ���"
		end if
	end if
end function

function daydownload(dh,dd)
	'dh�������ش���
	'dd���������ش���
	dim last
	if dd<>"" and dd<>0 then
		last=dd-dh
		if dh=<dd then
			daydownload="ÿ����������Ϊ"&dd&"�Σ�����<font color=blue>"&last&"</font>��"
		else
			daydownload="ÿ����������Ϊ"&dd&"�Σ��ѽ�ֹ"
		end if
	else
		daydownload="ÿ�����ش���������"
	end if
end function 

function classnum(expression)
	dim rs_class,sql_class
	rs_class=conn.execute("select userclass from usertitle where usertitle='"&expression&"'")
	classnum=rs_class("userclass")
end function

sub down_manage()
	dim rsd,sqld,flagd
	sqld="select flag from [admin] where username='"&membername&"'"
	set rsd=conndown.execute(sqld)
	if not(rsd.eof and rsd.bof) then
		flagd=rsd("flag")%>
		<br>
		<table class=tableborder1 cellspacing=1 width="<%=Forum_body(12)%>" align=center>
			<tr>
				<th valign=middle height=25 width="100%">�������Ĺ���</th>
			</tr>
			<tr>
				<td class=tablebody1 align=center height=25><b><%=membername%></b> �����Խ��еĹ�����Ŀ������&nbsp;&nbsp;&nbsp;<%
					if flagd=1 then
						%>��<a href="z_down_softadd.asp">����ϴ����</a>&nbsp;&nbsp;&nbsp;��<a href="z_down_freeadd.asp">����������</a>&nbsp;&nbsp;&nbsp;��<a href="z_down_adminedit.asp">�޸�ɾ�����</a>&nbsp;&nbsp;&nbsp;��<a href="admin_index.asp">��Ա������Ŀ���������̨</a><%
					elseif flagd=2 then
						%>��<a href="z_down_softadd.asp">����ϴ����</a>&nbsp;&nbsp;&nbsp;��<a href="z_down_freeadd.asp">����������</a>&nbsp;&nbsp;&nbsp;��<a href="z_down_adminedit.asp">�޸�ɾ�����</a><%
					elseif flagd=3 then
						%>��<a href="z_down_softadd.asp">����ϴ����</a>&nbsp;&nbsp;&nbsp;��<a href="z_down_freeadd.asp">����������</a><%
					elseif flagd=4 then
						%>��<a href="z_down_freeadd.asp">����������</a><%
					end if
				%></td>
			</tr>
		</table>
	<%end if
end sub

'-------------------�����û�--------------------
sub down_payuser()
	dim rsp,sqlp
	sqlp="select * from [user] where username='"&membername&"' "
	set rsp=conndown.execute(sqlp)
	if not(rsp.eof and rsp.bof) then
		if rsp("apply")=2 then%>
			<br>
			<table class=tableborder1 cellspacing=1 width="<%=Forum_body(12)%>" align=center>
				<tr>
					<th valign=middle height=25 width="100%">�����û��������</th>
				</tr>
				<tr>
					<td class=tablebody1 align=center height=30>���Ѿ��ݽ��������ˣ�����������δ��������ȴ���</td>
				</tr>
			</table>
		<%elseif rsp("apply")=1 then%>
			<br>
			<table class=tableborder1 cellspacing=1 width="<%=Forum_body(12)%>" align=center>
				<tr>
					<th valign=middle height=25 width="100%">�����û��������</th>
				</tr>
				<%if session("payuser")<>membername then%>
					<form action=z_down_paylogin.asp?action=chk method=post>
					<tr>
						<td class=tablebody1 align=center height=30><b><%=membername%></b> ����������������븶���û�������棺��<input maxLength=20 name=password size=12 type=password>&nbsp;&nbsp;<input type=submit name=submit value="�� ��"></td>
					</tr>
					</form>
				<%elseif session("payuser")=membername then%>
					<tr>
						<td class=tablebody1 align=center height=30><b><%=membername%></b> &nbsp;&nbsp;[<a href=z_down_paycontrol.asp>�鿴��ϸ��Ϣ</a>]&nbsp;&nbsp;&nbsp;[<a href=z_down_paylogout.asp>�˳�</a>]</td>
					</tr>
				<%end if%>
			</table>
		<%end if
	elseif membername<>"" then%>
		<br>
		<table class=tableborder1 cellspacing=1 width="<%=Forum_body(12)%>" align=center>
			<tr>
				<th valign=middle height=25 width="100%">�����û�����</th>
			</tr>
			<tr>
				<td class=tablebody1 align=center height=30><a href=z_down_applypayuser.asp><font color=red>[�����Ϊ�����û�]</font></a></td>
			</tr>
		</table>
	<%end if
end sub

'--------------------���ѵ�ַ��ʾ
sub paybuy()
	dim rsu,foundsee,seenameid,seecount
	sql="select id from [user] where username='"&membername&"' and lockuser=1 and apply=1"
	set rsu=conndown.execute(sql)
	if rsu.eof and rsu.bof then
		userid=-1
	else
		userid=rsu("id")
	end if
	foundsee=false
	sql="select * from [download] where id="&sid
	set rsu=conndown.execute(sql)
	if rsu.eof and rsu.bof then
		response.write "����������ڣ�"
	else
		if lcase(trim(rsu("adduser")))=lcase(trim(membername)) or (flagdown>=1 and flagdown<=3) then
			foundsee=true
		else
			seenameid=split(rsu("buyuser"),",")
			seecount=ubound(seenameid)
			for i = 0 to seecount
				if strcomp(seenameid(i),userid)=0 then
					foundsee=true
					exit for
				end if
			next
		end if
		if foundsee then
			if rs("softsx")=1 then     
				play="z_down_play.asp?action=rm"     
				pimg="images/img_down/rm.gif"     
			elseif rs("softsx")=2 then    
				play="z_down_play.asp?action=mpeg"     
				pimg="images/img_down/mpeg.gif"
			elseif rs("softsx")=3 then
				play="z_down_play.asp?action=swf"     
				pimg="images/img_down/swf.gif"
			elseif rs("softsx")=4 then
				play="z_down_play.asp?action=mp3"     
				pimg="images/img_down/mp3.gif" 
			end if
			if rs("softsx")>0 then 
				%><b>���߲��ţ�</b><%
				if rs("filename")<>"" then
					%><A href="#" onClick="windowOpen('<%=play%>&id=<%=rs("id")%>&downid=1')"><IMG src="<%=pimg%>" BORDER=0 alt=���߲��ŵ�ַһ></A><%
				end if
				if rs("filename1")<>"" then
					%><A href="#" onClick="windowOpen('<%=play%>&id=<%=rs("id")%>&downid=2')"><IMG src="<%=pimg%>" BORDER=0 alt=���߲��ŵ�ַ��></A><%
				end if
				if rs("filename2")<>"" then
					%><A href="#" onClick="windowOpen('<%=play%>&id=<%=rs("id")%>&downid=3')"><IMG src="<%=pimg%>" BORDER=0 alt=���߲��ŵ�ַ��></A><%
				end if
				%><br><br><%
			end if
			if lcase(trim(rsu("adduser")))=lcase(trim(membername)) then
				%><font color=<%=forum_body(8)%>>��������ṩ�����븶Ǯ</font><br>[���������ش����]<br><%
			elseif flagdown>=1 and flagdown<=3 then
				%><font color=<%=forum_body(8)%>>�����������Ĺ��������븶Ǯ</font><br>[���������ش����]<br><%
			else
				%><font color=<%=forum_body(8)%>>���Ѿ����� <b><%=rsu("point")%></b> Ԫ�ֽ�</font><br>[���������ش����]<br><%
			end if
			if lcase(left(rs("filename"),6))="ftp://" then
				%><b>���ص�ַ1��</b><a href="<%=rs("filename")%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=���ص�ַһ></a><%
			else
				if rs("ldown")=1 and instr(rs("filename"),"$")>0 then
					%><b>���ص�ַ1��</b><a href="z_down_list.asp?id=<%=rs("id")%>&down=1&dpwd=<%=md5(downpwd&left(split(rs("filename"),"$")(1),6))%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=���ص�ַһ></a><%
				else
					%><b>���ص�ַ1��</b><a href="z_down_list.asp?id=<%=rs("id")%>&down=1&dpwd=<%=md5(downpwd)%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=���ص�ַһ></a><%
				end if
			end if
			if rs("filename1")<>"" then
				if lcase(left(rs("filename1"),6))="ftp://" then
					%><br><b>���ص�ַ2��</b><a href="<%=rs("filename1")%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=���ص�ַ��></a><%
				else
					%><br><b>���ص�ַ2��</b><a href="z_down_list.asp?id=<%=rs("id")%>&down=2&dpwd=<%=md5(downpwd)%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=���ص�ַ��></a><%
				end if
			end if
			if rs("filename2")<>"" then
				if lcase(left(rs("filename2"),6))="ftp://" then
					%><br><b>���ص�ַ3��</b><a href="<%=rs("filename2")%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=���ص�ַ��></a><%
				else
					%><br><b>���ص�ַ3��</b><a href="z_down_list.asp?id=<%=rs("id")%>&down=3&dpwd=<%=md5(downpwd)%>" ><IMG SRC="images/img_down/download.gif" BORDER=0 alt=���ص�ַ��></a><%
				end if
			end if         
		else
			if userid>0 then
				%><form action=z_down_paybuy.asp method=post><b>���ص�ַ��</b><font color=red>���ش������Ҫ <b><%=rsu("point")%></b> Ԫ�ֽ�</font><br>&nbsp;&nbsp;<font color=gray>����������ش���������������ĸ����������룬Ȼ��������</font><br>&nbsp;&nbsp;<input maxLength=20 name=sid size=12 type=hidden value="<%=rsu("id")%>"><input maxLength=20 name=password size=12 type=password>&nbsp;&nbsp;<input type=submit name=submit value="����"></form><%
			else
				%><b>���ص�ַ��</b><font color=red>���ش������Ҫ <b><%=rsu("point")%></b> Ԫ�ֽ�</font><%               
			end if
		end if
	end if
end sub%>
