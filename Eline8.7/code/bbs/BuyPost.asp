<!-- #include file="conn.asp" -->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/ubbcode.asp" -->
<%
	response.buffer=true
	dim rootid
	dim AnnounceID
	stats="��������"
	if BoardID="" or not isInteger(BoardID) or BoardID=0 then
		Errmsg=Errmsg+"<br>"+"<li>����İ����������ȷ�����Ǵ���Ч�����ӽ��롣"
		founderr=true
	else
		BoardID=clng(BoardID)
	end if
	if request("id")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
	elseif not isInteger(request("id")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
	else
		rootid=request("id")
	end if
	if request("replyID")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ָ��������ӡ�"
	elseif not isInteger(request("replyID")) then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ������Ӳ�����"
	else
		AnnounceID=request("replyID")
	end if
	dim FoundTable
	FoundTable=false
	if request("PostTable")<>"" then
	For i=0 to ubound(AllPostTable)
		if request("PostTable")=AllPostTable(i) then
			FoundTable=true
			exit for
		end if
	next
	else
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ��Ĳ�����"
	end if
	if not FoundTable then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ��Ĳ�����"
	end if
	if not founduser then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>���¼����в�����"
	end if

if founderr then
	call nav()
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call nav()
	call head_var(1,boarddepth,"","")
	call main()
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub main()
	dim re,test
	dim po,ii,jj
	dim reContent
	dim strContent
	dim PostBuyUser
	dim Title,Topic,Body
	dim rs2,sql2
	po=0
	ii=0
	jj=0
	dim usermoney
	set rs=conn.execute("select userWealth from [user] where userid="&userid)
	usermoney=rs(0)
	set rs=server.createobject("adodb.recordset")
	sql="select body,PostBuyUser,username,PostUserID,rootId,Topic from "&request("PostTable")&" where Announceid="&Announceid
	rs.open sql,conn,1,3
	if rs.eof and rs.bof then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��������Ӳ�����"
		exit sub
	else
		strContent=HTMLcode(rs(0))
		PostBuyUser=rs(1)
		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=True
		re.Pattern="\[USEMONEY=*([0-9]*)\]"
		Test=re.Test(strContent)
		if Test then
			re.Pattern="\[\/USEMONEY\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[USEMONEY=*([0-9]*)\]"
				strContent=re.replace(strContent, chr(1) & "USEMONEY=$1" & chr(2))
				re.Pattern="\[\/USEMONEY\]"
				strContent=re.replace(strContent, chr(1) & "/USEMONEY" & chr(2))
				re.Pattern="(\x01USEMONEY=*([0-9]*)\x02)(\x01\/USEMONEY\x02)"
				strContent=re.Replace(strContent,"")
				re.Pattern="(^.*)(\x01USEMONEY=*([0-9]*)\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)(.*)"
				po=re.Replace(strContent,"$3")
				if IsNumeric(po) then
					ii=int(po) 
				else
					ii=0
				end if
			end if
		else
			re.Pattern="\[USEMONEY=*([0-9]*),([0|1])\]"
			Test=re.Test(strContent)
			if Test then
				re.Pattern="\[\/USEMONEY\]"
				Test=re.Test(strContent)
				if Test then
					re.Pattern="\[USEMONEY=*([0-9]*),([0|1])\]"
					strContent=re.replace(strContent, chr(1) & "USEMONEY=$1,$2" & chr(2))
					re.Pattern="\[\/USEMONEY\]"
					strContent=re.replace(strContent, chr(1) & "/USEMONEY" & chr(2))
					re.Pattern="(\x01USEMONEY=*([0-9]*),([0|1])\x02)(\x01\/USEMONEY\x02)"
					strContent=re.Replace(strContent,"")
					re.Pattern="(^.*)(\x01USEMONEY=*([0-9]*),([0|1])\x02)(.[^\x01]*)(\x01\/USEMONEY\x02)(.*)"
					po=re.Replace(strContent,"$3")
					if IsNumeric(po) then
						ii=int(po) 
					else
						ii=0
					end if
					po=re.Replace(strContent,"$4")
					if IsNumeric(po) then
						jj=int(po) 
					else
						jj=0
					end if
				end if
			end if
		end if
		if membername=rs(2) then
			response.write "<script>alert('�Ǻǣ���Ҫ��Ǯ�����Լ������������');</script>"
		elseif usermoney>ii then
			if (not isnull(PostBuyUser)) and PostBuyUser<>"" then
				if instr("|"&PostBuyUser&"|","|"&membername&"|")>0 then
					response.write "<script>alert('�Ǻǣ����Ѿ��������ѽ��');</script>"
				else
					conn.execute("update [user] set userWealth=userWealth-"&ii&" where userid="&userid)
					conn.execute("update [user] set userWealth=userWealth+"&ii&" where userid="&rs(3))
					rs(1)=rs(1) & "|" & membername
					rs.update
					if rs("topic")="" or isnull(rs("topic")) then
						if len(rs("body"))>35 then
							Title=left(reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true),35)+"..."
						else
							Title=reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true)
						end if
					else
						if len(rs("Topic"))>35 then
							Title=left(htmlencode(rs("Topic")),35)+"..."
						else
							Title=htmlencode(rs("Topic"))
						end if
					end if
					if jj=1 then
						Topic="����"&Forum_info(0)&"���۵��������˹�����"
						body=rs("username")&"�����ã�"&chr(10)
						body=body & "����[color=blue]"&membername&"[/color]��[b]"&now&"[/b]����������"&Forum_info(0)&"���������[color=red]��"&Title&"��[/color]��"
						body=body & "�뵽���µ�ַ������������ݡ�"&chr(10)
						body=body & "[align=center][URL=dispbbs.asp?boardid="&boardid&"&id="&rs("rootid")&"&replyid="&announceid&"&skin=1][color=blue][b][u]�鿴��������[/u][/b][/color][/URL][/align]"
						set rs2=server.createobject("adodb.recordset")
						sql2="select * from [message]"      
						rs2.open sql2,conn,1,3      
						rs2.addnew      
						rs2("sender")=forum_info(0)      
						rs2("incept")=rs("username")
						rs2("title")=topic
						rs2("content")=body
						rs2("flag")=0 
						rs2("issend")=1
						rs2.update
						rs2.close
						set rs2=nothing
					end if
				end if
				response.write "<script>alert('����ɹ���');</script>"
			else
				conn.execute("update [user] set userWealth=userWealth-"&ii&" where userid="&userid)
				conn.execute("update [user] set userWealth=userWealth+"&ii&" where userid="&rs(3))
				rs(1)=membername
				rs.update
				if rs("topic")="" or isnull(rs("topic")) then
					if len(rs("body"))>35 then
						Title=left(reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true),35)+"..."
					else
						Title=reubbcode(htmlencode(replace(rs("body"),chr(10),"")),true)
					end if
				else
					if len(rs("Topic"))>35 then
						Title=left(htmlencode(rs("Topic")),35)+"..."
					else
						Title=htmlencode(rs("Topic"))
					end if
				end if
				if jj=1 then
					Topic="����"&Forum_info(0)&"���۵��������˹�����"
					body=rs("username")&"�����ã�"&chr(10)
					body=body & "����[color=blue]"&membername&"[/color]��[b]"&now&"[/b]����������"&Forum_info(0)&"���������[color=red]��"&Title&"��[/color]��"
					body=body & "�뵽���µ�ַ������������ݡ�"&chr(10)
					body=body & "[align=center][URL=dispbbs.asp?boardid="&boardid&"&id="&rs("rootid")&"&replyid="&announceid&"&skin=1][color=blue][b][u]�鿴��������[/u][/b][/color][/URL][/align]"
					set rs2=server.createobject("adodb.recordset")
					sql2="select * from [message]"      
					rs2.open sql2,conn,1,3      
					rs2.addnew      
					rs2("sender")=forum_info(0)      
					rs2("incept")=rs("username")
					rs2("title")=topic
					rs2("content")=body
					rs2("flag")=0 
					rs2("issend")=1
					rs2.update
					rs2.close
					set rs2=nothing
				end if
				response.write "<script>alert('����ɹ���');</script>"
			end if
		else
			response.write "<script>alert('����û��Ǯѽ��');</script>"
		end if
	end if
	rs.close
	set rs=nothing
	response.redirect "dispbbs.asp?boardid="&request("boardid")&"&id="&rootid&"&replyID="&AnnounceID&"&skin=1"
end sub
%>