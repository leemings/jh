<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<!--#include file=z_down_function.asp-->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%dim filename,filename1,filename2,showname,classid,Nclassid,note,hot,system,size1,fromurl
dim order,downshow2,utime,daydown,softsx,money,piclink,addu,hots,hide,localdown,gudin
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	founderr=false
	if trim(request.form("txtfilename"))="" then
		founderr=true
		errmsg="<li>���ص������ַ����Ϊ��</li>"
	end if
	if trim(request.form("txtshowname"))="" then
		founderr=true
		errmsg=errmsg+"<li>��ʾ���Ʋ���Ϊ��</li>"
	end if
	if trim(request.form("Content"))="" then
		founderr=true
		errmsg=errmsg+"<li>�����鲻��Ϊ��</li>"
	end if
	if request.form("downshow")=5 and request.form("money")="" then
		founderr=true
		errmsg=errmsg+"<li>����۸���Ϊ��</li>"
	end if
	if trim(request.form("Classid"))="" then
		founderr=true
		errmsg=errmsg+"<li>�������಻��Ϊ��</li>"
	end if
	if trim(request.form("NClassid"))="" then
		founderr=true
		errmsg=errmsg+"<li>���С���಻��Ϊ��</li>"
	end if
	if trim(request.form("size1"))="" then
		founderr=true
		errmsg=errmsg+"<li>�����С����Ϊ��</li>"
	end if
	if request.form("daydown")="" then
		founderr=true
		errmsg=errmsg+"<li>�������ش������Ʋ���Ϊ�գ�</li>"
	else
		if not isInteger(request.form("daydown")) then
			founderr=true
			errmsg=errmsg+"<li>�������ش�������ֻ����������</li>"
		end if
	end if
	if founderr=false then
		filename=request("txtfilename")
		filename1=request("txtfilename1")
		filename2=request("txtfilename2")
		showname=htmlencode(request("txtshowname"))
		classid=request("classid")
		Nclassid=request("Nclassid")
		note=Checkstr(trim(request("Content")))
		note=replace(note,"''","'")
		hot=htmlencode(request("hot"))
		system=htmlencode(request("system"))
		size1=htmlencode(request("size1"))
		fromurl=htmlencode(request("fromurl"))
		order=htmlencode(request("order"))
		downshow2=request("downshow")
		utime=request("usetime")
		daydown=request("daydown")
		softsx=request("softsx")
		money=request("money")
		if money="" then
			money=0
		end if
		piclink=request("pic")
		addu=membername
		if request.form("hots")="on" then
			hots=1
		else
			hots=0
		end if
		if request.form("hide")="on" then
			hide=1
		else
			hide=0
		end if
		if request.form("localdown")="on" then
			localdown=1
		else
			localdown=0
		end if
		if request.form("gudin")="on" then
			gudin=1
			hots=1
		else
			gudin=0
		end if
		set rs=server.createobject("adodb.recordset")
		if request("action")="add" then
			call newsoft()
		elseif request("action")="edit" then
			call editsoft()
		end if%>
		<table width="95%" border="0" cellpadding="3" cellspacing="1" align=center class=tableborder>
			<tr>
				<th height="25"><%
					if request("action")="add" then
						%>���<%
					else
						%>�޸�<%
					end if
				%>����ɹ�</th>
			</tr>
			<tr>
				<td class=forumrow><br>���ص�ַ1Ϊ��<%=filename%><br>���ص�ַ2Ϊ��<%=filename1%><br>���ص�ַ3Ϊ��<%=filename2%><br>��ʾ����Ϊ��<%=showname%></p>�����Խ�����������</td>
			</tr>
		</table>
	<%else
		response.write "�������µ�ԭ���ܱ������ݣ�"
		response.write errmsg
	end if
end if%>
</body>
<%set rs=nothing

sub newsoft()
	dim mess
	sql="select * from download where (id is null)" 
	rs.open sql,conndown,1,3
	rs.addnew
	rs("filename")=filename
	if filename1<>"" then
		rs("filename1")=filename1
	end if
	if filename2<>"" then
		rs("filename2")=filename2
	end if
	rs("showname")=showname
	rs("downshow")=downshow2
	rs("classid")=classid
	rs("Nclassid")=Nclassid
	rs("fromurl")=fromurl
	rs("note")=note
	rs("system")=system
	rs("hot")=hot
	rs("hots")=hots
	rs("stop")=hide
	rs("size")=size1
	rs("orders")=order
	rs("softsx")=softsx
	rs("daydown")=daydown
	rs("lasthits")=date()
	rs("dateandtime")=date()
	rs("dayhits")=0
	rs("weekhits")=0
	rs("hits")=0
	if piclink<>"" then
		rs("pic")=piclink
	end if
	rs("adduser")=addu
	rs("point")=money
	rs("buyuser")=0
	rs("usetime")=utime
	rs("gudin")=gudin
	rs("ldown")=localdown
	rs.update
	rs.close
	sql="select * from events"
	rs.open sql,conndown,1,3
	mess=""&membername&"����ˡ�"&showname&"��"	
	rs.addnew
	rs("event")=mess
	rs("addtime")=date()
	rs.update
	rs.close
end sub

sub editsoft()
	dim mess
	sql="select * from download where id="&request("id")
	rs.open sql,conndown,1,3
	rs("filename")=filename
	if filename1<>"" then
		rs("filename1")=filename1
	else
		rs("filename1")=NULL
	end if
	if filename2<>"" then
		rs("filename2")=filename2
	else
		rs("filename2")=NULL
	end if
	rs("showname")=showname
	rs("classid")=classid
	rs("Nclassid")=Nclassid
	rs("downshow")=downshow2
	rs("fromurl")=fromurl
	rs("note")=note
	rs("system")=system
	rs("hot")=hot
	rs("hots")=hots
	rs("stop")=hide
	rs("softsx")=softsx
	rs("size")=size1
	rs("orders")=order
	if piclink<>"" then
		rs("pic")=piclink
	else
		rs("pic")=NULL
	end if
	rs("lasthits")=date()
	rs("daydown")=daydown
	rs("point")=money
	rs("gudin")=gudin
	rs("usetime")=utime
	rs("ldown")=localdown
	rs.update
	rs.close
	sql="select * from events"
	rs.open sql,conndown,1,3
	mess=""&membername&"�޸��ˡ�"&showname&"��"	
	rs.addnew
	rs("event")=mess
	rs("addtime")=date()
	rs.update
	rs.close
end sub%>
