<!--#include file=conn.asp-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_visual_const.asp" -->
<%dim menu,body
menu=request.querystring("menu")%>
<html>
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	if menu="" then
 		call main()
	elseif menu=1 then
	 	call defaultvalue()
	elseif menu=2 then
 		call feeman()
	elseif menu=3 then
 		call fixnew()
	elseif menu=4 then
 		call deltime()
	elseif menu=5 then
 		call deluser()
	elseif menu=6 then
 		call importitems()
	elseif menu=7 then
 		call delperiod()
	elseif menu=8 then
 		call photoman()
	elseif menu=9 then
 		call cleanbg()
	end if 	
end if%>

<%sub main() 
dim sql1,rs1,sql2,rs2 
sql1="select * from visual_config" 
set rs1=conn.execute(sql1) %>
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<form action=?menu=4 method=post>
	<tr>
		<th valign=middle colspan=2 height=23 align=left>ɾ��ָ���������¼�</th>
	</tr>
	<tr>
		<td valign=middle width=40% class=forumrow>ɾ��������ǰ���¼�(��д����)</td><td class=forumrow><input type=text name="TimeLimited" value=60 size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td>
	</tr>
	<tr>
		<td valign=middle width=40%  class=forumrow>�¼�����</td>
		<td class=forumrow><select name="deltype" size=1><option value="0">������־</option><option value="1">������־</option><option value="2">������־</option></select></td>
	</tr>
	</form>
	<form action=?menu=5 method=post>
	<tr>
		<th valign=middle colspan=2 height=23 align=left>ɾ��ĳ�û��������¼�</th>
	</tr>
	<tr>
		<td valign=middle width=40%  class=forumrow>�������û���</td>
		<td class=forumrow><input type=text name="username" size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td>
	</tr>
	<tr>
		<td valign=middle width=40%  class=forumrow>�¼�����</td>
		<td class=forumrow><select name="deltype" size=1><option value="0">������־</option><option value="1">������־</option><option value="2">������־</option></select></td>
	</tr>
	</form>
</table>
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
	<form action=?menu=7 method=post>
	<tr>
		<th valign=middle colspan=2 height=23 align=left>������ڵ���Ʒ</th>
	</tr>
	<tr>
		<td valign=middle width=40% class=forumrow>������ڶ�������Ʒ(��д����)</td><td class=forumrow><input type=text name="DelPeriod" value=60 size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td>
	</tr>
	</form>
</table>
<br>
<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
<tr><th align=center colspan=4 height=1 align="left"><b>������ƷĬ��ֵ����</b></th></tr> 
<form name="form" method="post" action="?menu=1"> 
<tr> 
	<td align=right width="25%" height="25" class=forumrow>Ĭ�ϵı�ۣ�</td> 
	<td align=left width="25%" height="25" class=forumrow><INPUT type=text name=price1 size="20" value="<%=rs1("price1")%>"></td> 
	<td align=right width="25%" height="25" class=forumrow>Ĭ�ϵ�VIP�ۣ�</td> 
	<td align=left width="25%" height="25" class=forumrow><INPUT type=text name=price2 size="20" value="<%=rs1("price2")%>"></td> 
</tr> 
<tr> 
	<td align=right width="25%" height="25" class=forumrow>Ĭ�ϵ���Ч�ڣ�</td> 
	<td align=left width="25%" height="25" class=forumrow><INPUT type=text name=period size="20" value="<%=rs1("period")%>"></td> 
	<td align=right width="25%" height="25" class=forumrow>Ĭ�ϵĿ�棺</td> 
	<td align=left width="25%" height="25" class=forumrow><INPUT type=text name=quantity size="20" value="<%=rs1("quantity")%>"></td> 
</tr> 
<tr> 
	<td align=right width="25%" height="25" class=forumrow>Ĭ�ϵ��޹��ߣ�</td> 
	<td align=left width="25%" height="25" class=forumrow><select size=1 name="flag"><option value=0<%if rs1("flag")=0 then%> selected<%end if%>>������Ʒ</option><option value=1<%if rs1("flag")=1 then%> selected<%end if%>>���л�Ա</option><option value=2<%if rs1("flag")=2 then%> selected<%end if%>>��������</option><option value=3<%if rs1("flag")=3 then%> selected<%end if%>>��������</option><option value=4<%if rs1("flag")=4 then%> selected<%end if%>>����Ա</option></select></td> 
	<td align=right width="25%" height="25" class=forumrow>Ĭ���Ƿ�VIPר����</td> 
	<td align=left width="25%" height="25" class=forumrow><input type=radio name="vip" value="0"<%if not rs1("vip") then%> checked<%end if%>>&nbsp;��&nbsp;&nbsp;<input type=radio name="vip" value="1"<%if rs1("vip") then%> checked<%end if%>>&nbsp;��</td> 
</tr> 
<tr> 
	<td align=center colspan=4 height="27" class=forumrow><INPUT type=submit value=�޸� name=submit></td> 
</tr> 
</form> 
</table> 
<br> 
<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
<tr><th align=center colspan=4 height=1 align="left"><b>������Ʒ�������ѹ���</b></th></tr> 
<form name="form" method="post" action="?menu=2"> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>������Ʒ��������(��׼�۸�İٷֱ�)��</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=send size="20" value="<%=rs1("send")%>"> %</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>������Ʒ��������(�ɽ��۸�İٷֱ�)��</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=sale size="20" value="<%=rs1("sale")%>"> %</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>������Ʒ����ͼ�(��׼�۸�İٷֱ�)��</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=salemin size="20" value="<%=rs1("salemin")%>"> %</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>�ͼ���Ʒ�ı�׼(���ڴ˼۸����Ʒ��Ϊ�ͼ���Ʒ)��</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=pricemin size="20" value="<%=rs1("pricemin")%>"> Ԫ</td> 
</tr> 
<tr> 
	<td align=center colspan=2 height="27" class=forumrow><INPUT type=submit value=�޸� name=submit></td> 
</tr> 
</form> 
</table> 
<br> 
<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
<tr><th align=center colspan=4 height=1 align="left"><b>������Ʒ�ķ��¹���</b></th></tr> 
<form name="form" method="post" action="?menu=3"> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>����һ����Ч�ڵķ���(���ܼ۸�/��Ч�ڵİٷֱ�)��</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=new1 size="20" value="<%=rs1("new1")%>"> %</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>��������Զ��Ч�ķ���(���ܼ۸�ı���)��</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=new0 size="20" value="<%=rs1("new0")%>"> ��</td> 
</tr> 
<tr> 
	<td align=center colspan=2 height="27" class=forumrow><INPUT type=submit value=�޸� name=submit></td> 
</tr> 
</form> 
</table> 
<br> 
<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder">
<tr><th align=center colspan=4 height=1 align="left"><b>�������</b></th></tr> 
<form name="form" method="post" action="?menu=8"> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>��ͨ��Ա����ļ۸�</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=photoprice1 size="20" value="<%=rs1("photoprice1")%>"> Ԫ/��</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>VIP��Ա����ļ۸�</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=photoprice2 size="20" value="<%=rs1("photoprice2")%>"> Ԫ/��</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>�����������Чʱ�䣺</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=photoperiod size="20" value="<%=rs1("photoperiod")%>"> ��</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>�����Ƭ��Ĭ��״̬��</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=radio name=photostatus value="0"<%if rs1("photostatus")=0 then%> checked<%end if%>> ����&nbsp;&nbsp;<INPUT type=radio name=photostatus value="1"<%if rs1("photostatus")=1 then%> checked<%end if%>> ����</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>��������Ӱ��������</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=photousercount size="20" value="<%=rs1("photousercount")%>"> ��</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>��������ϴ������Ĵ�С��</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=photobgsize size="20" value="<%=rs1("photobgsize")%>"> K</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>��ͨ��Ա�ϴ������ļ۸�</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=photoupprice1 size="20" value="<%=rs1("photoupprice1")%>"> Ԫ/K</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>VIP��Ա�ϴ������ļ۸�</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=photoupprice2 size="20" value="<%=rs1("photoupprice2")%>"> Ԫ/K</td> 
</tr> 
<tr> 
	<td align=right width="50%" height="25" class=forumrow>�����п���ʹ�õ��������(�Զ��ŷָ�)��</td> 
	<td align=left width="50%" height="25" class=forumrow><INPUT type=text name=accouterment size="20" value="<%=rs1("accouterment")%>"></td> 
</tr> 
<tr> 
	<td align=center colspan=2 height="27" class=forumrow><INPUT type=submit value=�޸� name=submit></td> 
</tr> 
</form> 
<form name="form" method="post" action="?menu=9"> 
<tr> 
	<td align=center colspan=2 height="27" class=forumrow><INPUT type=submit value=�������ñ��� name=submit></td> 
</tr> 
</form> 
</table> 
<br> 
<table width="95%" border="0" cellspacing="1" cellpadding="3" align=center class="tableBorder"> 
<tr><th align=center height=1 colspan="2" align="left">������Ʒ����</p</th></tr>
<tr> 
<td width="35%" class=forumrow><b>���������Ʒ</b><br>�����������뵼�����Ʒ���ݿ��ļ�����</td>
<form name="form1" method="post" action="?menu=6"> 
<td width="65%" class=forumrow height="45">����NewVisual<%=right("00"&month(now()),2)&right("00"&day(now()),2)&".mdb"%><br><INPUT type=text name=newvisualfile size="40">&nbsp; <INPUT type=submit value=ȷ�� name=submit></td> 
</form> 
</tr> 
<tr> 
<td width="35%" class=forumrow><b>������Ʒ����</b></td>
<form name="form1" method="get" action="z_admin_visual_items.asp">
<td width="65%" class=forumrow height="40"><INPUT type=submit value='�� �� �� ��' name=submit></td> 
</form>
</tr> 
<tr> 
<td width="35%" class=forumrow><b>������Ʒ�������</b></td>
<form name="form1" method="get" action="z_admin_visual_type.asp">
<td width="65%" class=forumrow height="40"><INPUT type=submit value='�� �� �� ��' name=submit></td> 
</form>
</tr> 
</table> 
<br> 
<%end sub

sub defaultvalue() 
	dim price1,price2,period,quantity,flag,vip,rs,sql
	if not isInteger(request.form("price1")) or not isInteger(request.form("price2")) or not isInteger(request.form("period")) or not isInteger(request.form("quantity")) or not isInteger(request.form("flag")) or not isInteger(request.form("vip")) then
		response.write "��������ݱ���Ϊ������"
		response.end 
	else 
		if request.form("price1")="" then 
			price1=100
		else 
			price1=int(request.form("price1")) 
		end if 
		if request.form("price2")="" then 
			price2=50 
		else 
			price2=int(request.form("price2")) 
		end if 
		if request.form("period")="" then 
			period=30 
		else 
			period=int(request.form("period")) 
		end if 
		if request.form("quantity")="" then 
			quantity=15 
		else 
			quantity=int(request.form("quantity")) 
		end if 
		if request.form("flag")="" then 
			flag=1 
		else 
			flag=int(request.form("flag")) 
		end if 
		if request.form("vip")="" then 
			vip=0
		else 
			vip=int(request.form("vip")) 
		end if 
	end if 
	set rs= server.createobject ("adodb.recordset") 
	sql = "select * from visual_config" 
	rs.Open sql,conn,1,3 
	if not(rs.eof and rs.bof) then 
		rs("price1") = price1 
		rs("price2") = price2 
		rs("period") = period
		rs("quantity") = quantity 
		rs("flag") = flag 
		if vip=0 then
			rs("vip")=false
		else
			rs("vip")=true
		end if
		rs.Update 
	end if 
	rs.close 
	set rs=nothing 
	response.write "�޸ĳɹ�!" 
end sub 

sub feeman() 
	dim send,sale,salemin,pricemin,rs,sql
	if not isInteger(request.form("send")) or not isInteger(request.form("sale")) or not isInteger(request.form("salemin")) or not isInteger(request.form("pricemin")) then
		response.write "��������ݱ���Ϊ������"
		response.end 
	else 
		if request.form("send")="" then 
			send=5
		else 
			send=int(request.form("send")) 
		end if 
		if request.form("sale")="" then 
			sale=5
		else 
			sale=int(request.form("sale")) 
		end if 
		if request.form("salemin")="" then 
			salemin=30 
		else 
			salemin=int(request.form("salemin")) 
		end if 
		if request.form("pricemin")="" then 
			pricemin=5 
		else 
			pricemin=int(request.form("pricemin")) 
		end if 
	end if 
	set rs= server.createobject ("adodb.recordset") 
	sql = "select * from visual_config" 
	rs.Open sql,conn,1,3 
	if not(rs.eof and rs.bof) then 
		rs("send") = send
		rs("sale") = sale
		rs("salemin") = salemin
		rs("pricemin") = pricemin
		rs.Update 
	end if 
	rs.close 
	set rs=nothing 
	response.write "�޸ĳɹ�!" 
end sub 

sub fixnew() 
	dim new1,new0,rs,sql
	if not isInteger(request.form("new1")) or not isInteger(request.form("new0")) then
		response.write "��������ݱ���Ϊ������"
		response.end 
	else 
		if request.form("new1")="" then 
			new1=50
		else 
			new1=int(request.form("new1")) 
		end if 
		if request.form("new0")="" then 
			new0=5
		else 
			new0=int(request.form("new0")) 
		end if 
	end if 
	set rs= server.createobject ("adodb.recordset") 
	sql = "select * from visual_config" 
	rs.Open sql,conn,1,3 
	if not(rs.eof and rs.bof) then 
		rs("new1") = new1
		rs("new0") = new0
		rs.Update 
	end if 
	rs.close 
	set rs=nothing 
	response.write "�޸ĳɹ�!" 
end sub 

sub deltime()
	dim dtime
	dtime=request("TimeLimited")
	if not isInteger(dtime) then
		response.write "�����������ҪΪ������"
	else
		conn.execute("delete from visual_events where datediff('d',dateandtime,now())>"&dtime&" and type="&request("deltype"))
		response.write "ɾ���ɹ���"
	end if
end sub 

sub delperiod()
	dim dtime
	dtime=request("DelPeriod")
	if not isInteger(dtime) then
		response.write "�����������ҪΪ������"
	else
		conn.execute("delete from visual_myitems where datediff('d',fixdate,now())-period>"&dtime&" and period>0")
		response.write "ɾ���ɹ���"
	end if
end sub 

sub deluser()
	dim username
	username=request("username")
	set rs=conn.execute("select userid from [user] where username='"&username&"'")
	if rs.eof and rs.bof then
		response.write "���û����Ǵ��ڡ�"
	else
		if request("deltype")=0 then
			conn.execute("delete from visual_events where (type=0 or type=2) and username='"&username&"'")
		elseif request("deltype")=1 then
			conn.execute("delete from visual_events where type=1 and fromuser='"&username&"'")
		elseif request("deltype")=2 then
			conn.execute("delete from visual_events where type=2 and fromuser='"&username&"'")
		end if
		response.write "ɾ���ɹ���"
	end if
	rs.close
end sub 

sub importitems()
	dim visualname,TCONN,icount,pcount
	visualname=Checkstr(trim(request("newvisualfile")))
	if visualname="" then
		ErrMsg=ErrMsg+"<Br>"+"<li>����д��������ݿ���"
		response.write Errmsg
		exit sub
	end if
	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(visualname)
	dim trs,nrs
	set trs=tconn.execute(" select * from visual_items ")
	sql="select * from visual_config" 
	set rs=conn.execute(sql)
	icount=0
	pcount=0
	do while not trs.eof
		if conn.execute("select top 1 id from visual_items where id="&trs("id")).eof then
			conn.execute("insert into visual_items (id,name,sex,type,content,price1,price2,period,quantity,flag,vip) values ("&trs("id")&",'"&trs("name")&"',"&trs("sex")&","&trs("type")&",'"&trs("content")&"',"&rs("price1")&","&rs("price2")&","&rs("period")&","&rs("quantity")&","&rs("flag")&","&rs("vip")&")")
			icount=icount+1
		else
			pcount=pcount+1
		end if
		trs.movenext
	loop
	trs.close
	set trs=nothing
	rs.close
	response.write "���ݵ���ɹ�(����"&icount&"������Ʒ������"&pcount&"��������Ʒ)��"
end sub

sub photoman() 
	dim photoprice1,photoprice2,photoperiod,photostatus,photousercount,photobgsize,photoupprice1,photoupprice2,accouterment,sql
	if not isInteger(request.form("photoprice1")) or not isInteger(request.form("photoprice2")) or not isInteger(request.form("photoperiod")) or not isInteger(request.form("photostatus")) or not isInteger(request.form("photousercount")) or not isInteger(request.form("photobgsize")) or not isInteger(request.form("photoupprice1")) or not isInteger(request.form("photoupprice2")) then
		response.write "��������ݱ���Ϊ������"
		response.end 
	else 
		if request.form("accouterment")="" then 
			response.write "��������ָ��һ�������п���ʹ�õ����"
			response.end 
		else
			accouterment=trim(request.form("accouterment"))
		end if 
		if request.form("photoprice1")="" then 
			photoprice1=20
		else 
			photoprice1=int(request.form("photoprice1")) 
		end if 
		if request.form("photoprice2")="" then 
			photoprice2=10
		else 
			photoprice2=int(request.form("photoprice2")) 
		end if 
		if request.form("photoperiod")="" then 
			photoperiod=5
		else 
			photoperiod=int(request.form("photoperiod")) 
		end if 
		if photoperiod<=0 then
			response.write "�����������Чʱ���������㣡"
			response.end 
		end if
		if request.form("photostatus")="" then 
			photostatus=1
		else 
			photostatus=int(request.form("photostatus")) 
		end if 
		if request.form("photousercount")="" then 
			photousercount=5
		else 
			photousercount=int(request.form("photousercount")) 
		end if 
		if photousercount<=0 then
			response.write "��������Ӱ��������������㣡"
			response.end 
		end if
		if request.form("photobgsize")="" then 
			photobgsize=100
		else 
			photobgsize=int(request.form("photobgsize")) 
		end if 
		if photobgsize<=0 then
			response.write "��������ϴ������Ĵ�С��"
			response.end 
		end if
		if request.form("photoupprice1")="" then 
			photoupprice1=2
		else 
			photoupprice1=int(request.form("photoupprice1")) 
		end if 
		if request.form("photoupprice2")="" then 
			photoupprice2=10
		else 
			photoupprice2=int(request.form("photoupprice2")) 
		end if 
	end if 
	set rs= server.createobject ("adodb.recordset") 
	sql = "select * from visual_config" 
	rs.Open sql,conn,1,3 
	if not(rs.eof and rs.bof) then 
		rs("photoprice1") = photoprice1
		rs("photoprice2") = photoprice2
		rs("photoperiod") = photoperiod
		rs("photostatus") = photostatus
		rs("photousercount") = photousercount
		rs("photobgsize") = photobgsize
		rs("photoupprice1") = photoupprice1
		rs("photoupprice2") = photoupprice2
		rs("accouterment")=accouterment
		rs.Update 
	end if 
	rs.close 
	set rs=nothing 
	response.write "�޸ĳɹ�!" 
end sub

sub cleanbg()
	dim objFSO
	dim uploadfolder
	dim uploadfiles
	dim upname
	dim upfilename
	dim selectfile
	dim delfile
	
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set uploadFolder=objFSO.GetFolder(Server.MapPath("UserVisual\"))
	Set uploadFiles=uploadFolder.Files
	i=0
	For Each Upname In uploadFiles
		delfile=0
		upfilename=PhotoPath&"/"&upname.name
		if instr(upname.name,"_")<=0 then
			set rs=conn.execute("select id from visual_photo where isUpload and Background='"&upname.name&"'")
			if rs.eof  then
				delfile=1
			end if
			rs.close
			set rs=nothing
		else
			selectfile=split(upname.name,"_")
			if lcase(selectfile(0))<>"photo" then
				delfile=1
			end if
		end if
		if delfile=1 then
			i=i+1
			objFSO.DeleteFile(Server.MapPath(upfilename))
			response.write upfilename&" �ѱ�ɾ����<br>"
		end if
	next
	response.write"��ɾ����"&i&"�������ñ����ļ���"
	set uploadFolder=nothing
	set uploadFiles=nothing
	set objFSO=nothing
end sub%> 
</html>