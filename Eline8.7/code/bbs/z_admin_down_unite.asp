<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file="z_down_conn.asp"-->
<html>
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<%dim str
if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else
	if Request("action") = "unite" then
		call unite()
	elseif Request("action") = "npwd" then
		call nodownpwd()
	elseif Request("action") = "deltime" then
		call deltime()
	elseif Request("action") = "deluser" then
		call deluser()
	elseif Request("action") = "dpath" then
		call dpath()
	elseif Request("action") = "show" then
		call show()
	elseif Request("action") = "show2" then
		call show2()
	elseif Request("action") = "save" then
		call save()
	elseif Request("action") = "vipsave" then
		call vipsave()
	elseif Request("action") = "upset" then
		call upset()
	elseif Request("action") = "retime" then
		call retime()
	elseif Request("action") = "renfile" then
		call renfile()
	else
		call boardinfo()
	end if
end if
conndown.close
set conndown=nothing%>
</html>

<%sub boardinfo()%>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=unite method=post>
		<tr>
			<th>�ϲ��ӷ�������</th>
		</tr>
		<tr>
			<td class=forumrow>ע����������棬��������Ŀǰ���е���̳�б������Խ�һ���ӷ��������һ���ӷ�����кϲ����ϲ��������ӷ����������ת��ϲ�����ӷ��ࡣ</td>
		</tr>
		<tr>
			<td class=forumrow><%
				dim NClassString
				set rs = server.CreateObject ("Adodb.recordset")
				sql="select * from DNclass"
				rs.open sql,conndown,1,1
				if rs.eof and rs.bof then
					NClassString="û���ӷ���"
				else
					NClassString=""
					do while not rs.eof
						NClassString=NClassString&"<option value="&rs("nclassid")&">"&rs("nclass")&"</option>"
						rs.movenext
					loop
					response.write "���ӷ���&nbsp;"
					response.write "<select name=oldboard size=1>"
					response.write NClassString
					response.write "</select>&nbsp;&nbsp;"
					response.write "�ϲ���&nbsp;"
					response.write "<select name=newboard size=1>"
					response.write NClassString
					response.write "</select>&nbsp;&nbsp;"
				end if
				rs.close
				set rs=nothing
				response.write "<input type=submit name=Submit value=ִ��>"
			%></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<tr>
			<th>�� �� ��</th>
		</tr>
		<tr>
			<td class=forumrow>���������������Ч�ķ�ֹ���˴��ⲿ�����������ص�ַ���붨���޸�<br>����Ч�ķ������ֶλ��Ƕ��ڸ������ص�ַ</td>
		</tr>
		<%set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from notdown"
		rs.open sql,conndown,1,1	%>
		<form action=?action=npwd method=post>
		<tr>
			<td class=forumrow><b>���������룺</b><input maxLength=20 name=downpwd size=15 value="<%=rs("nodown")%>">&nbsp;&nbsp;<input type=submit name=submit value="�� ��"></td>
		</tr>
		</form>
		<tr>
			<td class=forumrow>�����޸�����������Ŀ¼�����ƣ�������Ч�ķ�ֹ���˵����������<br><font color="#FF0000">˵����������Ŀռ�֧��FSO���ܣ���ô����д���ļ�����������޸ĺ󼴿����Ŀ¼�����޸�<br>������Ŀռ䲻֧��FSO����ô�����޸���������ֹ��޸�ԭ����Ŀ¼����ʹ������޸ĵ�Ŀ¼��һ��<br></font><font color="#FF0000">��д����</font>��������д���ļ��б����index.asp�ļ���ͬһ��Ŀ¼�����Ҳ�Ҫ��д��Ŀ¼����<br><font color="#FF0000">��ȷ��д��</font>��uploadimages</td>
		</tr>
		<form action=?action=dpath method=post>
		<tr>
			<td class=forumrow><b>������Ŀ¼��</b><input maxLength=20 name=downpath size=15 value="<%=rs("downpath")%>">&nbsp;&nbsp;<input type=submit name=submit value="�� ��"></td>
		</tr>
		</form>
		<%rs.close
		set rs=nothing%>
		<form action=?action=renfile method=post>
		<tr>
			<td class=forumrow><b>�ļ����������</b>&nbsp;<input type=submit name=submit value="�� ��"></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=deltime method=post>
		<tr>
			<th valign=middle colspan=2 height=23 align=left>ɾ��ָ���������¼�</th>
		</tr>
		<tr>
			<td valign=middle width=40% class=forumrow>ɾ��������ǰ���¼�(��д����)</td><td class=forumrow><input name="TimeLimited" value=0 size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td>
		</tr>
		<tr>
			<td valign=middle width=40%  class=forumrow>�¼�����</td>
			<td class=forumrow><select name="deltype" size=1><option value="1">����Ա������־</option><option value="2">�����û�������־</option><option value="3">�����û�������־</option></select></td>
		</tr>
		</form>
		<form action=?action=deluser method=post>
		<tr>
			<th valign=middle colspan=2 height=23 align=left>ɾ��ĳ�û��������¼�</th>
		</tr>
		<tr>
			<td valign=middle width=40%  class=forumrow>�������û���</td>
			<td class=forumrow><input type=text name="username" size=30>&nbsp;<input type=submit name="submit" value="�� ��"></td>
		</tr>
		<tr>
			<td valign=middle width=40%  class=forumrow>�¼�����</td>
			<td class=forumrow><select name="deltype" size=1><option value="1">�����û�������־</option><option value="2">�����û�������־</option></select></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<%dim ShowSet,upsizemax,uppicsizemax
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from showpage"
		rs.open sql,conndown,1,1
		showset=split(rs("showset"),",")
		upsizemax=int(int(rs("upsizemax"))/1024)
		uppicsizemax=int(int(rs("uppicsizemax"))/1024)%>
		<form action=?action=upset method=post>
		<tr>
			<th valign=middle height=23 align=left width="100%" colspan="2">�ϴ������趨</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="18%">�ϴ�<b>�ļ�</b>����(����|�ָ�)��</td> 
	  	<td valign=middle class=forumrow align=left width="57%"><input type=text name="uptype" size=74 value="<%=rs("uptype")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="18%">�ϴ�<b>�ļ���С</b>�趨��</td>
			<td valign=middle class=forumrow align=left width="57%"><input type=text name="upsizemin" value="<%=rs("upsizemin")%>" size=13><b>K</b>&nbsp; &lt; �ϴ��ļ���С &lt; <input type=text name="upsizemax" value="<%=upsizemax%>" size=13><b>K</b></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="18%">�ϴ�<b>���ͼƬ</b>����(����|�ָ�)��</td>
			<td valign=middle class=forumrow align=left width="57%"><input type=text name="uppictype" size=74 value="<%=rs("uppictype")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="18%">�ϴ�<b>���ͼƬ��С</b>�趨��</td>
			<td valign=middle class=forumrow align=left width="57%"><input type=text name="uppicsizemin" value="<%=rs("uppicsizemin")%>" size=13><b>K</b>&nbsp;&lt; �ϴ����ͼƬ��С &lt; <input type=text name="uppicsizemax" value="<%=uppicsizemax%>" size=13><b>K</b></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=center colspan=2><input type=submit name="submit" value="�� ��"></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=vipsave method=post>
		<tr>
			<th valign=middle height=23 align=left width="100%">VIP/��ͨ���л�</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="100%"><font color="#FF0000">��Ҫ˵����</font>���������̳û�а�װVIP�������ô��ֻ��ʹ����ͨ�棬���ѡ��VIP�棬������ִ���<br>���������̳��װ��VIP�����������ʹ��VIP�汾<br>����VIP�汾�������ͨ���һ��VIP��Ա���أ�������ܿ���ʹ���ؽ�������̳VIP��Ա<br>ע�⣺<font color="#FF0000">�������ʹ����VIP��һ��ʱ����л�����ͨ�棬����̳�����еĽ���VIP���ص��������ɽ��޻�Ա����</font>������ϸ���Ǻ����</td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="100%" align="center">VIP��Ա��&nbsp;<input type=radio name="vipshow" value="2" <%if rs("vipshow")=2 then%>checked<%end if%>>&nbsp;&nbsp;&nbsp;&nbsp;��ͨ�汾&nbsp;<input type=radio name="vipshow" value="1" <%if rs("vipshow")=1 then%>checked<%end if%>></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="100%" align="center"><input type=submit name="submit" value="�� ��"></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=retime method=post>
		<tr>
			<th valign=middle height=23 align=left width="100%">ҳ��ˢ��ʱ���޶�</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="100%">˵�����û���<b>�����ϸ����ҳ��</b>ˢ��ʱ�䲻�ó������趨ʱ�䣬�趨ʱ�������Ӱ���û������ʱ��̫�̿��ܻ���������ˢ�¡�����ϸ�趨</td> 
		</tr>
		<tr>
			<td valign=middle class=forumrow align=center width="100%"><input type=text name="retime" value="<%=rs("retime")%>" size=13><b>��</b>&nbsp;&nbsp;&nbsp;&nbsp;<input type=submit name="submit" value="�� ��"></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=show method=post>
		<tr>
			<th valign=middle colspan=4 height=23 align=left width="100%">ҳ����ʾ�����趨</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="25%">��������TOP����</td>
			<td valign=middle class=forumrow width="175"><input type=text name="daytop" size=15 value="<%=rs("daytop")%>"></td>
			<td valign=middle class=forumrow width="163">��������TOP����</td>
			<td valign=middle class=forumrow width="189"><input type=text name="weektop" size=15 value="<%=rs("weektop")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="25%">�ۼ�����TOP����</td>
			<td valign=middle class=forumrow width="175"><input type=text name="alltop" size=15 value="<%=rs("alltop")%>"></td>
			<td valign=middle class=forumrow width="163">�����Ƽ�TOP����</td>
			<td valign=middle class=forumrow width="189"><input type=text name="hotnum" size=15 value="<%=rs("hotnum")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="25%">�̶��Ƽ�TOP����</td>
			<td valign=middle class=forumrow width="175"><input type=text name="gudinnum" size=15 value="<%=rs("gudinnum")%>"></td>
			<td valign=middle class=forumrow width="163">������ҽ��ÿҳ��ʾ��</td>
			<td valign=middle class=forumrow width="189"><input type=text name="sonum" size=15 value="<%=rs("sonum")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="25%">����б���ÿҳ��Ŀ��</td>
			<td valign=middle class=forumrow width="175"><input type=text name="listnum" size=15 value="<%=rs("listnum")%>"></td>
			<td valign=middle class=forumrow width="163">�½�����б�����</td>
			<td valign=middle class=forumrow width="189"><input type=text name="newnum" size=15 value="<%=rs("newnum")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="25%">�޸�ɾ��ҳ��ÿҳ�����Ŀ��</td>
			<td valign=middle class=forumrow width="175"><input type=text name="editnum" size=15 value="<%=rs("editnum")%>"></td>
			<td valign=middle class=forumrow width="163">�����¼�ҳ��ÿҳ�¼���Ŀ��</td>
			<td valign=middle class=forumrow width="189"><input type=text name="eventnum" size=15 value="<%=rs("eventnum")%>"></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow colspan=4 align=center><input type=submit name="submit" value="�� ��"></td>
		</tr>
		</form>
		<%rs.close
		set rs=nothing%>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=show2 method=post>
		<tr>
			<th valign=middle colspan=6 height=23 align=left width="100%">������ʾ�趨</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="20%" align="center"></td>
			<td valign=middle class=forumrow width="16%" align="center">�̶��Ƽ�</td>
			<td valign=middle class=forumrow width="16%" align="center">�����Ƽ�</td>
			<td valign=middle class=forumrow width="16%" align="center">��������</td>
			<td valign=middle class=forumrow width="16%" align="center">��������</td>
			<td valign=middle class=forumrow width="16%" align="center">�ۼ�����</td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="20%" nowrap>����������ҳ(z_down_default.asp)</td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="allhot1"   value="1" <%if showset(0)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="hot1" value="1" <%if showset(1)=1   then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="daytop1" value="1" <%if showset(2)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="weektop1"   value="1" <%if showset(3)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="alltop1" value="1" <%if showset(4)=1   then%>checked<%end if%>></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="20%" nowrap>�������ҳ��(z_down_index.asp)</td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="allhot2"   value="1" <%if showset(5)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="hot2" value="1" <%if showset(6)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="daytop2" value="1" <%if showset(7)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="weektop2"   value="1" <%if showset(8)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="alltop2" value="1" <%if showset(9)=1   then%>checked<%end if%>></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow width="20%" nowrap>�����ϸ����ҳ��(z_down_list.asp)</td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="allhot3" value="1" <%if showset(10)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="hot3" value="1" <%if showset(11)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="daytop3" value="1" <%if showset(12)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="weektop3" value="1" <%if showset(13)=1 then%>checked<%end if%>></td>
			<td valign=middle class=forumrow width="16%" align="center"><input type="checkbox" name="alltop3" value="1" <%if showset(14)=1 then%>checked<%end if%>></td>
		</tr>
		<tr>
			<td valign=middle class=forumrow colspan=6 align=center><input type=submit name="submit" value="�� ��"></td>
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<tr>
			<th valign=middle colspan=2 height=23 align=left width="100%">�ϴ��ļ�������ҪFSO֧�֣�</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="25%"> ˵���� </td>
			<td valign=middle class=forumrow align=left width="75%">��Ŀռ����Ҫ֧��FSO����ʹ�ñ�����</td>
		</tr> 
		<tr>
			<form name="form1" method="get" action="z_admin_down_uploadlist.asp">
			<td valign=middle class=forumrow colspan=2 align=center><INPUT type=submit value='�� �� �� ��' name=submit></td>
			</form>
		</tr> 
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=?action=save method=post>
		<tr>
			<th valign=middle height=23 align=left width="100%">�������ݿ����</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="100%">ָ�<Input type="text" name="SQL_Statement" Size=60> <Input type="Submit" Value="�ͳ�"></td> 
		</tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=z_admin_down_rb.asp?action=backupdata method=post>
		<tr>
			<th valign=middle colspan=2 height=23 align=left width="100%">�������ݿⱸ�ݣ���ҪFSO֧�֣�</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="25%">��ǰ���ݿ�·����&nbsp;</td>
			<td valign=middle class=forumrow align=left width="75%"><input type=text size=24 name=DBpath value="<%=dbdown%>"> ����ȷ��д����ǰʹ�õ����ݿ�·����</td>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="25%">�������ݿ�Ŀ¼��</td>
			<td valign=middle class=forumrow align=left width="75%"><input type=text size=24 name=bkfolder value=Databackup> ���Ŀ¼�����ڣ������Զ�������</td>
		</tr> 
		<tr>
			<td valign=middle class=forumrow align=left width="25%">�������ݿ����ƣ�</td>
			<td valign=middle class=forumrow align=left width="75%"><input type=text size=24 name=bkDBname value=downbackup.mdb> �������Ŀ¼�и��ļ��������ǣ����û�У������Զ�������</td>
		</tr> 
	  <tr>
	  	<td valign=middle class=forumrow colspan=2 align=center><input type=submit value="��ʼ����"></td>
	  </tr>
		</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=z_admin_down_rb.asp?action=restoredata method=post>
		<tr>
			<th valign=middle colspan=2 height=23 align=left width="100%">�������ݿ�ָ�����ҪFSO֧�֣�</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="25%"> �������ݿ�·����</td>
			<td valign=middle class=forumrow align=left width="75%"><input type=text size=37 name=dbpath value="Databackup/downbackup.mdb">&nbsp;����ļ������ڣ������ָܻ�</td>
		</tr> 
		<tr>
			<td valign=middle class=forumrow align=left width="25%"> Ŀ�����ݿ�·����</td>
			<td valign=middle class=forumrow align=left width="75%"><input type=text size=37 name=backpath value="<%=dbdown%>">&nbsp;��д����ǰʹ�õ����ݿ�·����</td>
		</tr> 
	  <tr><td valign=middle class=forumrow colspan=2 align=center><input type=submit value="��ʼ�ָ�"></td></tr>
	</form>
	</table>
	<br>
	<table width="95%" border="0" cellspacing="1" cellpadding="3"  align=center class="tableBorder">
		<form action=z_admin_down_rb.asp?action=comp method=post>
		<tr>
			<th valign=middle colspan=2 height=23 align=left width="100%">ѹ���������ݿ⣨��ҪFSO֧�֣�</th>
		</tr>
		<tr>
			<td valign=middle class=forumrow align=left width="100%" colspan="2"> <b>ע�⣺</b><br>�������ݿ��������·��,�����������ݿ����ƣ�����ʹ�������ݿⲻ��ѹ������ѡ�񱸷����ݿ����ѹ��������</td>
		</tr> 
		<tr>
			<td valign=middle class=forumrow align=left width="25%">ѹ�����ݿ⣺</td>
			<td valign=middle class=forumrow align=left width="75%"><input type="text" name="dbpath" value=Databackup/downbackup.mdb size="39">&nbsp;<input type="submit" value="��ʼѹ��"></td>
		</tr> 
	  <tr>
	  	<td valign=middle class=forumrow align=left width="100%" colspan="2"><input type="checkbox" name="boolIs97" value="True"> ���ʹ�� Access 97 ���ݿ���ѡ��(Ĭ��Ϊ Access 2000 ���ݿ�)</td>
		</tr>
		</form>
	</table>
	<br>
<%end sub

sub unite()
	dim newboard,newclassid,oldboard
	if cint(request("newboard"))=cint(request("oldboard")) then
		response.write "�벻Ҫ����ͬ�ӷ����ڽ��в�����"
	else
		newboard=cint(request("newboard"))
		oldboard=cint(request("oldboard"))
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from DNclass where nclassid="&newboard&""
		rs.open sql,conndown,1,1
		newclassid=rs("classid")
		rs.close
		set rs=nothing
		conndown.execute("update download set nclassid="&newboard&",classid="&newclassid&" where nclassid="&oldboard)
		response.write "�ϲ��ɹ����Ѿ������ϲ��ӷ�����������ת�������ϲ��ӷ��ࡣ"
	end if
end sub

sub nodownpwd()
	dim nopwd
	if request("downpwd")="" then
		response.write "���������벻��Ϊ�ա�"
	else
		nopwd=request("downpwd")
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from notdown"
		rs.open sql,conndown,1,3
		rs("nodown")=nopwd
		rs.update
		rs.close
		sql="select * from events"
		rs.open sql,conndown,1,3
		rs.addnew
		rs("event")=membername&"�޸��˷��������룡"
		rs("addtime")=date()
		rs.update
		rs.close
		set rs=nothing
		response.write "�޸ĳɹ���"
	end if
end sub

sub dpath()
	dim downpath,downpathold,fsdir,fsdir1,fs,folder1,folder2
	downpath=request("downpath")
	if InStr(downpath,"/")<>0 or downpath="" then
		response.write("��������ļ�Ŀ¼������ȷ������ϸ�Ķ�˵��")
	else
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from notdown"
		rs.open sql,conndown,1,3
		downpathold=rs("downpath")
		if downpath=downpathold then
			response.write("��������ļ�Ŀ¼����ԭ�ļ�������ͬ���޷���������������ϸ�Ķ�˵��")
		else
			rs("downpath")=downpath
			rs.update
			response.write("���ݿ�������Ŀ¼���޸ĳɹ���<br>")
			fsdir=Server.MapPath("index.asp")	
			fsdir1=replace(fsdir,"\index.asp","\")
			set fs=createobject("scripting.filesystemobject") 
			folder1=fsdir1&"\"&downpathold
			folder2=fsdir1&"\"&downpath
			If not fs.folderexists(folder1) then
				response.write fsDir&"ԭ�ļ��в�����,�����������ļ�������<br>"
			elseIf fs.folderexists(folder2) then
				response.write "Ŀ���ļ���<font color=red>�Ѿ�����</font>,��ѡ�������ļ���<br>"
			else 
				fs.MoveFolder folder1,folder2
				If not fs.folderexists(folder1) and fs.folderexists(folder2) then
					response.write "�ļ�Ŀ¼���޸�<font color=red>�ɹ���</font><br>"
				Else
					response.write "�ļ�Ŀ¼���޸�<font color=red>ʧ�ܣ�</font><br>"
				End If
			end if
		end if
		rs.close
		set rs=nothing
	end if
end sub

sub deltime()
	dim dtime,deltype
	dtime=request("TimeLimited")
	if request("deltype")=1 then
		deltype="events"
	elseif request("deltype")=2 then
		deltype="downevent"
	elseif request("deltype")=3 then
		deltype="userevent"
	end if 
	if not isInteger(dtime) then
		response.write "�����������ҪΪ������"
	else
		conndown.execute("delete from "&deltype&" where addtime>now()-"&dtime)
		response.write "ɾ���ɹ���"
	end if
end sub 

sub deluser()
	dim username,deltype
	username=request("username")
	if request("deltype")=1 then
		deltype="downevent"
	elseif request("deltype")=2 then
		deltype="userevent"
	end if 
	set rs=conndown.execute("select id from [user] where username='"&username&"'")
	if rs.eof and rs.bof then
		response.write "���û����Ǹ����û���"
	else
		conndown.execute("delete from "&deltype&" where username='"&username&"'")
		response.write "ɾ���ɹ���"
	end if
	rs.close
end sub 

sub upset()
	dim upsizemax,uppicsizemax
	if request("uptype")="" or request("upsizemin")="" or request("upsizemax")="" or request("uppictype")="" or request("uppicsizemin")="" or request("uppicsizemax")="" then
		response.write "��������ݲ���Ϊ�գ�"
	elseif not isInteger(request("upsizemin")) or  not isInteger(request("upsizemax")) or not isInteger(request("uppicsizemin")) or  not isInteger(request("uppicsizemax")) then
		response.write "������ϴ��ļ���С����Ϊ������"
	else
		upsizemax= int(request("upsizemax")*1024)
		uppicsizemax=int(request("uppicsizemax")*1024)
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from showpage"
		rs.open sql,conndown,1,3
		rs("uptype")= request("uptype")
		rs("upsizemin")=int(request("upsizemin"))
		rs("upsizemax")=upsizemax
		rs("uppictype")= request("uppictype")
		rs("uppicsizemin")=int(request("uppicsizemin"))
		rs("uppicsizemax")=uppicsizemax
		rs.update
		rs.close
		set rs=nothing
		response.write "�޸ĳɹ���"
	end if
end sub

sub vipsave()
	if request("vipshow")=""  then
		response.write "��������ݲ���Ϊ�գ�"
	else
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from showpage"
		rs.open sql,conndown,1,3
		if request("vipshow")=2 or (request("vipshow")=1 and rs("vipshow")=1) then
			rs("vipshow")=request("vipshow")
			rs.update
			response.write("�޸ĳɹ���")
		elseif request("vipshow")=1 and rs("vipshow")=2 then
			conndown.execute"update [download] set downshow=2 where downshow=3" 
			rs("vipshow")=request("vipshow")
			rs.update
			response.write("�޸ĳɹ���")
		end if
		rs.close
		set rs=nothing
	end if
end sub

sub retime()
	if request("retime")="" then
		response.write "��������ݲ���Ϊ�գ�"
	elseif not isInteger(request("retime")) then
		response.write "�����ʱ�����Ϊ������"
	else
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from showpage"
		rs.open sql,conndown,1,3
		rs("retime")=request("retime")
		rs.update
		rs.close
		set rs=nothing
		response.write "�޸ĳɹ���"
	end if
end sub

sub show()
	if request("daytop")="" or request("weektop")="" or request("alltop")="" or request("listnum")="" or request("editnum")="" or request("eventnum")="" or request("hotnum")="" or request("newnum")="" or request("gudinnum")="" then
		response.write "��������ݲ���Ϊ�գ�"
	elseif not isInteger(request("daytop")) or not isInteger(request("weektop")) or not isInteger(request("alltop")) or not isInteger(request("listnum")) or not isInteger(request("editnum")) or not isInteger(request("eventnum")) or not isInteger(request("newnum")) or not isInteger(request("hotnum")) or not isInteger(request("gudinnum")) then
		response.write "��������ݱ���Ϊ������"
	else
		set rs = server.CreateObject ("Adodb.recordset")
		sql="select * from showpage"
		rs.open sql,conndown,1,3
		rs("daytop")=request("daytop")
		rs("weektop")=request("weektop")
		rs("alltop")=request("alltop")
		rs("listnum")=request("listnum")
		rs("editnum")=request("editnum")
		rs("eventnum")=request("eventnum")
		rs("newnum")=request("newnum")
		rs("hotnum")=request("hotnum")
		rs("gudinnum")=request("gudinnum")
		rs("sonum")=request("sonum")
		rs.update
		rs.close
		response.write "�޸ĳɹ���"
	end if
end sub

sub show2()
	dim allhot1,allhot2,allhot3,hot1,hot2,hot3,daytop1,daytop2,daytop3,weektop1,weektop2,weektop3,alltop1,alltop2,alltop3
	if request("allhot1")=1 then
		allhot1=1
	else
		allhot1=0
	end if
	if request("hot1")=1 then
		hot1=1
	else
		hot1=0
	end if
	if request("daytop1")=1 then
		daytop1=1
	else
		daytop1=0
	end if
	if request("weektop1")=1 then
		weektop1=1
	else
		weektop1=0
	end if
	if request("alltop1")=1 then
		alltop1=1
	else
		alltop1=0
	end if
	if request("allhot2")=1 then
		allhot2=1
	else
		allhot2=0
	end if
	if request("hot2")=1 then
		hot2=1
	else
		hot2=0
	end if
	if request("daytop2")=1 then
		daytop2=1
	else
		daytop2=0
	end if
	if request("weektop2")=1 then
		weektop2=1
	else
		weektop2=0
	end if
	if request("alltop2")=1 then
		alltop2=1
	else
		alltop2=0
	end if
	if request("allhot3")=1 then
		allhot3=1
	else
		allhot3=0
	end if
	if request("hot3")=1 then
		hot3=1
	else
		hot3=0
	end if
	if request("daytop3")=1 then
		daytop3=1
	else
		daytop3=0
	end if
	if request("weektop3")=1 then
		weektop3=1
	else
		weektop3=0
	end if
	if request("alltop3")=1 then
		alltop3=1
	else
		alltop3=0
	end if
	set rs = server.CreateObject ("Adodb.recordset")
	sql="select * from showpage"
	rs.open sql,conndown,1,3
	rs("showset")=allhot1&","&hot1&","&daytop1&","&weektop1&","&alltop1&","&allhot2&","&hot2&","&daytop2&","&weektop2&","&alltop2&","&allhot3&","&hot3&","&daytop3&","&weektop3&","&alltop3
	rs.update
	rs.close
	set rs=nothing
	response.write("�޸ĳɹ���")
end sub

sub save()
	dim SQL_Statement
	SQL_Statement=Request("SQL_Statement")
	if trim(SQL_Statement)<>"" then
		On Error Resume Next 
		conndown.Execute(SQL_Statement)
		if err.number="0" then
			response.write "ִ�гɹ�"
		else
			response.write "��������⣬����������£�"
			response.write Err.Description
			err.clear
		end if
	end if
end sub

sub renfile()
	dim filename,filenameold,filenamesplit,tempsplit,rannum
	dim downpath,fsdir,fsdir1,fs,file1,file2
	set rs=server.createobject("ADODB.Recordset")
	sql="select * from notdown"
	rs.open sql,conndown,1,3
	downpath=rs("downpath")
	rs.close
	sql="select * from download where ldown=1"
	rs.open sql,conndown,1,3
	do while not rs.eof
		filenameold=rs("filename")
		if InStr(filenameold,"/")<>0 or filenameold="" then
			response.write("�ļ���"&filenameold&"����ȷ������ϸ�Ķ�˵��")
		else
			filenamesplit=split(filenameold,".")
			tempsplit=split(filenamesplit(0),"$")
			do
				randomize
				rannum=right("000000"&int(999999*rnd),6)
				filename=tempsplit(0)&"$"&rannum
				for i=1 to ubound(filenamesplit)
					filename=filename&"."&filenamesplit(i)
				next
			loop while filename=filenameold
			fsdir=Server.MapPath("index.asp")	
			fsdir1=replace(fsdir,"\index.asp","\")
			set fs=createobject("scripting.filesystemobject") 
			file1=fsdir1&"\"&downpath&"\"&filenameold
			file2=fsdir1&"\"&downpath&"\"&filename
			If not fs.fileexists(file1) then
				response.write fsDir&"ԭ�ļ�������,���޸����"&filenameold&"���ļ�����<br>"
			elseIf fs.fileexists(file2) then
				response.write "Ŀ���ļ�"&filename&"�Ѿ�����,�����¸�����<br>"
			else 
				fs.MoveFile file1,file2
				If not fs.fileexists(file1) and fs.fileexists(file2) then
					response.write "�ļ����޸�<font color=red>�ɹ���</font><br>"
					rs("filename")=filename
					rs.update
				Else
					response.write "�ļ����޸�<font color=red>ʧ�ܣ�</font><br>"
				End If
			end if
		end if
		rs.movenext
	loop
	response.write("�����ļ��ĸ��������ɹ���<br>")
	rs.close
	set rs=nothing
end sub%>
