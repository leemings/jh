<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp" -->
<!--#include file=z_down_conn.asp-->
<!--#include file="z_down_function.asp" -->
<head>
<title><%=Forum_info(0)%>--����ҳ��</title>
<link rel="stylesheet" href="forum_admin.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<base target="footer">
</head>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0">
<%if not master or session("flag")="" then
	Errmsg=Errmsg+"<br>"+"<li>��ҳ��Ϊ����Աר�ã���<a href=admin_index.asp target=_top>��¼</a>����롣<br><li>��û�й���ҳ���Ȩ�ޡ�"
	call dvbbs_error()
else%>
	<script src="inc/ubbcode.js"></script>
	<table width="95%" border="0" cellpadding="3" cellspacing="1" align=center class=tableborder>
		<form method="POST" name=frmAnnounce action="z_admin_down_savesoft.asp?action=add">
		<tr>
			<th height="30" colspan=2>�� �� �� �� �� ��</td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>������ͣ�</b>&nbsp;</td>
			<td width="85%" class=forumrow>ѡ����ࣺ<select class="smallSel" name="classid" size="1" onChange="changelocation(document.frmAnnounce.classid.options[document.frmAnnounce.classid.selectedIndex].value)"><%
				dim DClassID
				DClassID=""
				set rs=server.createobject("adodb.recordset")
				sql="select * from Dclass"
				rs.open sql,conndown,1,1
				if rs.eof and rs.bof then
					response.write "<option value="""" name=classid>û�з���</option>"
				else
					DClassID=rs("classID")
					do while not rs.eof
						response.write "<option value='"&CStr(rs("classID"))&"' name=classid>"&rs("class")&"</option>"&vbNewLine
						rs.movenext
					loop
				end if	
				rs.close
				%></select>&nbsp;&nbsp;ѡ��С�ࣺ<select name="Nclassid" size="1"><%
				if DClassID="" then
					sql="select * from DNclass order by Nclassid"
				else
					sql="select * from DNclass where classid="&DClassID&" order by Nclassid"
				end if
				rs.open sql,conndown,1,1
				if rs.eof and rs.bof then
					response.write "<option value="""" name=Nclassid>û�з���</option>"
				else
					do while not rs.eof
						response.write "<option value='" + Cstr(rs("Nclassid")) + "' name=Nclassid>" + rs("Nclass") + "</option>"
						rs.MoveNext
					Loop
				end if
				rs.close
				set rs=nothing
			%></select>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>�������ԣ�</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><input type=radio name="downshow" value="1">�������� <input type=radio name="downshow" value="2" checked>��Ա����<%if vipshow=2 then%> <input type=radio name="downshow" value="3">VIP����<%end if%> <input type=radio name="downshow" value="4">��Լ���� <input type=radio name="downshow" value="5">�������� <input name="money" size=6 value="0"></td>
		</tr> 
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>������ԣ�</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><select name="softsx" size="1"><option selected value="0" name="softsx">��ͨ���</option><option value="1" name="softsx">RealPlay����</option><option value="2" name="softsx">MediaPlay����</option><option value="3" name="softsx">FLASH����</option></select>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>��ʾ���ƣ�</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><input type="text" name="txtshowname" size="103" class="smallinput" maxlength="100">&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr> 
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>���ص�ַ1��</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><input type="text" name="txtfilename" size="103" class="smallinput" maxlength="255">&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font>&nbsp;<font color=red>�Ƿ����</font><input type="checkbox" name="localdown" value="on"></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>���ص�ַ2��</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><input type="text" name="txtfilename1" size="103" class="smallinput" maxlength="255" value=""></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>���ص�ַ3��</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><input type="text" name="txtfilename2" size="103" class="smallinput" maxlength="255" value=""></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>ÿ�����ش������ƣ�</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><input type="text" name="daydown" size="15" class="smallinput" maxlength="100" value="0">�˴Σ���д��0����ʾ�����ƣ�&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>������Ч�ڣ�</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><select name="usetime" size="1"><option selected value="0" name="usetime" selected>������</option><option value="1" name="usetime">һ��</option><option value="2" name="usetime">����</option><option value="3" name="usetime">����</option><option value="7" name="usetime">һ��</option><option value="14" name="usetime">����</option><option value="30" name="usetime">һ����</option><option value="60" name="usetime">������</option><option value="90" name="usetime">������</option><option value="180" name="usetime">����</option><option value="360" name="usetime">һ��</option></select>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>�Ƽ��ȣ�</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><select name="hot" size="1"><option selected value="1" name="hot">һ�Ǽ�</option><option value="2" name="hot">���Ǽ�</option><option value="3" name="hot">���Ǽ�</option><option value="4" name="hot">���Ǽ�</option><option value="5" name="hot">���Ǽ�</option></select>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>���л�����</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><select name="system" size="1"><option selected value="ASP����" name="system">ASP����</option><option value="PHP����" name="system">PHP����</option><option value="CGI����" name="system">CGI����</option><option value="JSP����" name="system">JSP����</option><option value="Windows����" name="system">Windows����</option><option value="Linux����" name="system">Linux����</option></select>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>�����С��</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><input type="text" name="size1" size="20" class="smallinput" maxlength="100"></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>�����Դ��</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><input type="text" name="fromurl" size="103" class="smallinput" maxlength="100" value="http://"></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>��Ȩ��ʽ��</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><select name="order" size="1"><option selected value="�������" name="order">�������</option><option value="�������" name="order">�������</option><option value="��ҵ���" name="order">��ҵ���</option></select>&nbsp;�Ƿ��Ƽ�<input type="checkbox" name="hots" value="on">&nbsp;�Ƿ�����<input type="checkbox" name="hide" value="on">&nbsp;�Ƿ�̶��Ƽ�<input type="checkbox" name="gudin" value="on"></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>�ϴ�ͼƬ��</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><iframe name=ad frameborder=0 width=400 height=24 scrolling=no src=z_admin_down_picupload.asp></iframe></td>
		</tr>
		<tr>
			<td width="15%" align="right" height="30" class=forumrow><b>���ͼƬ��</b>&nbsp;</td>
			<td width="85%" height="30" class=forumrow><input type="text" name="pic" size="103" class="smallinput" maxlength="100" value=""></td>
		</tr>
		<tr>
			<td width="15%" align="right" valign="top" class=forumrow><br><b>��ʽ��飺</b>&nbsp;</td>
			<td width="85%" class=forumrow><!--#include file="inc/getubb.asp"--><textarea rows="6" name="Content" cols="103" class="smallarea"></textarea>&nbsp;<font color=<%=forum_body(8)%>><B>**</b></font></td>
		</tr>
		<tr>
			<td colspan=2 height="30" align=center class=forumrow><input type="submit" value=" �� �� " name="cmdok">&nbsp; <input type="reset" value=" �� �� " name="cmdcancel"></td>
		</tr>
		</form>
	</table>
<%end if%>
</body>
<%set rs=server.createobject("adodb.recordset")
sql = "select * from DNclass order by Nclassid asc"
rs.open sql,conndown,1,1%>
<SCRIPT language = "JavaScript">
	var onecount;
	onecount=0;
	subcat = new Array();
	<%dim count
	count = 0
	do while not rs.eof%>
		subcat[<%=count%>] = new Array("<%= trim(rs("Nclass"))%>","<%=cstr(rs("classid"))%>","<%=cstr(rs("Nclassid"))%>");
		<%count = count + 1
		rs.movenext
	loop
	rs.close%>
	onecount=<%=count%>;

	function changelocation(locationid) {
		document.frmAnnounce.Nclassid.length = 0; 
		var locationid=locationid;
		var i;
		for (i=0;i < onecount; i++) {
			if (subcat[i][1] == locationid) { 
				document.frmAnnounce.Nclassid.options[document.frmAnnounce.Nclassid.length] = new Option(subcat[i][0], subcat[i][2]);
			} 
		}
	} 
</SCRIPT> 
