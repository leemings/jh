<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%
'---------------------
' write by ��ˮ��ɽ
'---------------------
	stats="׫д���"

	call nav()
	call head_var(0,0,"���������","z_dglistall.asp")
		
if not founduser then
	errmsg=errmsg+"<br>"+"<li>���˲����Բ��ĵ���б�"
	errmsg=errmsg+"<br>"+"<li>����<a href=login.asp><font color=blue>��¼</font></a>,��û��<a href=reg.asp><font color=blue>ע��</font></a>"
	founderr=true
end if
if founderr then
	call dvbbs_error()
else
%>
	<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
		<tr>
			<th><a href=z_dglistall.asp><font color=white><b>���е���б�</b></font></a></th>
			<th><a href=z_dglistme.asp><font color=white><b>�ҵĵ���б�</b></font></a></th>
			<th><a href=z_dgwrite.asp><font color=white><b>��Ҫ���</b></font></a></th>
		</tr>
		<tr>
			<td class=tablebody2 align=center valign=middle colspan="3">���׸���ף����ĺ���</td>
		</tr>
	</table>
		<br>
	<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1 style="width:55%">
		<tr> 
			<th colspan="2"><b>����������������Ϣ</b></th> 
		</tr> 
			<tr>
				<td colspan="2" class=tablebody2><%
				if isnull(myvip) or myvip<>1 then
					response.write "<font color=forum_info(8)>����ǰ�����ֽ�"&mymoney&"Ԫ��Ϊȫ���Ա�����Ҫ�ֽ�"&forum_user(18)&"Ԫ��Ϊ��һ��Ա�����Ҫ�ֽ�"&forum_user(19)&"Ԫ</font>"
				else
					response.write "<font color=forum_info(8)>����ǰ�����ֽ�"&mymoney&"Ԫ��Ϊȫ���Ա�����Ҫ�ֽ�"&forum_user(20)&"Ԫ��Ϊ��һ��Ա�����Ҫ�ֽ�"&forum_user(21)&"Ԫ</font>"
				end if%></td>
			</tr>		
		<form name=codeform action="z_dgsave.asp" method=post> 
			<tr> 
				<td class=tablebody2 width="25%" valign=middle><b>�͸���</b><br>֧�ֶ��û�</td>
				<td class=tablebody2 width="75%" valign=middle> 
					<input name="incept" type=text id="incept" value="<%=Trim(Request.QueryString("name"))%>" size=40>
			&nbsp;<font color="#FF0000">*</font> <br>ʹ�ö���&quot;|&quot;�ֿ������5λ�û�		<input name="alluser" type=checkbox id="alluser" value="" onclick="{if(document.forms[0].incept.value=='ȫ���Ա') {document.forms[0].incept.value='';return;} document.forms[0].incept.value='ȫ���Ա';return;}">ȫ���Ա</td> 
			</tr> 
			<tr> 
				<td class=tablebody2 width="25%" height="24" valign=top ><b>�������֣�</b><br>50������</td>
				<td class=tablebody2 width="75%" valign=middle> 
					<input name="medianame" type=text id="medianame" size=40 maxlength=50> &nbsp;<font color="#FF0000">*</font> 
				</td>
			</tr> 
			<tr> 
				<td class=tablebody2 valign=top><b>���ֵ�ַ��</b></td>
				<td class=tablebody2 valign=middle>
					<input name="url" type="text" id="url" value="http://" size="50" maxlength="100">&nbsp;<font color="#FF0000">*</font><br>
					<font color=ff0000><b>��</b></font><a href=http://search.sogua.com/search/search.asp target=_blank title=��������><font color=blue><b><u>�������ȥ�Ҹ�</u></b></font></a><br>
        ֧�ָ�ʽ:mp3��mid��wav��wma��swf��mpg��rm��ra��asf<br>
        ֧��Э��:http��rtsp��mms��pnm </td>
			</tr> 
			<tr> 
				<td class=tablebody2 valign=top><b>ף�����ݣ�</b><br>200������
					<br>
					<br> 
				</td>
				<td class=tablebody2  valign=middle>
					<textarea name="content" cols="50" rows="5" id="content"></textarea> &nbsp;<font color="#FF0000">*</font>
				</td>
			</tr> 
			<tr> 
				<td class=tablebody2  valign=middle colspan=2 align=center>
					<INPUT TYPE="submit" NAME="Submit" VALUE="�ͳ�">������
					<INPUT TYPE="reset" NAME="Submit2" VALUE="��д">
				</td>
			</tr>
			<tr>
				<td colspan="2" class=tablebody2>&nbsp;</td>
			</tr>			
		</FORM>
	</table>
	<br>
	<table cellpadding=3 cellspacing=1 border=0 align=center class=tableborder1>
		<tr>
			<th><font color=white>--== ��̳���̨-׫д��� ==--</font></th>
		</tr>
	</table>
<%
call activeonline()
call footer() 
end if	
%>	