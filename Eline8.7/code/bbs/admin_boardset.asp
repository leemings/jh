<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<%

	stats="��������ҳ��"
	dim sql1,rs1
	if not founduser  then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>���¼����в�����"
	end if

	if request("boardid")="" then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>��ָ����̳���档"
	elseif not isInteger(request("boardid"))  or request("boardid")=0  then
		founderr=true
		Errmsg=Errmsg+"<br>"+"<li>�Ƿ��İ��������"
	else
		boardid=clng(request("boardid"))
	end if

	if not(  boardmaster or  master or  superboardmaster ) then
	founderr=true
	Errmsg=Errmsg+"<br>"+"<li>ֻ�й���Ա���ܵ�¼��"
	end if

	call nav()
if founderr then
	call head_var(1,BoardDepth,0,0)
	call dvbbs_error()
else
	call head_var(1,BoardDepth,0,0)
	dim title
	dim content
	set rs=server.createobject("adodb.recordset")
	call main()
end if

	set rs=nothing
	call footer()

	sub main()
%>
<TABLE cellpadding=0 cellspacing=1 class=tableborder1 align=center > 
        <tr >
          <th height=24 align=center colspan="2">��ӭ <%=htmlencode(membername)%>�����������ҳ��</th>
        </tr>
        <tr >
          <td height=24 align=center colspan="2" class=tablebody1>
        <b>����ѡ�<a href=admin_boardset.asp?boardid=<%=boardid%>>��̳���淢��</a>
        <%if master then%>
        | <b><a href=admin_boardset.asp?action=manage&boardid=<%=boardid%>>�������</a>
        <%end if%> | 
		<a href=admin_boardset.asp?action=editbminfo&boardid=<%=boardid%>>������Ϣ����</a> | 
		<a href=admin_boardset.asp?action=editbmset&boardid=<%=boardid%>>�������ù���</a> | 
		<!--<a href=admin_boardset.asp?action=editbmset&boardid=<%=boardid%>>�������ù���</a> | -->
		<a href=admin_boardset.asp?action=editbmcolor&boardid=<%=boardid%>>��ɫ���ù���</a> | 
		<!--<a href=admin_boardset.asp?action=editbmads&boardid=<%=boardid%>>�ְ������</a> | -->
		<a href=admin_boardset.asp?action=editbmads&boardid=<%=boardid%>>�ְ������</a>
		</b></td>
        </tr>
		<tr>
              <td width="30%" valign=top align=center class=tablebody1 >
		<table cellpadding=3 cellspacing=1   class=tableborder2 style="width:90%;word-break:break-all;" >
		<tr>
			<th width="100%" height=24  colspan="2">�� ������Ϣ�� ��
			</th>
		</tr>
		<tr>
			<td  height=24 class=tablebody2 colspan="2" align=center ><%=boardtype%>
			</td>
		</tr>
		<tr>
			<td width="60%" height=24 class=tablebody1 >����������
			</td>
			<td width="40%" height=24 class=tablebody1 ><FONT COLOR=RED><%=todaynum%></FONT>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody2 >�������ӣ�
			</td>
			<td  height=24 class=tablebody2 ><%=LastTopicNum%>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody1 >�������ӣ�
			</td>
			<td  height=24 class=tablebody1 ><%=LastBbsNum%>
			</td>
		</tr>
		<tr>
			<td width="30%" height=24 class=tablebody2 colspan="2">�����Ա��
		<%=boardmasterlist%>
			</td>
		</tr>
		<tr>
			<th width="100%" height=24  colspan="2">�� ����Ȩ�� ��
			</th>
		</tr>
		<tr>
			<td  height=24 class=tablebody1 >����������ɾ��������
			</td>
			<td  height=24 class=tablebody1 ><%if Board_Setting(33)=1 then%>��<%else%><FONT COLOR=RED>�ر�</FONT><%end if%>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody2 >���������޸���ɫ���ã�
			</td>
			<td  height=24 class=tablebody2 ><%if Board_Setting(34)=1 then%>��<%else%><FONT COLOR=RED>�ر�</FONT><%end if%>
			</td>
		</tr>
		<tr>
			<td  height=24 class=tablebody1 >���а��������޸���ɫ���ã�
			</td>
			<td  height=24 class=tablebody1 ><%if Board_Setting(35)=1 then%>��<%else%><FONT COLOR=RED>�ر�</FONT><%end if%>
			</td>
		</tr>
		<tr>
			<td width="100%" height=24  colspan="2" class=tablebody2>
		<b>ע�⣺</b>������������������Լ��������ɷ�������Ͱ������ã�����Ա���������а��淢����������Ϣ���й��������
			</td>
		</tr>
		</table>
	      </td>
              <td width="70%" valign=top class=tablebody1 align=center>
      		<table cellpadding=3 cellspacing=1  class=tableborder2 style="width:100%;word-break:break-all;">
		  <tr>
			<td width="100%" height=24 class=tablebody2 >
<B>ע��</B>��<BR>��ҳ��Ϊ����ר�ã�ʹ��ǰ�뿴������Ӧ�Ĺ����Ƿ�򿪣��ڽ��й������õ�ʱ�򣬲�Ҫ����������ã�������ģ�������д����������ȷ����д��
		  </td></tr>
		</table>
<%
	select case request("action")
	case "new"
		call savenews()
	case "manage"
		call manage()
	case "edit"
		call edit()
	case "updat"
		call update()
	case "del"
		call del()
	case "editbminfo"
		call editbminfo()
	case "saveditbm"
		call savebminfo()
	case "editbmset"
		call editbmset()
	case "savebmset"
		call savebmset()
	case "editbmcolor"
		call editbmcolor()
	case "savebmcolor"
		call savebmcolor()
	case "editbmads"
		call editbmads()
	case "savebmads"
		call savebmads()
	case else
	call news()
	end select
	if founderr then call dvbbs_error()
%>
        </td>
    </tr>
</table>
<%
end sub

sub news()
%>
<form action="admin_boardset.asp?action=new" method=post name=FORM>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
    <tr> 
      <td width="20%" valign=top class=tablebody1> 
        <div align="right">�������棺 </div>
      </td>
      <td width="80%" class=tablebody1> 
        <%if master then%>
        <%
   sql="select boardid,boardtype from board"
   rs.open sql,conn,1,1
%>
        <select name="boardid" size="1">
<!--           <option value="0">��̳��ҳ</option> -->
          <%
	do while not rs.eof
        response.write "<option value='"+CStr(rs("BoardID"))+"'>"+rs("Boardtype")+"</option>"+chr(13)+chr(10)
	rs.movenext
	loop
	rs.close
%>
        </select>
        <%else%>
        <%
	sql="select boardtype from board where boardid="&boardid
	rs.open sql,conn,1,1
	boardtype=rs("boardtype")
%>
        <select name="boardid" size="1">
          <option value="<%=boardid%>"><%=boardtype%></option>
        </select>
        <%end if%>
      </td>
    </tr>
    <tr> 
      <td width="20%" valign=top class=tablebody1> 
        <div align="right">�����ˣ� </div>
      </td>
      <td width="80%" class=tablebody1>
        <input type=text name=username size=36 value="<%=membername%>" disabled>
        <input type=hidden name=username value="<%=membername%>">
      </td>
    </tr>
    <tr> 
      <td width="20%" valign=top class=tablebody1> 
        <div align="right">���⣺ </div>
      </td>
      <td width="80%" class=tablebody1>
        <input type=text name=title size=60>
      </td>
    </tr>
    <tr> 
      <td width="20%" valign=top class=tablebody1> 
        <div align="right">���ݣ� </div>
      </td>
      <td width="80%" class=tablebody1>
        <textarea cols=60 rows=6 name="content"></textarea>
      </td>
    </tr>
    <tr>
      <td width="100%" valign=top colspan="2" align=center class=tablebody2> 
        <input type=Submit value="�� ��" name=Submit">
        &nbsp; 
        <input type="reset" name="Clear" value="�� ��">
      </td>
    </tr>
  </table>
</form>
<%
end sub

sub savenews()
	dim username,title,content
	if request("boardid")<>"" or (not isInteger(request("boardid"))) then
		boardid=clng(request("boardid"))
	else
		Errmsg=Errmsg+"<br>"+"<li>�������˴���Ĳ�����"
		founderr=true
	end if
	if request("username")="" then
		Errmsg=Errmsg+"<br>"+"<li>�����뷢���ߡ�"
		founderr=true
	else
		username=checkstr(request("username"))
	end if
	if request("title")="" then
		Errmsg=Errmsg+"<br>"+"<li>���������ű��⡣"
		founderr=true
	else
		title=checkstr(request("title"))
	end if
	if request("content")="" then
		Errmsg=Errmsg+"<br>"+"<li>�������������ݡ�"
		founderr=true
	else
		content=checkstr(request("content"))
	end if

	if founderr=true then
		call dvbbs_error()
		exit sub
	end if 
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(25))=1 then
		sql="select * from bbsnews"
		rs.open sql,conn,1,3
		rs.addnew
		rs("username")=username
		rs("title")=title
		rs("content")=content
		rs("addtime")=Now()
		rs("boardid")=boardid
		rs.update
		rs.close
		myCache.name="AnnounceMents"&BoardID
		myCache.makeEmpty
		call success()
		else
	Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
	founderr=true
	exit sub
	end if
end sub

sub manage()
dim caneditann
caneditann=false
if (master or superboardmaster or boardmaster) and Cint(GroupSetting(26))=1 then
	caneditann=true
else
	caneditann=false
end if
	if   not caneditann  then
	Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
	founderr=true
	exit sub
	end if
	sql="select * from bbsnews where boardid<>0"
	rs.open sql,conn,1,1
%>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
		  <tr><th width="80%" valign=top height=22 >
����
		  </th>
		  <th width="20%" >
����
		  </th></tr>
<%do while not rs.eof%>
		  <tr><td width="80%" valign=top height=22 class=tablebody1><a href=admin_boardset.asp?action=edit&id=<%=rs("id")%>&boardid=<%=rs("boardid")%>><%=rs("title")%></a>
		  </td>
		  <td width="20%" class=tablebody2><a href=admin_boardset.asp?action=del&id=<%=rs("id")%>&boardid=<%=boardid%>>ɾ��</a>
		  </td></tr>
<%
	rs.movenext
	loop
	rs.close
%></table>
<%
end sub

sub edit()
%>
<form action="admin_boardset.asp?action=updat&id=<%=request("id")%>" method=post>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
		  <tr><td  valign=top align=right class=tablebody1>�������棺
		  </td>
		  <td  class=tablebody1>
<%
	dim sel
   	sql="select boardid,boardtype from board"
   	rs.open sql,conn,1,1
%>
<select name="boardid" size="1">
<option value="0" <%if request("boardid")=0 then%>selected<%end if%>>��̳��ҳ</option>
<%
	do while not rs.eof
	if Clng(request("boardid"))=Clng(rs("boardid")) then
	sel="selected"
	else
	sel=""
	end if
        response.write "<option value='"+CStr(rs("BoardID"))+"' "&sel&">"+rs("Boardtype")+"</option>"+chr(13)+chr(10)
	rs.movenext
	loop
	rs.close
%>        
          </select>
		  </td></tr>
<%
	sql="select * from bbsnews where id="&cstr(request("id"))
	rs.open sql,conn,1,1
%>
		  <tr><td width="20%" valign=top align=right class=tablebody1>
�����ˣ�
		  </td>
		  <td width="80%"  class=tablebody1><input type=text name=username size=36 value=<%=rs("username")%>></td></tr>
		  <tr><td width="20%" valign=top align=right class=tablebody1>
���⣺
		  </td>
		  <td width="80%"  class=tablebody1><input type=text name=title size=60 value=<%=rs("title")%>></td></tr>
		  <tr><td width="20%" valign=top align=right class=tablebody1>
���ݣ�
		  </td>
		  <td width="80%" class=tablebody1><textarea cols=60 rows=6 name="content">
<%
	    content=replace(rs("content"),"<br>",chr(13))
        content=replace(content,"&nbsp;"," ")
            response.write ""&content&""
	    rs.close
%>
		  </textarea></td>
		  </tr>
		  <tr><td width="100%" valign=top colspan="2" align=center class=tablebody2>
<input type=Submit value="�� ��" name=Submit"> &nbsp; <input type="reset" name="Clear" value="�� ��">
		  </td></tr>
		</table>
</form>
<%
end sub

sub update()
	dim username,title,content
	dim caneditann
	caneditann=false
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(25))=1 then
	caneditann=true
	else
	caneditann=false
	end if
	if   not caneditann  then
	Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
	founderr=true
	exit sub
	end if
	if request("id")="" or not isnumeric(request("id"))  then
	Errmsg=Errmsg+"<br><li>��ѡ����ȷ�Ĺ��档"
	founderr=true
	exit sub
	end if
	if request("boardid")<>"" or (not isInteger(request("boardid")))  then
		boardid=clng(request("boardid"))
	else
		Errmsg=Errmsg+"<br><li>��ѡ����ȷ����̳��"
		founderr=true
		exit sub
	end if
	if request("username")="" then
		Errmsg=Errmsg+"<br>"+"<li>�����뷢���ߡ�"
		founderr=true
	else
		username=checkstr(request("username"))
	end if
	if request("title")="" then
		Errmsg=Errmsg+"<br>"+"<li>���������ű��⡣"
		founderr=true
	else
		title=checkstr(request("title"))
	end if
	if request("content")="" then
		Errmsg=Errmsg+"<br>"+"<li>�������������ݡ�"
		founderr=true
	else
		content=checkstr(request("content"))
	end if
	if founderr=true then
		call dvbbs_error()
	else
		sql="select * from bbsnews where id="&cstr(request("id"))
		rs.open sql,conn,1,3
		rs("username")=username
		rs("title")=title
		rs("content")=content
		rs("addtime")=Now()
		rs("boardid")=boardid
		rs.update
		rs.close
		myCache.name="AnnounceMents"&BoardID
		myCache.makeEmpty
		call success()
	end if
end sub

sub del()
	dim caneditann
	caneditann=false
	if (master or superboardmaster or boardmaster) and Cint(GroupSetting(26))=1 then
	caneditann=true
	else
	caneditann=false
	end if

	if   not caneditann  then
	Errmsg=Errmsg+"<br><li>��û��ִ�д˲�����Ȩ�ޡ�"
	founderr=true
	exit sub
	end if

	if request("id")="" and not isnumeric(request("id"))  then
	Errmsg=Errmsg+"<br><li>��ѡ����ȷ�Ĺ��档"
	founderr=true
	exit sub
	end if
	
	dim delid
	delid=replace(request("id"),"'","")
	delid=replace(delid,";","")
	delid=replace(delid,"--","")
	delid=replace(delid,")","")
	conn.execute("delete from bbsnews where id in ("&delid&")")
	myCache.name="AnnounceMents"&BoardID
	myCache.makeEmpty
	call success()
end sub

sub success()
%><br><br>
�ɹ������Ų���
<%
end sub

sub editbminfo()
dim master_1

%><form action ="admin_boardset.asp?action=saveditbm&boardid=<%=boardid%>" method=post> 
<%
set rs= server.CreateObject ("adodb.recordset")
sql = "select * from board where boardid="+CSTr(boardid)
rs.open sql,conn,1,1
if rs.eof and rs.bof then
   Errmsg=Errmsg+"<br>"+"<li>��û��ָ����Ӧ��̳ID�����ܽ��й���"
   call dvbbs_error()
exit sub
   end if
if not master then
	if Board_Setting(33)=1  then
		master_1=split(rs("boardmaster"),"|")
		if membername<>master_1(0)  then
		Errmsg=Errmsg+"<br>"+"<li>�����Ϊ������ר�á�"
		call dvbbs_error()
		exit sub
		end if
	else
	Errmsg=Errmsg+"<br>"+"<li>��δ���޸����õ�Ȩ�ޡ�"
	call dvbbs_error()
	exit sub
	end if
end if
%>
             <input type='hidden' name=editid value='<%=boardid%>'>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
    <tr> 
      <th width="20%" height=22 class=tablebody2><b>�ֶ����ƣ�</b> </td>
      <th  > 
        <div align="center"><b>����ֵ��</b></div>
      </th>
    </th>
    <tr> 
      <td height=22 class=tablebody1  align="center">��̳���ƣ�</td>
      <td  class=tablebody1>
	  <input type="text" name="boardtype" size="30" value='<%=htmlencode(rs("boardtype"))%>'>
	  </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody2  align="center">����˵����</td>
      <td  class=tablebody1> 
        <input type="text" name="readme" size="60" value='<%=htmlencode(rs("readme"))%>'>
      </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody1  align="center">�����޸ģ�</td>
      <td  class=tablebody1> 
        <input type="text" name="boardmaster" size="30" value='<%=rs("boardmaster")%>'>(������������|�ָ����磺ɳ̲С��|wodeail)
				<input type="hidden" name="oldboardmaster" value='<%=rs("boardmaster")%>'>
      </td>
    </tr>
    <tr> 
      <td height=22 class=tablebody2>&nbsp;</td>
      <td  class=tablebody2> 
        <input type="submit" name="Submit" value="�ύ">
      </td>
    </tr>
  </table>
</form>
<%
rs.close

end sub

sub savebminfo()
dim rname
dim readme,boardtype,boardmaster
	if request("boardid")<>"" or (not isInteger(request("boardid")))  then
		boardid=clng(request("boardid"))
	else
		Errmsg=Errmsg+"<br><li>��ѡ����ȷ����̳��"
		founderr=true
		exit sub
	end if
readme=checkStr(Request.form("readme"))
boardtype=checkStr(Request.form("boardtype"))
boardmaster=checkStr(Request.form("boardmaster"))
	if readme="" then
		Errmsg=Errmsg+"<br>"+"<li>��������̳��顣"
		founderr=true
	end if
	if boardtype="" then
		Errmsg=Errmsg+"<br>"+"<li>��������̳���ơ�"
		founderr=true
	end if
	if boardmaster="" then
		Errmsg=Errmsg+"<br>"+"<li>����������Ա��"
		founderr=true
	end if
	if founderr=true then 
	call dvbbs_error()
	exit sub
	end if
rname=split(boardmaster,"|")
for i=0 to ubound(rname)
sql="select top 1 username from [user] where username='"&replace(rname(i),"'","")&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	founderr=true
	errmsg=errmsg+"<br>"+"<li>��̳û������û����������Ϊ����"
	exit sub
	exit for
end if
set rs=nothing
next
set rs=server.createobject("adodb.recordset")
sql = "select * from board where boardid="+Cstr(request("boardid"))
rs.open sql,conn,1,3
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>��û��ָ����Ӧ��̳ID�����ܽ��й���"
call dvbbs_error()
exit sub
end if
	rs("boardmaster") = boardmaster
	rs("readme") = readme
	rs("boardtype")=boardtype
	rs.Update 
	rs.Close 
if request("oldboardmaster")<>boardmaster then call addmaster(boardmaster,request("oldboardmaster"),1)
	response.write "<p>��̳�޸ĳɹ���"
end sub

sub editbmset()
dim chkedit
dim master_1
chkedit=false
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>��û��ָ����Ӧ��̳ID�����ܽ��й���"
call dvbbs_error()
exit sub
end if
if isnull(rs(0)) or rs(0)="" then
Errmsg=Errmsg+"<br>"+"<li>����̳��δ�й���Ա��"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>�����Ϊ������ר�á� "
	call dvbbs_error()
else
%>
<form method="POST" action=admin_boardset.asp?action=savebmset&boardid=<%=request("boardid")%>>
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
<tr> 
<td height="23" colspan="2" class=Tablebody2><b>��̳ʹ������</b></td>
</tr>
<tr> 
<th height="23" colspan="2" align=left><a name="setting1"><b>�򿪻��߹ر���̳</b></a>[<a href="#top"><FONT color="#FFFFFF">����</FONT></a>]</th>
</tr>
<tr> 
<td width="50%" class=tablebody1>
<U>��������</U></td>
<td width="50%" class=tablebody1> 
<select size=1 name="forum_setting(9)">
<option value="lang_gb.asp">��������
</select>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1><U>��̳��ǰ״̬</U><BR>ά���ڼ�����ùر���̳ʹ��</td>
<td width="50%" class=tablebody1> 
<input type=radio name="Forum_Setting(21)" value=0 <%if cint(Forum_Setting(21))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(21)" value=1 <%if cint(Forum_Setting(21))=1 then%>checked<%end if%>>�ر�&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1><U>ά��˵��</U><BR>����̳�ر��������ʾ��֧��html�﷨</td>
<td width="50%" class=tablebody1> 
<textarea name="StopReadme" cols="40" rows="3"><%=StopReadme%></textarea>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1>
<U>�Ƿ����ö�ʱ������̳</U></td>
<td width="50%" class=tablebody1> 
<input type=radio name="forum_setting(69)" value=0 <%if cint(Forum_Setting(69))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(69)" value=1 <%if cint(Forum_Setting(69))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1>
<U>��̳����ʱ��</U><BR>��ȷ�����Ѿ��������ö�ʱ���ع���<BR>��ֹСʱ�÷��š�|���ֿ�</td>
<td width="50%" class=tablebody1> 
<input type=text size=10 name="forum_setting(70)" value="<%=forum_setting(70)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��ҳģʽ</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(4)" value=0 <%if cint(Forum_Setting(4))=0 then%>checked<%end if%>>�°���ʽ&nbsp;
<input type=radio name="Forum_Setting(4)" value=1 <%if cint(Forum_Setting(4))=1 then%>checked<%end if%>>�ɰ���ʽ&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�����˵�ģʽ</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(71)" value=0 <%if cint(Forum_Setting(71))=0 then%>checked<%end if%>>�����̶�&nbsp;
<input type=radio name="Forum_Setting(71)" value=1 <%if cint(Forum_Setting(71))=1 then%>checked<%end if%>>��໬��&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting3"></a><b>��̳������Ϣ</b>[<a href="#top">����</a>]</th>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��̳����</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="ForumName" size="35" value="<%=Forum_info(0)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��̳��url</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="ForumURL" size="35" value="<%=Forum_info(1)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��̳�Ľ�վ����</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="CreateDate" size="35" value="<%=Forum_info(12)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��ҳ����</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="CompanyName" size="35" value="<%=Forum_info(2)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��ҳURL</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="HostUrl" size="35" value="<%=Forum_info(3)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��̳��ҳLogo��ַ</U><BR>��ʾ����̳��ҳ�������̳���û����дlogo��ַ����ʹ�ø�����</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Logo" size="35" value="<%=Forum_info(6)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��̳ͼƬĿ¼</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="Picurl" size="35" value="<%=Forum_info(7)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��������Ŀ¼</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="Faceurl" size="35" value="<%=Forum_info(8)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��������Ŀ¼</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="emoturl" size="35" value="<%=Forum_info(10)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��̳ͷ��Ŀ¼</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="UserFaceurl" size="35" value="<%=Forum_info(11)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��Ȩ��Ϣ</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="Copyright" size="35" value="<%=Copyright%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting4"></a><b>��̳��ϵ����</b>[<a href="#top">����</a>]</th>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��̳����ԱEmail</U><BR>���û������ʼ�ʱ����ʾ����ԴEmail��Ϣ</td>
<td width="50%" class=tablebody1>  
<input type="text" name="SystemEmail" size="35" value="<%=Forum_info(5)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting6"></a><b>���Ļ�ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�¶���Ϣ��������</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(10)" value=0 <%if cint(Forum_Setting(10))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(10)" value=1 <%if cint(Forum_Setting(10))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting7"></a><b>��̳��ҳѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1>
<U>��ҳ��ʾ��̳���</U><BR>0����һ����1����2�����Դ�����<BR>���ù������̳��Ƚ�Ӱ����̳�������ܣ�������Լ���̳��������ã���������Ϊ1</td>
<td width="50%" class=tablebody1> 
<input type=text size=10 name="forum_setting(5)" value="<%=forum_setting(5)%>"> ��
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>���ٵ�¼λ��</U></td>
<td width="50%" class=tablebody1>  
<select name="Forum_Setting(31)">
<option value="1" <%if cint(Forum_Setting(31))=1 then%>selected<%end if%>>����
<option value="2" <%if cint(Forum_Setting(31))=2 then%>selected<%end if%>>�ײ�
<option value="0" <%if cint(Forum_Setting(31))="0" then%>selected<%end if%>>����ʾ
</select>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�Ƿ���ʾ�����ջ�Ա</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(29)" value=0 <%if cint(Forum_Setting(29))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(29)" value=1 <%if cint(Forum_Setting(29))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�Ƿ���ʾ��������Ա</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(43)" value=0 <%if cint(Forum_Setting(43))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(43)" value=1 <%if cint(Forum_Setting(43))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��ҳ�Ƿ���ʾ����б�(ͣ��)</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(12)" value=0 <%if cint(Forum_Setting(12))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(12)" value=1 <%if cint(Forum_Setting(12))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��ҳ�Ƿ���ʾ����վ</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(13)" value=0 <%if cint(Forum_Setting(13))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(13)" value=1 <%if cint(Forum_Setting(13))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��ҳ�Ƿ���ʾ����</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(16)" value=0 <%if cint(Forum_Setting(16))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(16)" value=1 <%if cint(Forum_Setting(16))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>���ǵ�һ�е���ʾ��ʽ</U></td>
<td width="50%" class=tablebody1><select name="Forum_Setting(17)"> 
<option value=1 <%if cint(Forum_Setting(17))=1 then%>selected<%end if%>>���շ�����</option>
<option value=2 <%if cint(Forum_Setting(17))=2 then%>selected<%end if%>>���ܷ�����</option>
<option value=3 <%if cint(Forum_Setting(17))=3 then%>selected<%end if%>>���·�����</option>
<option value=4 <%if cint(Forum_Setting(17))=4 then%>selected<%end if%>>���귢����</option>
<option value=5 <%if cint(Forum_Setting(17))=5 then%>selected<%end if%>>��������</option>
<option value=6 <%if cint(Forum_Setting(17))=6 then%>selected<%end if%>>���������</option>
<option value=7 <%if cint(Forum_Setting(17))=7 then%>selected<%end if%>>���Ů����</option>
<option value=8 <%if cint(Forum_Setting(17))=8 then%>selected<%end if%>>���յ÷���</option>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>���ǵڶ��е���ʾ��ʽ</U></td>
<td width="50%" class=tablebody1><select name="Forum_Setting(18)"> 
<option value=1 <%if cint(Forum_Setting(18))=1 then%>selected<%end if%>>���շ�����</option>
<option value=2 <%if cint(Forum_Setting(18))=2 then%>selected<%end if%>>���ܷ�����</option>
<option value=3 <%if cint(Forum_Setting(18))=3 then%>selected<%end if%>>���·�����</option>
<option value=4 <%if cint(Forum_Setting(18))=4 then%>selected<%end if%>>���귢����</option>
<option value=5 <%if cint(Forum_Setting(18))=5 then%>selected<%end if%>>��������</option>
<option value=6 <%if cint(Forum_Setting(18))=6 then%>selected<%end if%>>���������</option>
<option value=7 <%if cint(Forum_Setting(18))=7 then%>selected<%end if%>>���Ů����</option>
<option value=8 <%if cint(Forum_Setting(18))=8 then%>selected<%end if%>>���յ÷���</option>
</select>
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting8"></a><b>�û���ע��ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�Ƿ��������û�ע��</U><BR>�رպ���̳������ע��</td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(37)" value=0 <%if cint(forum_setting(37))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(37)" value=1 <%if cint(forum_setting(37))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>����û�������</U><BR>��д���֣�����С��1����50</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(40)" size="3" value="<%=forum_setting(40)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��û�������</U><BR>��д���֣�����С��1����50</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(41)" size="3" value="<%=forum_setting(41)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>ͬһIPע����ʱ��</U><BR>�粻�����ƿ���д0</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(22)" size="3" value="<%=Forum_Setting(22)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>Email֪ͨ����</U><BR>ȷ������վ��֧�ַ���mail������������Ϊϵͳ�������</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(23)" value=0 <%if cint(Forum_Setting(23))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(23)" value=1 <%if cint(Forum_Setting(23))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>һ��Emailֻ��ע��һ���ʺ�</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(24)" value=0 <%if cint(Forum_Setting(24))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(24)" value=1 <%if cint(Forum_Setting(24))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>ע����Ҫ����Ա��֤</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(25)" value=0 <%if cint(Forum_Setting(25))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(25)" value=1 <%if cint(Forum_Setting(25))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>����ע����Ϣ�ʼ�</U><BR>��ȷ���������ʼ�����</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(47)" value=0 <%if cint(Forum_Setting(47))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(47)" value=1 <%if cint(Forum_Setting(47))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�������Ż�ӭ��ע���û�</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(46)" value=0 <%if cint(Forum_Setting(46))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(46)" value=1 <%if cint(Forum_Setting(46))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting10"></a><b>ϵͳ����</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��̳����ʱ��</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="GMT" size="35" value="<%=Forum_info(9)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>������ʱ��</U></td>
<td width="50%" class=tablebody1>  
<select name="Forum_Setting(0)">
<%for i=-23 to 23%>
<option value="<%=i%>" <%if cint(i)=cint(Forum_Setting(0)) then%>selected<%end if%>><%=i%>
<%next%>
</select>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�ű���ʱʱ��</U><BR>Ĭ��Ϊ300��һ�㲻������</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(1)" size="3" value="<%=Forum_Setting(1)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�Ƿ���ʾҳ��ִ��ʱ��</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(30)" value=0 <%if cint(Forum_Setting(30))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(30)" value=1 <%if cint(Forum_Setting(30))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1><U>��ֹ���ʼ���ַ</U><BR>������ָ�����ʼ���ַ������ֹע�ᣬÿ���ʼ���ַ�á�|�����ŷָ�<BR>������֧��ģ����������������eway��ֹ������ֹeway@aspsky.net����eway@dvbbs.net����������ע��</td>
<td width="50%" class=tablebody1> 
<input type="text" name="Forum_Setting(52)" size="50" value="<%=Forum_Setting(52)%>">
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1><U>��ˢ�¹�����Ч��ҳ��</U><BR>��ȷ�������˷�ˢ�¹���<BR>��ָ����ҳ�潫�з�ˢ�����ã��û����޶���ʱ���ڲ����ظ��򿪸�ҳ�棬����һ��������Դ���ĵ�����<BR>ÿ��ҳ�������á�|�����Ÿ���</td>
<td width="50%" class=tablebody1> 
<input type="text" name="Forum_Setting(64)" size="50" value="<%=Forum_Setting(64)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting12"></a><b>���ߺ��û���Դ</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>������ʾ�û�IP</U><BR>�رպ���������û��顢��̳Ȩ�ޡ��û�Ȩ�����������û��������ɼ�</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(28)" value=0 <%if cint(Forum_Setting(28))=0 then%>checked<%end if%>>����&nbsp;
<input type=radio name="Forum_Setting(28)" value=1 <%if cint(Forum_Setting(28))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>������ʾ�û���Դ</U><BR>�رպ���������û��顢��̳Ȩ�ޡ��û�Ȩ�����������û��������ɼ�<BR>���������ܽ�������Դ</td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(36)" value=0 <%if cint(forum_setting(36))=0 then%>checked<%end if%>>����&nbsp;
<input type=radio name="forum_setting(36)" value=1 <%if cint(forum_setting(36))=1 then%>checked<%end if%>>����&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>���������б���ʾ�û���ǰλ��</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(33)" value=0 <%if cint(forum_setting(33))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(33)" value=1 <%if cint(forum_setting(33))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>���������б���ʾ�û���¼�ͻʱ��</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(34)" value=0 <%if cint(forum_setting(34))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(34)" value=1 <%if cint(forum_setting(34))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>���������б���ʾ�û�������Ͳ���ϵͳ</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(35)" value=0 <%if cint(forum_setting(35))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(35)" value=1 <%if cint(forum_setting(35))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>����������ʾ��������</U><BR>Ϊ��ʡ��Դ����ر�</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(15)" value=0 <%if cint(Forum_Setting(15))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(15)" value=1 <%if cint(Forum_Setting(15))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>����������ʾ�û�����</U><BR>Ϊ��ʡ��Դ����ر�</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(14)" value=0 <%if cint(Forum_Setting(14))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(14)" value=1 <%if cint(Forum_Setting(14))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>ɾ������û�ʱ��</U><BR>������ɾ�����ٷ����ڲ���û�<BR>��λ�����ӣ�����������</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(8)" size="3" value="<%=Forum_Setting(8)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>����̳����ͬʱ������</U><BR>�粻�����ƣ�������Ϊ0</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(26)" size="6" value="<%=Forum_Setting(26)%>">&nbsp;��
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting13"></a><b>�ʼ�ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�����ʼ����</U><BR>������ķ�������֧�������������ѡ��֧��</td>
<td width="50%" class=tablebody1>  
<select name="Forum_Setting(2)">
<option value="0" <%if Forum_Setting(2)=0 then%>selected<%end if%>>��֧�� 
<option value="1" <%if Forum_Setting(2)=1 then%>selected<%end if%>>JMAIL 
<option value="2" <%if Forum_Setting(2)=2 then%>selected<%end if%>>CDONTS 
<option value="3" <%if Forum_Setting(2)=3 then%>selected<%end if%>>ASPEMAIL 
</select>
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>SMTP Server��ַ</U><BR>ֻ������̳ʹ�������д��˷����ʼ����ܣ�����д���ݷ���Ч</td>
<td width="50%" class=tablebody1>  
<input type="text" name="SMTPServer" size="35" value="<%=Forum_info(4)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting14"></a><b>�ϴ�����</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>ͷ���ϴ�</U></td>
<td width="50%" class=tablebody1> 
<input type=radio name="Forum_Setting(7)" value=0 <%if Forum_Setting(7)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(7)" value=1 <%if Forum_Setting(7)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=tablebody1 width="50%"><U>��������ͷ���ļ���С</U></td>
<td class=tablebody1 width="50%"> 
<input type="text" name="forum_setting(56)" size="6" value="<%=forum_setting(56)%>">&nbsp;K
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting15"></a><b>�û�ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�������ǩ��</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(42)" value=0 <%if forum_setting(42)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(42)" value=1 <%if forum_setting(42)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�����û�ʹ��ͷ��</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(53)" value=0 <%if forum_setting(53)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(53)" value=1 <%if forum_setting(53)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=tablebody1 width="50%"><U>���ͷ����</U><BR>��������Ϊͷ��������</td>
<td class=tablebody1 width="50%"> 
<input type="text" name="forum_setting(57)" size="6" value="<%=forum_setting(57)%>">&nbsp;����
</td>
</tr>
<tr> 
<td class=tablebody1 width="50%"><U>���ͷ��߶�</U><BR>��������Ϊͷ������߶�</td>
<td class=tablebody1 width="50%"> 
<input type="text" name="forum_setting(58)" size="6" value="<%=forum_setting(58)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>Ĭ��ͷ����</U><BR>��������Ϊ��̳ͷ���Ĭ�Ͽ��</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(38)" size="6" value="<%=forum_setting(38)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>Ĭ��ͷ��߶�</U><BR>��������Ϊ��̳ͷ���Ĭ�Ͽ��</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(39)" size="6" value="<%=forum_setting(39)%>">&nbsp;����
</td>
</tr>
<tr> 
<td class=tablebody1 width="50%"><U>ʹ���Զ���ͷ������ٷ�����</U></td>
<td class=tablebody1 width="50%"> 
<input type="text" name="forum_setting(54)" size="6" value="<%=forum_setting(54)%>">&nbsp;ƪ
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>���������վ���ϴ�ͷ��</U><BR>�����Ƿ����ֱ��ʹ��http..������url��ֱ����ʾͷ��</td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(55)" value=0 <%if forum_setting(55)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="forum_setting(55)" value=1 <%if forum_setting(55)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�û�ǩ���Ƿ���UBB����</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(65)" value=0 <%if Forum_Setting(65)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(65)" value=1 <%if Forum_Setting(65)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�û�ǩ���Ƿ���HTML����</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(66)" value=0 <%if Forum_Setting(66)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(66)" value=1 <%if Forum_Setting(66)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�û��Ƿ�����ͼ��ǩ</U><BR>����ͼƬ��flash����ý���</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(67)" value=0 <%if Forum_Setting(67)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(67)" value=1 <%if Forum_Setting(67)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td class=tablebody1 width="50%"><U>�û������б����</U></td>
<td class=tablebody1 width="50%"> 
<input type="text" name="Forum_Setting(68)" size="6" value="<%=Forum_Setting(68)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�û�ͷ��</U><BR>�Ƿ������û��Զ���ͷ��</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(6)" value=0 <%if cint(Forum_Setting(6))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(6)" value=1 <%if cint(Forum_Setting(6))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�û�ͷ����󳤶�</U></td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(59)" size="6" value="<%=forum_setting(59)%>">&nbsp;byte
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�Զ���ͷ�����ٷ�����������</U><BR>��������������Ϊ0</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(60)" size="6" value="<%=forum_setting(60)%>">&nbsp;ƪ
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�Զ���ͷ��ע����������</U><BR>��������������Ϊ0</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(61)" size="6" value="<%=forum_setting(61)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�Զ���ͷ������������������һ������</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(62)" value=0 <%if cint(Forum_Setting(62))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(62)" value=1 <%if cint(Forum_Setting(62))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�Զ���ͷ����Ҫ���εĴ���</U><BR>ÿ�������ַ��á�|�����Ÿ���</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(63)" size="50" value="<%=forum_setting(63)%>">
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting17"></a><b>��ˢ�»���</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��ˢ�»���</U><BR>��ѡ�������д���������ˢ��ʱ��<BR>�԰����͹���Ա��Ч</td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(19)" value=0 <%if cint(Forum_Setting(19))=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="Forum_Setting(19)" value=1 <%if cint(Forum_Setting(19))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>���ˢ��ʱ����</U><BR>��д����Ŀ��ȷ�������˷�ˢ�»���<BR>���������б����ʾ����ҳ��������</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(20)" size="3" value="<%=Forum_Setting(20)%>">&nbsp;��
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting18"></a><b>��̳��ҳ����</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>ÿҳ��ʾ����¼</U><BR>������̳���кͷ�ҳ�йص���Ŀ�������б��������ӳ��⣩</td>
<td width="50%" class=tablebody1>  
<input type="text" name="Forum_Setting(11)" size="3" value="<%=Forum_Setting(11)%>">&nbsp;��
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting16"></a><b>����ѡ��</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>��Ϊ���Ż�����������ֵ</U><BR>��׼Ϊ����ظ���</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(44)" size="3" value="<%=forum_setting(44)%>">&nbsp;��
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�༭����������ʾ"��xxx��yyy�༭"����Ϣ</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(48)" value=0 <%if cint(forum_setting(48))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(48)" value=1 <%if cint(forum_setting(48))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>����Ա�༭����ʾ"��XXX�༭"����Ϣ</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(49)" value=0 <%if cint(forum_setting(49))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(49)" value=1 <%if cint(forum_setting(49))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�ȴ�"��XXX�༭"��Ϣ��ʾ��ʱ��</U><BR>�����û��༭�Լ������Ӷ��������ӵײ���ʾ"��XXX�༭"��Ϣ��ʱ��(�Է���Ϊ��λ)</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(50)" size="3" value="<%=forum_setting(50)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�༭����ʱ��</U><BR>�༭�������ӵ�ʱ������(�Է���Ϊ��λ, 1����1440����) �������ʱ������, ֻ�й���Ա�Ͱ������ܱ༭��ɾ������. �������ʹ�������, ������Ϊ0</td>
<td width="50%" class=tablebody1>  
<input type="text" name="forum_setting(51)" size="3" value="<%=forum_setting(51)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�Ƿ����ü�¼�����Ķ��û�����</U><BR>ֻ���û�����ʱѡ���¼�Ķ������û�����Ч<BR>���������ܽ����Ĳ���ϵͳ��Դ<BR>�����û�������Ա�Ͱ����ɿ��������Ķ��û��б�</td>
<td width="50%" class=tablebody1>  
<input type=radio name="forum_setting(3)" value=0 <%if cint(forum_setting(3))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="forum_setting(3)" value=1 <%if cint(forum_setting(3))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�����й��ģʽ</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(27)" value=0 <%if cint(Forum_Setting(27))=0 then%>checked<%end if%>>��̬���&nbsp;
<input type=radio name="Forum_Setting(27)" value=1 <%if cint(Forum_Setting(27))=1 then%>checked<%end if%>>��������&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�����л�Ա������ʾ��ʽ</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(72)" value=0 <%if cint(Forum_Setting(72))=0 then%>checked<%end if%>>������ϸ&nbsp;
<input type=radio name="Forum_Setting(72)" value=1 <%if cint(Forum_Setting(72))=1 then%>checked<%end if%>>�����&nbsp;
<input type=radio name="Forum_Setting(72)" value=2 <%if cint(Forum_Setting(72))=2 then%>checked<%end if%>>ͷ����ϸ&nbsp;
</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�������Ƿ���ʾ���͹�Ʊ</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(45)" value=0 <%if cint(Forum_Setting(45))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(45)" value=1 <%if cint(Forum_Setting(45))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<th height=25 colspan=2 align=left id=tabletitlelink><a name="setting19"></a><b>��������</b>[<a href="#top">����</a>]</td>
</tr>
<tr> 
<td width="50%" class=tablebody1> <U>�Ƿ�����̳����</U></td>
<td width="50%" class=tablebody1>  
<input type=radio name="Forum_Setting(32)" value=0 <%if cint(Forum_Setting(32))=0 then%>checked<%end if%>>��&nbsp;
<input type=radio name="Forum_Setting(32)" value=1 <%if cint(Forum_Setting(32))=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr bgcolor=<%=Forum_body(3)%>> 
<td width="50%" class=Tablebody1> &nbsp;</td>
<td width="50%" class=Tablebody1>  
<div align="center"> 
<input type="submit" name="Submit" value="�� ��">
</div>
</td>
</tr>
</table>
</form>
<%
end if
end sub

'====================��������===================
sub savebmset()
dim chkedit
dim master_1
dim Forum_copyright
chkedit=false
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>��û��ָ����Ӧ��̳ID�����ܽ��й���"
call dvbbs_error()
exit sub
end if
if isnull(rs(0)) then
Errmsg=Errmsg+"<br>"+"<li>����̳��δ�й���Ա��"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
dim Forum_setting
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>�����Ϊ������ר�á� "
	call dvbbs_error()
else
Forum_Setting=request.Form("Forum_Setting(0)") & "," & request.Form("Forum_Setting(1)") & "," & request.Form("Forum_Setting(2)") & "," & request.Form("Forum_Setting(3)") & "," & request.Form("Forum_Setting(4)") & "," & request.Form("Forum_Setting(5)") & "," & request.Form("Forum_Setting(6)") & "," & request.Form("Forum_Setting(7)") & "," & request.Form("Forum_Setting(8)") & "," & request.Form("Forum_Setting(9)") & "," & request.Form("Forum_Setting(10)") & "," & request.Form("Forum_Setting(11)") & "," & request.Form("Forum_Setting(12)") & "," & request.Form("Forum_Setting(13)") & "," & request.Form("Forum_Setting(14)") & "," & request.Form("Forum_Setting(15)") & "," & request.Form("Forum_Setting(16)") & "," & request.Form("Forum_Setting(17)") & "," & request.Form("Forum_Setting(18)") & "," & request.Form("Forum_Setting(19)") & "," & request.Form("Forum_Setting(20)") & "," & request.Form("Forum_Setting(21)") & "," & request.Form("Forum_Setting(22)") & "," & request.Form("Forum_Setting(23)") & "," & request.Form("Forum_Setting(24)") & "," & request.Form("Forum_Setting(25)") & "," & request.Form("Forum_Setting(26)") & "," & request.Form("Forum_Setting(27)") & "," & request.Form("Forum_Setting(28)") & "," & request.Form("Forum_Setting(29)") & "," & request.Form("Forum_Setting(30)") & "," & request.Form("Forum_Setting(31)") & "," & request.Form("Forum_Setting(32)") & "," & request.Form("Forum_Setting(33)") & "," & request.Form("Forum_Setting(34)") & "," & request.Form("Forum_Setting(35)") & "," & request.Form("Forum_Setting(36)") & "," & request.Form("Forum_Setting(37)") & "," & request.Form("Forum_Setting(38)") & "," & request.Form("Forum_Setting(39)") & "," & request.Form("Forum_Setting(40)") & "," & request.Form("Forum_Setting(41)") & "," & request.Form("Forum_Setting(42)") & "," & request.Form("Forum_Setting(43)") & "," & request.Form("Forum_Setting(44)") & "," & request.Form("Forum_Setting(45)") & "," & request.Form("Forum_Setting(46)") & "," & request.Form("Forum_Setting(47)") & "," & request.Form("Forum_Setting(48)") & "," & request.Form("Forum_Setting(49)") & "," & request.Form("Forum_Setting(50)") & "," & request.Form("Forum_Setting(51)") & "," & request.Form("Forum_Setting(52)") & "," & request.Form("Forum_Setting(53)") & "," & request.Form("Forum_Setting(54)") & "," & request.Form("Forum_Setting(55)") & "," & request.Form("Forum_Setting(56)") & "," & request.Form("Forum_Setting(57)") & "," & request.Form("Forum_Setting(58)") & "," & request.Form("Forum_Setting(59)") & "," & request.Form("Forum_Setting(60)") & "," & request.Form("Forum_Setting(61)") & "," & request.Form("Forum_Setting(62)") & "," & request.Form("Forum_Setting(63)") & "," & request.Form("Forum_Setting(64)") & "," & request.Form("Forum_Setting(65)") & "," & request.Form("Forum_Setting(66)") & "," & request.Form("Forum_Setting(67)") & "," & request.Form("Forum_Setting(68)") & "," & request.Form("Forum_Setting(69)") & "," & request.Form("Forum_Setting(70)") & "," & request.Form("Forum_Setting(71)") & "," & request.Form("Forum_Setting(72)")


Forum_info=Request("ForumName") & "," & Request("ForumURL") & "," & Request("CompanyName") & "," & Request("HostUrl") & "," & Request("SMTPServer") & "," & Request("SystemEmail") & "," & Request("Logo") & "," & Request("Picurl") & "," & Request("Faceurl") & "," & Request("GMT") & "," & Request("emoturl") & "," & Request("userfaceurl") & "," & Request("CreateDate")
Forum_copyright=request("copyright")
sql="update config set Forum_Setting='"&Forum_Setting&"',StopReadme='"&request.Form("StopReadme")&"',Forum_info='"&Forum_info&"',Forum_copyright='"&Forum_copyright&"' where id="&skinid
conn.execute(sql)
response.write "���óɹ���"
end if
end sub

sub editbmcolor()
dim master_1,chkedit
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>��û��ָ����Ӧ��̳ID�����ܽ��й���"
call dvbbs_error()
exit sub
end if
if isnull(rs(0)) then
Errmsg=Errmsg+"<br>"+"<li>����̳��δ�й���Ա��"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>�����Ϊ������ר�á� "
	call dvbbs_error()
else
%>
<form method="POST" action="?action=savebmcolor&boardid=<%=request("boardid")%>">
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
<script>
function color(para_URL){var URL =new String(para_URL)
window.open(URL,'','width=300,height=220,noscrollbars')}
</SCRIPT>
<tr> 
<th height="23" colspan="3" align=left>��̳��������&nbsp;���� <a href="javascript:color('color.asp')"><font color=#FFFFFF>����</font></a> ʹ��������ɫʰȡ�����޸��������ñ���߱�һ����ҳ֪ʶ��</th>
</tr>
<tr> 
<td width="45%" class=tablebody1>��̳BODY��ǩ<br>
����������̳���ı�����ɫ���߱���ͼƬ��</td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<input type="text" name="ForumBody" size="50" value="<%=Forum_body(11)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>Body��CSS����<BR>�ɶ�������Ϊ��ҳ������ɫ��������������߿��<BR><FONT COLOR="red">���¶����У�����CSS����ľͲ�һ����������ɫ�ϵ������ˣ��������������е����塢������</FONT></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="iebarcolor" cols="50" rows="3"><%=Forum_body(25)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>��̳�����ܵ�CSS����<BR>�ɶ�������Ϊ����������ɫ����ʽ��<BR></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="ForumCssLink" cols="50" rows="3"><%=Forum_body(6)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>��̳����ܵ�CSS����<BR>����Ϊ��̳�ܵı���壬Ϊһ����ķ�����ã���ҳ���г��˵�������Ϣ�б���Ĳ��֣�β����ͷ���ȣ����ɶ�������Ϊ������������ɫ����ʽ��<BR></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="ForumCssTD" cols="50" rows="3"><%=Forum_body(10)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>�����˵����CSS����(Logo & Banner�Ϸ�)<BR><FONT COLOR="#000099">���ã�Class=TopDarkNav</FONT></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="NavDarkcolor" cols="50" rows="3"><%=Forum_body(24)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>�����˵����CSS����(Logo & Banner�·�)<BR><FONT COLOR="#000099">���ã�Class=TopLighNav</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Navlighcolor" cols="50" rows="3"><%=Forum_body(23)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>�����˵����CSS����(�����˵�)<BR><FONT COLOR="#000099">���ã�Class=TopLighNav1</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Navlighcolor1" cols="50" rows="3"><%=Forum_body(26)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>�����˵����CSS����(Logo & banner����)<BR><FONT COLOR="#000099">���ã�Class=TopLighNav2</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Navlighcolor2" cols="50" rows="3"><%=Forum_body(28)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>���߿���ɫ����<br>
�����涨��һ�Ͷ�CSS������Ʋ����ĵط�����ñ��ֺͱ߿�CSS����һ��ͬ������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(27)%>"></td>
<td width="50%" class=tablebody1> 
<input type=text name="btablebackcolor" value="<%=Forum_body(27)%>" size=35>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>���߿��壩CSS����һ<br>
һ��ҳ�棬�ɶ�������Ϊ����񱳾�������ͼ���߿򡢿�ȵ�<BR><FONT COLOR="#000099">���ã�Class=TableBorder1</font></td>
<td width="5%"  class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Tablebackcolor" cols="50" rows="3"><%=Forum_body(0)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>���߿��壩CSS�����<br>
����̳�������߿򼰲���ҳ�棬�ɶ�������Ϊ����񱳾�������ͼ���߿򡢿�ȵ�<BR><FONT COLOR="#000099">���ã�Class=TableBorder2</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="aTablebackcolor" cols="50" rows="3"><%=Forum_body(1)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>������CSS����һ�������<br>
һ��Ϊ���ı���״̬����ɫ���߱����ϵĶ��壬�ɶ�������Ϊ����������ͼ�����弰����ɫ��<BR><FONT COLOR="#000099">���ã�th</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Tabletitlecolor" cols="50" rows="3"><%=Forum_body(2)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>��������������CSS���壨�����<br>
���������Ϊ���������������ر����ڱ����������ӵ���ɫ���������ɫһ�������������ｫ����Ĭ�ϵ�������ɫ<BR>ע�⣺�����#TableTitleLink����ȥ�������Խ�������Ŀ�ֿ�д���ﵽ���ӵĲ�ͬЧ��<BR><FONT COLOR="#000099">���ã�id=TableTitleLink</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="TableTitleLink" cols="50" rows="3"><%=Forum_body(7)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>������CSS�������ǳ������<br>
Ϊ��������״̬����ɫ���߱����ϵĶ��壬����ҳ����̳���������ɶ�������Ϊ����������ͼ�����弰����ɫ��<BR><FONT COLOR="#000099">���ã�Class=TableTitle2</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="aTabletitlecolor" cols="50" rows="3"><%=Forum_body(3)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>�����CSS����һ<BR>�ɶ�������Ϊ����������ͼ�����弰����ɫ��<BR><FONT COLOR="#000099">���ã�Class=TableBody1</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="Tablebodycolor" cols="50" rows="3"><%=Forum_body(4)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>�������ɫ��(1��2��ɫ����ʾ�д���)<BR>�ɶ�������Ϊ����������ͼ�����弰����ɫ��<BR><FONT COLOR="#000099">���ã�Class=TableBody2</font></td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<textarea name="aTablebodycolor" cols="50" rows="3"><%=Forum_body(5)%></textarea>
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>�����</td>
<td width="5%" class=tablebody1></td>
<td width="50%" class=tablebody1> 
<input type="text" name="TableWidth" size="35" value="<%=Forum_body(12)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>��������������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(8)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="AlertFontColor" size="35" value="<%=Forum_body(8)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>��ʾ���ӵ�ʱ��������ӣ�ת�����ӣ��ظ��ȵ���ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(9)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="ContentTitle" size="35" value="<%=Forum_body(9)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>��ҳ������ɫ<BR>���������</td>
<td width="5%" bgcolor="<%=Forum_body(14)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="BoardLinkColor" size="35" value="<%=Forum_body(14)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>һ���û�����������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(15)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="user_fc" size="35" value="<%=Forum_body(15)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>һ���û������ϵĹ�����ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(16)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="user_mc" size="35" value="<%=Forum_body(16)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>��������������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(17)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="bmaster_fc" size="35" value="<%=Forum_body(17)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>���������ϵĹ�����ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(18)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="bmaster_mc" size="35" value="<%=Forum_body(18)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>����Ա����������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(19)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="master_fc" size="35" value="<%=Forum_body(19)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>����Ա�����ϵĹ�����ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(20)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="master_mc" size="35" value="<%=Forum_body(20)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>�������������ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(21)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="vip_fc" size="35" value="<%=Forum_body(21)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody1>��������ϵĹ�����ɫ</td>
<td width="5%" bgcolor="<%=Forum_body(22)%>"></td>
<td width="50%" class=tablebody1> 
<input type="text" name="vip_mc" size="35" value="<%=Forum_body(22)%>">
</td>
</tr>
<tr> 
<td width="45%" class=tablebody2>&nbsp;</td>
<td width="5%" class=tablebody2></td>
<td width="50%" class=tablebody2> 
<input type="submit" name="Submit" value="�� ��">
</td>
</tr>
</table>
</form>
<%
end if
end sub
sub savebmcolor()
dim Forum_body,master_1
dim chkedit
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>��û��ָ����Ӧ��̳ID�����ܽ��й���"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>�����Ϊ������ר�á� "
	call dvbbs_error()
else
Forum_body=request.form("Tablebackcolor") & "|||" & request.form("aTablebackcolor") & "|||" & request.form("Tabletitlecolor") & "|||" & request.form("aTabletitlecolor") & "|||" & request.form("Tablebodycolor") & "|||" & request.form("aTablebodycolor") & "|||" & request.form("ForumCssLink") & "|||" & request.form("TableTitleLink") & "|||" & request.form("AlertFontColor") & "|||" & request.form("ContentTitle") & "|||" & request.form("ForumCssTD") & "|||" & request.form("ForumBody") & "|||" & request.form("TableWidth") & "|||" & request.form("BodyFontColor") & "|||" & request.form("BoardLinkColor") & "|||" & request.form("user_fc") & "|||" & request.form("user_mc") & "|||" & request.form("bmaster_fc") & "|||" & request.form("bmaster_mc") & "|||" & request.form("master_fc") & "|||" & request.form("master_mc") & "|||" & request.form("vip_fc") & "|||" & request.form("vip_mc") & "|||" & request.form("NavLighColor") & "|||" & request.form("NavDarkColor") & "|||" & request.form("IEbarColor") & "|||" & request.form("Navlighcolor1") & "|||" & request.form("btablebackcolor") & "|||" & request.form("Navlighcolor2")
response.write skinid

sql = "update config set Forum_body='"&checkstr(Forum_body)&"' where id="&skinid
conn.execute(sql)

response.write "���óɹ���"
end if
end sub

sub editbmads()
dim master_1,chkedit
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>��û��ָ����Ӧ��̳ID�����ܽ��й���"
call dvbbs_error()
exit sub
end if
if isnull(rs(0))	then
Errmsg=Errmsg+"<br>"+"<li>����̳��δ�й���Ա��"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>�����Ϊ������ר�á� "
	call dvbbs_error()
else
%>
<form method="POST" action="?action=savebmads&boardid=<%=request("boardid")%>">
<TABLE cellPadding=1 cellSpacing=1 class=tableborder1 align=center style="width:96%;word-break:break-all;">
<tr> 
<th height="23" colspan="2" class="tableHeaderText"><b>��̳�������</b>����Ϊ���÷���̳�����Ƿ���̳��ҳ��棬����ҳ��Ϊ������ʾҳ�棩</th>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��ҳ����������</B></font></td>
<td width="60%" class="Tablebody1"> 
<textarea name="index_ad_t" cols="50" rows="3"><%=Forum_ads(0)%></textarea>
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��ҳβ��������</B></font></td>
<td width="60%" class="Tablebody1"> 
<textarea name="index_ad_f" cols="50" rows="3"><%=Forum_ads(1)%></textarea>
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>������ҳ�������</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type=radio name="index_moveFlag" value=0 <%if Forum_ads(2)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="index_moveFlag" value=1 <%if Forum_ads(2)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ�������ͼƬ��ַ</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="MovePic" size="35" value="<%=Forum_ads(3)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ����������ӵ�ַ</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="MoveUrl" size="35" value="<%=Forum_ads(4)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ�������ͼƬ���</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="move_w" size="3" value="<%=Forum_ads(5)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ�������ͼƬ�߶�</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="move_h" size="3" value="<%=Forum_ads(6)%>">&nbsp;����
</td>
</tr>
<input type=hidden name="Board_moveFlag" value=0>
<tr> 
<td width="40%" class="Tablebody1"><B>������ҳ���¹̶����</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type=radio name="index_fixupFlag" value=0 <%if Forum_ads(13)=0 then%>checked<%end if%>>�ر�&nbsp;
<input type=radio name="index_fixupFlag" value=1 <%if Forum_ads(13)=1 then%>checked<%end if%>>��&nbsp;
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ���¹̶����ͼƬ��ַ</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixupPic" size="35" value="<%=Forum_ads(8)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ���¹̶�������ӵ�ַ</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixupUrl" size="35" value="<%=Forum_ads(9)%>">
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ���¹̶����ͼƬ���</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixup_w" size="3" value="<%=Forum_ads(10)%>">&nbsp;����
</td>
</tr>
<tr> 
<td width="40%" class="Tablebody1"><B>��̳��ҳ���¹̶����ͼƬ�߶�</B></font></td>
<td width="60%" class="Tablebody1"> 
<input type="text" name="fixup_h" size="3" value="<%=Forum_ads(11)%>">&nbsp;����
</td>
</tr>
<input type=hidden name="Board_fixupFlag" value=0>
<tr> 
<td width="40%" class="Tablebody1">&nbsp;</td>
<td width="60%" class="Tablebody1"> 
<div align="center"> 
<input type="submit" name="Submit" value="�� ��">
</div>
</td>
</tr>
</table>
</form>
<%
end if
end sub
sub savebmads()
dim master_1
dim chkedit
dim Forum_adsinfo
set rs=conn.execute("select boardmaster from board where boardid="&request("boardid"))
if rs.eof and rs.bof then
Errmsg=Errmsg+"<br>"+"<li>��û��ָ����Ӧ��̳ID�����ܽ��й���"
call dvbbs_error()
exit sub
end if
master_1=split(rs(0),"|")
if Board_Setting(35)=1 then
chkedit=true
else
	if Board_Setting(34)=0 then
	chkedit=false
	elseif Board_Setting(34)=1 and membername=master_1(0) then
	chkedit=true
	else
	chkedit=false
	end if
end if
if master then
chkedit=true
end if
if chkedit=false then
	Errmsg=Errmsg+"<br>"+"<li>�����Ϊ������ר�á� "
	call dvbbs_error()
else

Forum_adsinfo=request("index_ad_t") & "$" & request("index_ad_f") & "$" & request("index_moveFlag") & "$" & request("MovePic") & "$" & request("MoveUrl") & "$" & request("move_w") & "$" & request("move_h") & "$" & request("Board_moveFlag") & "$" & request("fixupPic") & "$" & request("FixupUrl") & "$" & request("Fixup_w") & "$" & request("Fixup_h") & "$" & request("Board_fixupFlag") & "$" & request("index_fixupFlag")
'response.write Forum_ads
sql = "update config set Forum_ads='"&CheckStr(Forum_adsinfo)&"' where id="&skinid
conn.execute(sql)
	response.write skinid&"���óɹ���"
end if
end sub

sub addmaster(s,o,n)
dim arr,pw,oarr
dim classname,titlepic
set rs=conn.execute("select title from usergroups where usergroupid=3")
classname=rs(0)
set rs=conn.execute("select titlepic from usertitle where usergroupid=3 order by Minarticle desc")
if not (rs.eof and rs.bof) then
titlepic=rs(0)
end if
randomize
pw=Cint(rnd*9000)+1000
arr=split(s,"|")
oarr=split(o,"|")
set rs=server.createobject("adodb.recordset")
for i=0 to Ubound(arr)
sql="select * from [user] where username='"& arr(i) &"'"
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	rs.addnew
	rs("username")=arr(i)
	rs("userpassword")=md5(pw)
	rs("userclass")=classname
	rs("UserGroupID")=3
	rs("titlepic")=titlepic
	rs("userWealth")=100
	rs("userep")=30
	rs("usercp")=30
	rs("userisbest")=0
	rs("userdel")=0
	rs("userpower")=0
	rs("lockuser")=0
	rs.update
	str=str&"������������û���<b>" &arr(i) &"</b> ���룺<b>"& pw &"</b><br><br>"
else
	if rs("UserGroupID")>3 then
		rs("userclass")=classname
		rs("UserGroupID")=3
		rs("titlepic")=titlepic
		rs.update
	end if
end if
rs.close
next

'�ж�ԭ���������������Ƿ񻹵��ΰ�������û�е����򳷻����û�ְλ
if n=1 then
dim iboardmaster
dim UserGrade,article
iboardmaster=false
for i=0 to ubound(oarr)
	set rs=conn.execute("select boardmaster from board")
	do while not rs.eof
		if instr("|"&trim(rs("boardmaster"))&"|","|"&trim(oarr(i))&"|")>0 then
			iboardmaster=true
			exit do
		end if
	rs.movenext
	loop
	if not iboardmaster then
	set rs=conn.execute("select userid,UserGroupID,article from [user] where username='"&trim(oarr(i))&"'")
	if not (rs.eof and rs.bof) then
		if rs(1)>2 then
		if not isnumeric(rs(2)) then article=0
		set UserGrade=conn.execute("select top 1 usertitle,titlepic,UserGroupID from usertitle where Minarticle<="&rs(2)&" and not MinArticle=-1 order by MinArticle desc,usertitleid")
		if not (UserGrade.eof and UserGrade.bof) then
			conn.execute("update [user] set UserGroupID="&UserGrade(2)&",titlepic='"&UserGrade(1)&"',userclass='"&UserGrade(0)&"' where userid="&rs(0))
		end if
		end if
	end if
	end if
	iboardmaster=false
next
end if
set rs=nothing
end sub
%>

