<!--#include FILE="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!-- #include file="inc/chkinput.asp" -->
<%
stats="�����ϴ�����"
call nav()

Dim TopicCount
Dim Pcount,endpage,star,page_count
if request("star")="" or not isnumeric(request("star")) then
	star=1
else
	star=clng(request("star"))
end if

if not founduser then
  	errmsg=errmsg+"<br>"+"<li>��û��<a href=login.asp target=_blank>��¼</a>��"
	founderr=true
end if
if founderr=true then
	call head_var(2,0,"","")
	call dvbbs_error()
else
	call head_var(0,0,membername & "�Ŀ������","usermanager.asp")
	select case request("action")
	case "edit"
		call edit()
	case "fsave"
		call filesave()
	case "fadd"
		call addnew()	
	case "fsnew"
		call savenew()
	case "fdel"
		call fdel()
	case "alldel"
		call alldel()
	case else
		call main()
	end select
	if founderr then call dvbbs_error()
end if
call activeonline()
call footer()

sub main()
dim sname
dim stype
dim searchsql
stype=request("Stype")

TopicCount=filenum("")
if stype="" or not  isnumeric(stype)  then
	sname="�����ļ�"
	searchsql=""
	TopicCount=filenum("")
else
	select case stype
	case 1
	sname="ͼƬ��"
	searchsql="F_Type=1 and"
	case 2
	sname="FLASH��"
	searchsql="F_Type=2 and"
	case 3
	sname="���ּ�"
	searchsql="F_Type=3 and"
	case 4
	sname="��Ӱ��"
	searchsql="F_Type=4 and"
	case 0
	sname="�ļ���"
	searchsql="F_Type=0 and"
	case else
	sname="�����ļ�"
	searchsql=""
	end select
	TopicCount=filenum(clng(stype))
end if
%>
<br>
<!--#include file="z_controlpanel.asp"-->
<br>
	<table cellpadding=3 cellspacing=1  align=center  class=tableborder2 >
	<form action="<%=ScriptName%>" method=get>
	<tr><td><a href=?action=fadd >�������ļ����</a>��<a href=show.asp?username=<%=trim(membername)%> >�۸���չ����</a>��Ŀǰ����<%=filenum("")%>���ļ�
<% if 	stype<>"" and isnumeric(stype) then %>
	�����С�<%=sname%>����<font color=<%=Forum_body(8)%>><%=filenum(clng(stype))%></font>����
<%end if
%>
	</td>
	<td width=200 align=right>
	<select name=Stype onchange='javascript:submit()'>
	<option value=all>�鿴�ļ�����
	<option value=0>�ļ���
	<option value=1>ͼƬ��
	<option value=2>FLASH��
	<option value=3>���ּ�
	<option value=4>��Ӱ��
	<option value=5>�����ļ�
	</select>
	</td></tr></form>
	</table>
	<script language="JavaScript">
<!--
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
//-->
</script>
<%
dim F_Type,F_typename
'---------------------
response.write"<table cellpadding=3 cellspacing=1 align=center class=tableborder1>"&_
              "<form action="""&ScriptName&"?action=fdel""  method=post name=myfile><tr>"&_
              "<th colspan=7 height=25 align=left>-=> �����ϴ��ļ�</th></tr>"&_
            "<tr>"&_
                "<td align=center valign=middle width=30 class=tabletitle2><b>����</b></td>"&_
                "<td align=center valign=middle width=* class=tabletitle2><b>�ļ�˵��</b></td>"&_
                "<td align=center valign=middle width=60 class=tabletitle2><b>��С</b></td>"&_
				"<td align=center valign=middle width=60 class=tabletitle2><b>����/���</b></td>"&_
                "<td align=center valign=middle width=120 class=tabletitle2><b>����</b></td>"&_
                "<td align=center valign=middle width=60 class=tabletitle2><b>����</b></td>"&_
				"<td align=center valign=middle width=120 class=tabletitle2><b>����</b></td>"&_
            "</tr>"
'---------------------
set rs=server.createobject("ADODB.Recordset")
rs.open "select * from [DV_Upfile] where  "&searchsql&"  F_UserID="&userid&" order by F_ID desc",conn,1,1
if rs.eof and rs.bof then
	response.write"<tr><td class=tablebody1 align=center valign=middle colspan=7>�����ļ�����û���κ����ݡ�</td></tr>"
else
	if TopicCount mod Cint(Forum_Setting(11))=0 then
		Pcount= TopicCount \ Cint(Forum_Setting(11))
	else
		Pcount= TopicCount \ Cint(Forum_Setting(11))+1
	end if
	if star > Pcount then star = Pcount
	if star < 1 then star = 1
	rs.PageSize = Cint(Forum_Setting(11))
	rs.AbsolutePage=star
	page_count=0
	do while not rs.eof and page_count < Cint(Forum_Setting(11))
F_Type=rs("F_Type")
if F_Type=1 then
	F_typename="ͼƬ��"
elseif F_Type=2 then
	F_typename="FLASH��"
elseif F_Type=3 then
	F_typename="���ּ�"
elseif F_Type=4 then
	F_typename="��Ӱ��"
else
	F_typename="�ļ���"
end if
response.write	"<tr>"
response.write	"<td align=center valign=middle  class=tablebody2><img src=""images/files/"&rs("F_FileType")&".gif"" border=0></td>"
response.write	"<td align=left valign=middle  class=tablebody1><a href=""fileshow.asp?boardid="&rs("F_BoardID")&"&id="&rs("F_ID") &"""  target=""_blank"" >"
if rs("F_Readme")<>"" or not isnull(rs("F_Readme")) then
if len(rs("F_Readme"))>26 then
response.write ""&left(htmlencode(replace(rs("F_Readme"),chr(10)," ")),26)&"..."
else
response.write Server.htmlencode(rs("F_Readme"))
end if
end if
response.write"</a></td>"
response.write	"<td align=right valign=middle  class=tablebody2>"&rs("F_FileSize")&" <b>B</b></td>"
response.write	"<td align=center valign=middle  class=tablebody1>"&rs("F_DownNum")&"/"&rs("F_ViewNum")&"</td>"
response.write	"<td align=left valign=middle  class=tablebody2>"&rs("F_AddTime")&"</td>"
response.write	"<td align=center valign=middle  class=tablebody1>"&F_typename&"</td>"
response.write	"<td align=center valign=middle  class=tablebody2>"

if GroupSetting(48)=1 then
	if instr(rs("F_Filename"),"/") and cint(rs("F_Flag"))=1 then
		response.write	" <input type=checkbox name=delid value="""&rs("F_ID")&"""> "
	else
		response.write	" <input type=checkbox  value="""" disabled > "
	end if
	response.write"<a href=?action=edit&editid="&rs("F_ID")&"  >�༭</a>  | <a href=fileshow.asp?action=send&id="&rs("F_ID")&"  >����</a>"
else
	response.write	" ���� "
end if
response.write	"</td>"
response.write	"</tr>"
page_count = page_count + 1
		rs.movenext
		loop
end if
rs.close
set rs=nothing
if GroupSetting(48)=1 then
	response.write "<tr><td  colspan=7 align=right class=tablebody2>��ѡ�����ɾ�����ļ���<input type=checkbox name=chkall value=on onclick=""CheckAll(this.form)"">ȫѡ <input type=submit name=Submit value=ִ��  onclick=""{if(confirm('��ȷ��ִ�д˲�����?')){this.document.myfile.submit();return true;}return false;}""></td></tr>"
end if
response.write"</form></table>"

call list()
end sub

'��ҳ����
sub list()
response.write "<table cellpadding=0 cellspacing=3 border=0 width="&Forum_body(12)&" align=center><form method=post action="&ScriptName&" ><tr><td valign=middle nowrap align=left>"&stats&"�ļ���<b>"&TopicCount&"</b>������<b><font color=red>"&Pcount&"</font></b>ҳ</td><td valign=middle  align=right>"
call DispPageNum(clng(star),Pcount,"""?stype="&request("Stype")&"&star=","""")
response.write " ת��:<input type=text name=star size=3 maxlength=10  value='"& star &"'><input type=submit value=Go  id=button1 name=button1 >"     
response.write "</td></tr></form></table>"
end sub

'�༭�ļ�
sub edit()
dim editid
dim F_Type,F_typename,filename,chefile,con,body
dim F_Username,postid,F_rootid,F_bbsid
dim myurl

myurl=false
editid=trim(request("editid"))
if not isnumeric(editid) or isnull(editid) then 
	errmsg=errmsg+"<br>"+"<li>ִ�е����ݲ����ڡ�"
	founderr=true
	exit sub
end if
set rs=conn.execute("select * from [DV_Upfile] where F_ID="&editid)
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>�༭���ļ��Ѳ����ڡ�"
	founderr=true
	exit sub
else
	F_Username=rs("F_Username")
	F_Type=rs("F_Type")
	filename=rs("F_Filename")
	con=rs("F_Readme")

	if instr(filename,"/")=0 then
	myurl=true
	filename="UploadFile/"&filename
	end if
	if F_Type=1  then
	F_typename="<img src='"&filename&"'  style='border: 1 solid #000000' width=80 height=80 >"
	else
	F_typename=rs("F_FileType")&"�ļ�"
	end if

	if con<>"" then
	body=replace(con,"<br>",chr(13))
	body=replace(body,"&nbsp;","")
	body=body+chr(13)
	end if
	
%>
 <table cellpadding=3 cellspacing=1 align=center class=tableborder1>
 <form action="?action=fsave" method=post >
	<tr>
	<th width="100%" height="26" colspan="2" align=left>�༭<%=F_Username%>�ϴ��ļ�</th>
	</tr>
	<tr>
	<td width="15%" height="26" class=tablebody2 align=right>�༭���ļ���</td>
	<td width="85%"  height="26" class=tablebody1 align=left>
	<a href="fileshow.asp?boardid=<%=rs("F_BoardID")%>&id=<%=rs("F_ID") %>"  target="_blank" ><%=F_typename%></a>
	</td>
	</tr>
<% if myurl=false then %>
	<tr>
	<td width="15%" height="26" class=tablebody2 align=right>�ļ���ַ��</td>
	<td width="85%"  height="26" class=tablebody1 align=left><input type=text name="fileurl" size="80%"  value="<%=rs("F_Filename")%>">
	</td>
	</tr>
<% else 
	if instr(rs("F_AnnounceID"),"|") then
		postid=split(rs("F_AnnounceID"),"|")
		F_rootid=postid(0)
		F_bbsid=postid(1)
	end if
%>
	<tr>
	<td valign=middle align=right  class=tablebody2 >�ļ�����ϵ����</td>
	<td width="100%" height="26" class=tablebody1>
	<% if (master or superboardmaster or boardmaster) and GroupSetting(48)=1 then%>
	��̳ID:<input type=text name="F_BoardID" value="<%= rs("F_BoardID") %>" size=8>&nbsp;
	����ID:<input type=text name="F_AnnounceID" value="<%= rs("F_AnnounceID") %>">&nbsp;<font color=<%=Forum_body(8)%>>�����ļ�����������δ�иĶ�,�벻Ҫ�޸����ϵ������</font>
	<%else%>
	<input type=hidden name="F_BoardID" value="<%= rs("F_BoardID") %>" >
	<input type=hidden name="F_AnnounceID" value="<%= rs("F_AnnounceID") %>">
	<%end if%>
	<a href="dispbbs.asp?Boardid=<%=rs("F_BoardID")%>&ID=<%=F_Rootid%>&replyID=<%=F_bbsid%>&skin=1" target="_blank" title="����������">[<font color=<%=Forum_body(8)%>>�鿴�������</font>]</a>
	</td>
    	</tr>
<%end if%>
	<tr>
	<td width="15%" height="26" class=tablebody2 align=right>�ļ���˵����</td>
	<td width="85%"  height="26" class=tablebody1 align=left>
	<textarea name="F_Readme" rows=6 cols=80%  wrap=VIRTUAL><%response.write Server.htmlencode(body)%></textarea>
	<input type=hidden name="saveid" value="<%= rs("F_ID") %>">�������ó��������������ֻ򣲣������ַ���
	</td>
	</tr>
	 <%if myurl=true then %>
	<tr>
	<td valign=middle align=right  class=tablebody2 >�ļ�����ɾ����</td>
	<td valign=middle class=tablebody1>
	  <input type=radio name=Fflag value=0 <%if rs("F_Flag")=0 then %> checked<%end if%>>��&nbsp;
	  <input type=radio name=Fflag value=2 <%if rs("F_Flag")=2 then %> checked<%end if%>>��
	</td>
	</tr>
	<%else%>
	<tr>
	<td valign=middle align=right  class=tablebody2 >�ļ�����ɾ����</td>
	<td valign=middle class=tablebody1>
	  <input type=radio name=Fflag value=1 <%if rs("F_Flag")=1 then %> checked<%end if%>>��&nbsp;
	  <input type=radio name=Fflag value=2 <%if rs("F_Flag")=2 then %> checked<%end if%>>��
	</td>
	</tr>
	<%end if%>
	<tr>
	<td width="100%" height="26" colspan="2" class=tablebody1>
	  ˵����<li>�������ļ���ɾ�����򵱺��и��ļ������ӱ�ɾ�����ļ��п����౻ɾ����
	  <li>�ļ�˵�����ó��������������ֻ򣲣������ַ���
	</td>
    	</tr>
	<tr>
	<th width="100%" height="26" colspan="2">
	<input type=Submit value="����" name=Submit>&nbsp;
	</th>
	</tr>
	</form>
  </table>
<%
end if 
rs.close
set rs=nothing
end sub 


'�����޸�
sub filesave()
if GroupSetting(48)=0 then
	errmsg=errmsg+"<br>"+"<li>��û����̳�ļ������Ȩ��,�������Ա��ϵ��"
	founderr=true
	exit sub
end if
dim saveid,F_Readme,Fflag,fileurl
dim F_BoardID,F_AnnounceID

F_BoardID=checkStr(trim(request.form("F_BoardID")))
F_AnnounceID=checkStr(trim(request.form("F_AnnounceID")))
F_Readme=checkStr(trim(request.form("F_Readme")))
saveid=trim(request.form("saveid"))
Fflag=trim(request.form("Fflag"))
if not isnumeric(Fflag) or not isnumeric(F_BoardID) then 
	errmsg=errmsg+"<br>"+"<li>ִ�е����ݲ�����,�򲻺Ϸ�Ҫ��"
	founderr=true
	exit sub
end if
if not isnumeric(saveid) or isnull(saveid) then 
	errmsg=errmsg+"<br>"+"<li>ִ�е����ݲ����ڡ�"
	founderr=true
	exit sub
end if

if  F_Readme="" or isnull(F_Readme) then 
	errmsg=errmsg+"<br>"+"<li>˵�����ݲ���Ϊ�ա�"
	founderr=true
	exit sub
end if

if strLength(F_Readme)>250 then
   		ErrMsg=ErrMsg+"<Br>"+"<li>˵�����ݲ��ô���" & 250 & "bytes"
   		FoundErr=true
		exit sub
end if
fileurl=checkStr(trim(request.form("fileurl")))
if fileurl<>"" then
	if instr(fileurl,"/")=0 or instr(fileurl,"://")=0  or instr(fileurl,".")=0 then
   		ErrMsg=ErrMsg+"<Br>"+"<li>����д�������ļ����ӵ�ַ��"
   		FoundErr=true
		exit sub
	end if
	conn.execute("update [DV_Upfile]  set F_Filename='"&fileurl&"',F_Readme='"&F_Readme&"' ,F_Flag="&Fflag&" where  F_ID="&saveid)
else
	conn.execute("update [DV_Upfile]  set F_BoardID="&F_BoardID&",F_AnnounceID='"&F_AnnounceID&"',F_Readme='"&F_Readme&"' ,F_Flag="&Fflag&" where  F_ID="&saveid)
end if
	sucmsg=sucmsg+"<br>"+"<li><b>�༭�ɹ���<a href=myfile.asp� >[���ع���ҳ��]</a>"
	call dvbbs_suc()
end sub

'�����ļ�
sub addnew()
%>
 <table cellpadding=3 cellspacing=1 align=center class=tableborder1>
 <form action="?action=fsnew" method=post >
    <tr>
      <th width="100%" height="26" colspan="2" align=left>�����Ҹ����ļ�</th>
    </tr>
	    <tr>
      <td width="15%" height="26" class=tablebody2 align=right>�������ļ���</td>
      <td width="85%"  height="26" class=tablebody1 align=left>
	  <input type=text name="fileurl" size="80%"  value="http://"> ������д��ɵ��ļ����ӵ�ַ��
	</td>
    </tr>
    <tr>
      <td width="15%" height="26" class=tablebody2 align=right>�ļ���˵����</td>
      <td width="85%"  height="26" class=tablebody1 align=left>
	  <textarea name="F_Readme" rows=6 cols=80%  wrap=VIRTUAL value=""></textarea> �����ó��������������ֻ򣲣������ַ���
	</td>
    </tr>
	<tr>
	 <td valign=middle align=right  class=tablebody2 >�ļ������ͣ�</td>
      <td valign=middle class=tablebody1>
	  <select name="filetype" size=1>
<option value="">�ļ�����</option>
<%
dim iupload
iupload="�ļ���|ͼƬ��|FLASH��|���ּ�|��Ӱ��"
iupload=split(iupload,"|")
for i=0 to ubound(iupload)
response.write "<option value="&i&">"&iupload(i)&"</option>"
next
%>
</select>
	</td>
	</tr>
	 <td valign=middle align=right  class=tablebody2 >�ļ�����ɾ����</td>
      <td valign=middle class=tablebody1>
	  <input type=radio name=Fflag value=1  checked >��&nbsp;
	  <input type=radio name=Fflag value=2  >��
	</td>
	</tr>
	<tr>
		<td width="100%" height="26" colspan="2" class=tablebody1>
	  ˵����<li>�������ļ���ɾ����������ʱ���Զ������
	  <li>�ļ�˵�����ó��������������ֻ򣲣������ַ���
	  <li>���ãȣԣԣУ�������ƣԣУ�������ȣԣԣУӣ�������ͷ���������ӵ�ַ��
		</td>
    </tr>
    <tr>
		<th width="100%" height="26" colspan="2">
	  <input type=Submit value="����" name=Submit>&nbsp;<INPUT TYPE="reset" value="���">
		</th>
    </tr>
	</form>
  </table>
  <%
end sub 

'���������ļ�
sub savenew()
if GroupSetting(48)=0 then
	errmsg=errmsg+"<br>"+"<li>��û����̳�ļ������Ȩ��,�������Ա��ϵ��"
	founderr=true
	exit sub
end if
dim F_Readme,fileurl,filetype,fileExt,fileExt_a,filename
dim F_Type,F_Flag
F_Readme=checkStr(trim(request.form("F_Readme")))
fileurl=checkStr(trim(request.form("fileurl")))
filetype=trim(request.form("filetype"))
F_Flag=trim(request.form("Fflag"))

if  fileurl="" or isnull(fileurl) then 
	errmsg=errmsg+"<br>"+"<li>�ļ����Ӳ���Ϊ�ա�"
	founderr=true
	exit sub
else
	if strLength(fileurl)>250 then
   		ErrMsg=ErrMsg+"<Br>"+"<li>�ļ����Ӳ��ô���" & 250 & "bytes"
   		FoundErr=true
		exit sub
	end if
end if

if  F_Readme="" or isnull(F_Readme) then 
	errmsg=errmsg+"<br>"+"<li>˵�����ݲ���Ϊ�ա�"
	founderr=true
	exit sub
end if
if strLength(F_Readme)>250 then
   		ErrMsg=ErrMsg+"<Br>"+"<li>˵�����ݲ��ô���" & 250 & "bytes"
   		FoundErr=true
		exit sub
end if
if not isnumeric(F_Flag)  then 
	errmsg=errmsg+"<br>"+"<li>ִ�е����ݲ����ڡ�"
	founderr=true
	exit sub
end if

	if instr(fileurl,"/")=0 or instr(fileurl,"://")=0  or instr(fileurl,".")=0 then
   		ErrMsg=ErrMsg+"<Br>"+"<li>����д�������ļ����ӵ�ַ��"
   		FoundErr=true
		exit sub
		else
		filename=split(fileurl,"/")
		fileExt=lcase(filename(ubound(filename)))
	end if

	fileExt_a=split(fileExt,".")
	fileExt=lcase(fileExt_a(ubound(fileExt_a)))
	
	if fileEXT="asp" and fileEXT="asa" and fileEXT="aspx" then
	   	ErrMsg=ErrMsg+"<Br>"+"<li>�ļ���ʽ��֧�֣���������д��"
   		FoundErr=true
		exit sub
	end if
	
	if lcase(fileExt)="gif" or lcase(fileExt)="jpg" or lcase(fileExt)="jpeg" or lcase(fileExt)="bmp" or lcase(fileExt)="png" then
	F_Type=1
	elseif lcase(fileExt)="swf" or lcase(fileExt)="swi" then
	F_Type=2
	elseif lcase(fileExt)="mid" or lcase(fileExt)="wav" or lcase(fileExt)="mp3" or lcase(fileExt)="rmi" or lcase(fileExt)="cda" then
	F_Type=3
	elseif lcase(fileExt)="avi" or lcase(fileExt)="wov" or lcase(fileExt)="asf" or lcase(fileExt)="mpg" or lcase(fileExt)="mpeg" or lcase(fileExt)="ra" or lcase(fileExt)="ram" then
	F_Type=4
	else
	F_Type=0
	end if
	BoardID=0
conn.execute("insert into dv_upfile (F_BoardID,F_UserID,F_Username,F_Filename,F_FileType,F_Type,F_Readme,F_Flag ) values ("&BoardID&","&UserID&",'"&membername&"','"&replace(fileurl,"|","")&"','"&replace(fileExt,".","")&"',"&F_Type&",'"&F_Readme&"',"&F_Flag&" )")
	sucmsg=sucmsg+"<br>"+"<li><b>�༭�ɹ���<a href=myfile.asp >[���ع���ҳ��]</a>"
	call dvbbs_suc()
end sub

'ɾ���ļ�
sub fdel()
dim delid
if GroupSetting(48)=0 then
	errmsg=errmsg+"<br>"+"<li>��û����̳�ļ������Ȩ��,�������Ա��ϵ��"
	founderr=true
	exit sub
end if
delid=replace(request("delid"),"'","")
delid=replace(delid,";","")
delid=replace(delid,"--","")
delid=replace(delid,")","")
if delid="" or isnull(delid) then
Errmsg=Errmsg+"<li>"+"��ѡ����ز�����"
founderr=true
exit sub
else
	if  master or superboardmaster or boardmaster  then
	conn.execute("delete from DV_Upfile where   F_ID in ("&delid&")")
	else
	conn.execute("delete from DV_Upfile where F_Flag=1 and  F_ID in ("&delid&")")
	end if
	sucmsg=sucmsg+"<br>"+"<li><b>���Ѿ�ɾ��ѡ�����ļ���¼��"
	call dvbbs_suc()
end if
end sub

'ɾ�������ļ�
sub alldel()
if GroupSetting(48)=0 then
	errmsg=errmsg+"<br>"+"<li>��û����̳�ļ������Ȩ��,�������Ա��ϵ��"
	founderr=true
	exit sub
end if
	conn.execute("delete from DV_Upfile where F_Flag=1 and  F_UserID="&userid)
	sucmsg=sucmsg+"<br>"+"<li><b>���Ѿ�ɾ���������ļ���¼��"
	call dvbbs_suc()
end sub 

function filenum(types)
dim stype
if isnumeric(types)  then
	select case types
	case 0
	stype="F_Type=0 and"
	case 1
	stype="F_Type=1 and"
	case 2
	stype="F_Type=2 and"
	case 3
	stype="F_Type=3 and"
	case 4
	stype="F_Type=4 and"
	case else
	stype=""
	end select
else
	stype=""
end if
	rs=conn.execute("Select Count(F_ID) From DV_Upfile Where  "&stype&" F_UserID="&userid&"")
	filenum=rs(0)
	set rs=nothing
	if isnull(filenum) then filenum=0
end function

%>


