<%
Rem ==========��̳ͨ�ú���=========
'ҳ�������ʾ��Ϣ
sub dvbbs_error()
%>
<br>
<table cellpadding=3 cellspacing=1 align=center width=95% bgcolor=<%=forum_body(27)%>>
<tr align=center>
<th width="100%" height=25 colspan=2>��̳������Ϣ</th>
</tr>
<tr>
<td width="100%" class=tablebody1 colspan=2>
<b>��������Ŀ���ԭ��</b><br><br>
<li>���Ƿ���ϸ�Ķ���<a href="boardhelp.asp?boardid=<%=boardid%>">�����ļ�</a>����������û�е�½���߲�����ʹ�õ�ǰ���ܵ�Ȩ�ޡ�
<font color="<%=Forum_body(8)%>"><%=errmsg%></font>
</td></tr>
<%if not founduser then
dim regjm
randomize timer
regjm=int(rnd*8998)+1000
%>
<script language="JavaScript">
ie = (document.all)? true:false
if (ie){function ctlent(eventobject){if(window.event.keyCode==13 && check()!=false){this.document.login.submit();}}}
function check()
{
var youname=document.login.name.value;
var mima=document.login.pass.value;
var regno1=document.login.regjm.value;
var regno2=document.login.regjm1.value;
if(youname=="" || youname==null){window.alert("��������Ҫ���棬�������������^_^!");return false;}
if( CheckIfEnglish(youname) )
{
window.alert("�����в����зǷ��ַ���Ӣ�ġ����֣�ֻ��ʹ������Ƭ������^_^!");
return false;
}
if(mima=="" || mima==null){window.alert("��������Ҫ���棬û��������ô����ѽ��^_^!");return false;}
if(regno1!=regno2){window.alert('��������ȷ����֤��:'+regno1+'��');return false;}
return true;
}
function CheckIfEnglish( String )
{
var Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890`=~!@#$%^&*()_+[]{}\\|/?.>,<;:'\-?<>/�����������������������������������������������������������������ѡ������ ������������ �� �� ������I�ơ� ���� ���ΦƦء� �ӡ��� ���ǅe�� �������� �ɡ衩 ���� �� �ԨK�L ������ੳ ���� ������n���� �� �٢ڢۢܢݢޢߢ� \"";
var i;
var c;
for( i = 0; i < String.length; i ++ )
{
c = String.charAt( i );
if (Letters.indexOf( c ) < 0)
return false;
}
return true;
}

</script>
<!--<form action="login.asp?action=chk" method=post> -->
<form action="../check0.asp" method="post" name="login" onsubmit="return(check())" target="_self" ;> 
    <input type=hidden name=regjm value="<%=regjm%>">
    <input type=hidden name=bbs value="1">
    <input type=hidden name=CookieDate value=0 checked>

    <tr>
    <th valign=middle colspan=2 align=center height=25><font color=#cc0000>��������׼ȷ�������Ľ������ϵ�½</font></td></tr>
    <tr>
    <td valign=middle class=tablebody1 align=center colspan=2><img alt="���ֽ���" border="0" height="32" src="images/eline.gif" width="162"><br>
    <script language=JavaScript src="../online.asp"></script><br>
    </td>
    
    <tr>
    <td valign=middle class=tablebody1>�����������û���</td>
    <td valign=middle class=tablebody1>
    <input type="text" onKeyDown="ctlent()"  name="name">
     &nbsp; <a href=reg.asp>û��ע�᣿</a></td></tr>
    <tr>
    <td valign=middle class=tablebody1>��������������</font></td>
    <td valign=middle class=tablebody1>
    <input type="password"  onKeyDown="ctlent()" name="pass"> &nbsp; <a href=lostpass.asp>�������룿</a></td></tr>
    <tr>
    <td valign=middle class=tablebody1>������������֤��</font></td>
    <td valign=middle class=tablebody1><input type=text onKeyDown="ctlent()" name="regjm1"> &nbsp; <%=regjm%></td></tr>
    <tr>
    <td valign=middle class=tablebody1>��¼����</font></td>
    <td valign=middle class=tablebody1>
    <select name="txwz">
                        <option value="0" selected>ͼ�ΰ�</option>
<!--                        <option value="1">���ְ�</option>-->
                      </select>                    
            </td></tr>
    <tr>
    <td valign=top width=30% class=tablebody1><b>�����½</b><BR> ������ѡ�������½����̳��Ա�����û��б�����������Ϣ��</td>
    <td valign=middle class=tablebody1>                <input type=radio name=userhidden value=2 checked>������½<br>
                <input type=radio name=userhidden value=1>�����½<br>
                </td></tr>
	<input type=hidden name=comeurl value="<%=Request.ServerVariables("HTTP_REFERER")%>">
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><input type=submit name=submit value="�� ½">&nbsp;&nbsp;<input type=button name="back" value="�� ��" onclick="location.href='<%=Request.ServerVariables("HTTP_REFERER")%>'"></td></tr>
</form>
<%else%>
    <tr>
    <td class=tablebody2 valign=middle colspan=2 align=center><a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><<������һҳ</a></td></tr>
<%end if%>
</table>
<%
end sub

sub dvbbs_suc()
%>
<br>
<table cellpadding=3 cellspacing=1 align=center width=95% bgcolor=<%=forum_body(27)%>>
<tr align=center>
<th width="100%">��̳�ɹ���Ϣ
</td>
</tr>
<tr>
<td width="100%" class=tablebody1>
<b>�����ɹ���</b><br><br>
<%=sucmsg%>
</td></tr>
<tr align=center><td width="100%" class=tablebody2>
<a href="<%=Request.ServerVariables("HTTP_REFERER")%>"> << ������һҳ</a>
</td></tr>
</table>
<%
end sub

rem �����ַ�
function ChkBadWords(fString)
    dim bwords,ii
    if not(isnull(BadWords) or isnull(fString)) then
    bwords = split(BadWords, "|")
    for ii = 0 to ubound(bwords)
        fString = Replace(fString, bwords(ii), string(len(bwords(ii)),"*")) 
    next
    ChkBadWords = fString
    end if
end function

Rem ����HTML����
function HTMLEncode(fString)
if not isnull(fString) then
    fString = replace(fString, ">", "&gt;")
    fString = replace(fString, "<", "&lt;")

    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(9), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")

    fString=ChkBadWords(fString)
    HTMLEncode = fString
end if
end function

Rem ���˱��ַ�
function HTMLcode(fString)
if not isnull(fString) then
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
    fString = Replace(fString, CHR(10), "<BR>")
    HTMLcode = fString
end if
end function

'�û�IP����
function LockIP(sip)
	dim str1,str2,str3,str4
	dim num
	LockIP=false
	if isnumeric(left(sip,2)) then
		str1=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str2=left(sip,instr(sip,".")-1)
		sip=mid(sip,instr(sip,".")+1)
		str3=left(sip,instr(sip,".")-1)
		str4=mid(sip,instr(sip,".")+1)
		if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then
	
		else
			num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
			sql="select count(*) from LockIP where ip1 <="&num&" and ip2 >="&num&""
			set rs=conn.execute(sql)
			if rs(0)>0 then 
				LockIP=true
			end if
			set rs=nothing
		end if
	end if
end function

Rem �жϷ����Ƿ������ⲿ
function ChkPost()
	dim server_v1,server_v2
	chkpost=false
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		chkpost=false
	else
		chkpost=true
	end if
end function

Rem �ж��û���Դ
function address(sip)
	dim str1,str2,str3,str4
	dim num
	dim country,city
	dim irs
	if isnumeric(left(sip,2)) then
	if sip="127.0.0.1" then sip="192.168.0.1"
	str1=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str2=left(sip,instr(sip,".")-1)
	sip=mid(sip,instr(sip,".")+1)
	str3=left(sip,instr(sip,".")-1)
	str4=mid(sip,instr(sip,".")+1)
	if isNumeric(str1)=0 or isNumeric(str2)=0 or isNumeric(str3)=0 or isNumeric(str4)=0 then

	else
		num=cint(str1)*256*256*256+cint(str2)*256*256+cint(str3)*256+cint(str4)-1
		sql="select Top 1 country,city from address where ip1 <="&num&" and ip2 >="&num&""
		set irs=server.createobject("adodb.recordset")
		irs.open sql,conn,1,1
		if irs.eof and irs.bof then 
		country="����"
		city=""
		else
		country=irs(0)
		city=irs(1)
		end if
		irs.close
		set irs=nothing
	end if
	address=country&city
	else
	address="δ֪"
	end if
end function

function iif(expression,returntrue,returnfalse)
	if expression=0 then
		iif=returnfalse
	else
		iif=returntrue
	end if
end function

function iiif(express,expression,returntrue,returnfalse)
	if iif(isnull(express),0,express)>iif(isnull(expression),0,expression) then
		iiif=returnfalse
	else
		iiif=returntrue
	end if
end function

function iimg(expression,returnfalse,returntrue)
	if expression="" or isnull(expression) then
		iimg=returnfalse
	else
		iimg=returntrue
	end if
end function

Rem ����SQL�Ƿ��ַ�
function checkStr(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	end if
	checkStr=replace(str,"'","''")
end function

Rem �û�����
sub activeonline()
dim ComeFrom,actCome,statuserid
statuserid=replace(replace(Request.ServerVariables("REMOTE_HOST"),".",""),"'","")
'response.Write(Request.ServerVariables("REMOTE_HOST"))
if not founduser then
	session("userid")=statuserid
	sql="select id,boardid from online where id="&cstr(session("userid"))
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		ComeFrom=""
		actCome=""
		sql="insert into online(id,username,userclass,ip,startime,lastimebk,boardid,browser,stats,actforip,UserGroupID,actCome,userhidden) values ("&statuserid&",'����','����','"&replace(Request.ServerVariables("REMOTE_HOST"),"'","")&"',Now(),Now(),"&boardid&",'"&replace(Request.ServerVariables("HTTP_USER_AGENT"),"'","")&";"&replace(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE"),"'","")&"','"&replace(stats,"'","")&"','"&replace(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),"'","")&"',7,'"&actCome&"',"&userhidden&")"
	else
		sql="update online set lastimebk=Now(),boardid="&boardid&",stats='"&replace(stats,"'","")&"' where id="&cstr(session("userid"))
	end if
	conn.execute(sql)
else
	if founderr then
		boardid=0
		stats="������Ϣ"
	end if
	sql="select id,boardid from online where userid="&userid
	set rs=conn.execute(sql)
	if rs.eof and rs.bof then
		ComeFrom=""
		actCome=""
		sql="insert into online(id,username,userclass,ip,startime,lastimebk,boardid,browser,stats,actforip,UserGroupID,actCome,userhidden,userid,usersex,uservip) values ("&statuserid&",'"&membername&"','"&memberclass&"','"&replace(Request.ServerVariables("REMOTE_HOST"),"'","")&"',Now(),Now(),"&boardid&",'"&replace(Request.ServerVariables("HTTP_USER_AGENT"),"'","")&";"&replace(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE"),"'","")&"','"&replace(stats,"'","")&"','"&replace(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),"'","")&"',"&UserGroupID&",'"&actCome&"',"&userhidden&","&userid&","&mysex&","&myvip&")"
	else
		sql="update online set lastimebk=Now(),boardid="&boardid&",stats='"&replace(stats,"'","")&"' where userid="&userid
	end If
'	response.Write(sql)
	conn.execute(sql)
	rs.close
	if session("userid")<>"" then
		Conn.Execute("delete from online where id="&session("userid"))
		session("userid")=""
	end if
end if
set rs=nothing
Rem ɾ����ʱ�û�
sql="Delete FROM online WHERE DATEDIFF('s', lastimebk, now()) > "&Forum_Setting(8)&"*60"
Conn.Execute sql
end sub

sub footer()
endtime=timer()%>
<p>
<TABLE cellSpacing=0 cellPadding=0 border=0 align=center id=footer>
<tr><td align=center>
<%=Forum_ads(1)%>
</td></tr>
<!--<tr>
<td align=center>
<a href="file:///::{450D8FBA-AD25-11D0-98A8-0800361B1103}" target="_blank"><img src=images/img_icon/mydocuments.gif border=0></a>&nbsp;
<a href="file:///::{20D04FE0-3AEA-1069-A2D8-08002B30309D}" target="_blank"><img src=images/img_icon/mycomputer.gif border=0></a>&nbsp;
<a href="file:///::{208D2C60-3AEA-1069-A2D7-08002B30309D}" target="_blank"><img src=images/img_icon/networkneighborhood.gif border=0></a>&nbsp;
<a href="file:///::{645FF040-5081-101B-9F08-00AA002F954E}" target="_blank"><img src=images/img_icon/recyclebin.gif border=0></a>&nbsp;
<a href="file:///::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}" target="_blank"><img src=images/img_icon/controlpanel.gif border=0></a>&nbsp;
<a href="file:///::{2227A280-3AEA-1069-A2DE-08002B30309D}" target="_blank"><img src=images/img_icon/printer.gif border=0></a>
</td>
</tr>-->
<tr><td align=center><%=Version%><br><%=Copyright%><%if Cint(forum_setting(30))=1 then%> , ҳ��ִ��ʱ�䣺<%=FormatNumber((endtime-startime)*1000,3)%>����<%end if%></td></tr>
<!--<tr><td align=center><script language=javascript src="z_counter.asp?style=3"></script></td></tr>-->
<!--<tr><td align=center><img src="z_counter.asp?style=0&BoardID=<%=request("BoardID")%>" border=0></td></tr>-->
</table>
<a name=bottom></a>
</body>
</html>
<%
	CloseDatabase
	set myCache=nothing
end sub

sub nav()
%>
<html>
<head>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta name=keywords content="���׵���,eline,51eline,һ��,����,���ֽ���,eline_email@etang.com,happyjh.com,bbs.happyjh.com,yxt_email@eatng.com,<%=forum_info(0)%>">
<title><%=Forum_info(0)%> - http://bbs.happyjh.com --<%=HTMLEncode(stats)%></title>
<!--#include file="Forum_css.asp"-->
<!--#include file="Forum_js.asp"-->
<%if forum_setting(71)=1 then
if founduser then%>
<SCRIPT language=javascript src="inc/menu.js"></SCRIPT>
<%end if
end if%>
</head>
<body <%=Forum_body(11)%> onmousemove="HideMenu()">
<a name=top></a>
<div id=menuDiv style='Z-INDEX: 2000; VISIBILITY: hidden; WIDTH: 1px; POSITION: absolute; HEIGHT: 1px; BACKGROUND-COLOR: #9cc5f8'></div>
<table cellspacing=0 cellpadding=0 align=center style="border:1px <%=Forum_body(27)%> solid; border-top-width: 0px; border-right-width: 1px; border-bottom-width: 0px; border-left-width: 1px;width:<%=Forum_body(12)%>;">
<tr><td width=100%>
<table width=100% align=center border=0 cellspacing=0 cellpadding=0>
<tr><td class=TopDarkNav height=9></td></tr>
<tr><td height=70 class=TopLighNav2>
<TABLE border=0 width="100%" align=center>
<TR>
<TD align=left width="25%"><a href="<%= Forum_info(3) %>"><img border=0 src='<%= Forum_info(6) %>'></a></TD>
<TD Align=center width="*"><%
	if not isnull(Forum_ads(0)) and Forum_ads(0)<>"" then
		%><%=Forum_ads(0)%><%
	end if
%></td>
<td align=right style="line-height: 15pt" width="50" nowrap><a href=#><span style="CURSOR: hand" onClick="window.external.AddFavorite('<%=Forum_info(1)%>', '<%=Forum_info(0)%>')">�����ղ�</span></a><br><a href="mailto:<%=Forum_info(5)%>">��ϵ����</a><br><a href="boardhelp.asp?boardid=<%=boardid%>">��̳����</a></td>
</tr>
</table>
</td></tr>
<tr><td class=TopLighNav height=9></td></tr>
        <tr> 
          <td class=TopLighNav1 height=22  valign="middle">&nbsp;&nbsp;
<%if not founduser then
	%><img src=pic/nologin.gif> ��ע�⣬����δ <a href="login.asp"><font color="<%=Forum_body(8)%>">��½</color></a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="reg.asp">ע��</a><%
else
	%><img src=pic/login.gif> ��ӭ�� <a href=dispuser.asp?name=<%=membername%> target=_blank><font color=blue><B><%=membername%></B></font></a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="#" onMouseOver='ShowMenu(mystatus,100)' id="mystat"><%
	if userhidden=1 then
		%>����<%
	elseif userhidden=2 then
		%>����<%
	elseif userhidden=3 then
		%>æµ<%
	elseif userhidden=4 then
		%>���ϻ���<%
	elseif userhidden=5 then
		%>�뿪<%
	elseif userhidden=6 then
		%>�����绰<%
	elseif userhidden=7 then
		%>����Ͳ�<%
	elseif userhidden=8 then
		%>����<%
	end if
	%></a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="login.asp">�ص�</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="usermanager.asp" onMouseOver='ShowMenu(manage,100)'>����</a> <%
end if
%> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="query.asp?boardid=<%=boardid%>">����</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="#" onMouseOver='ShowMenu(stylelist,100)'>���</a> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="boardstat.asp?boardid=<%=boardid%>" onMouseOver='ShowMenu(boardstat,100)'>״̬</a> <!--<img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="show.asp?boardid=<%=boardid%>" onMouseOver='ShowMenu(downlist,100)'>չ��</a> --><%
if founduser then
if forum_setting(71)<>1 then%> 
<!--��������������ʼ-->
<%set rs= server.createobject ("adodb.recordset")
sqlgroup = "select * from [group] order by groupid"
rs.open sqlgroup,connplus,0,1
if not (rs.bof and rs.eof) then
	do while not rs.eof
		response.write " <img src="&Forum_info(7)&"navspacer.gif align=absmiddle> <a onMouseOver='ShowMenu(plus"&rs("groupid")&",100)' href='#'>"&rs("groupname")&"</a>"
		rs.movenext
	loop
end if
rs.close
set connplus=nothing%>
<!--����������������-->
<%end if%>
<img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="logout.asp">�˳�</a><%else%> <img src=<%=Forum_info(7)%>navspacer.gif align=absmiddle> <a href="dispuser.asp?boardid=<%=boardid%>&action=permission">������ʲô</a><%end if%>
<%if master then response.write " <img src="&Forum_info(7)&"navspacer.gif align=absmiddle> <a href=admin_guanli.asp>����</a> <img src="&Forum_info(7)&"navspacer.gif align=absmiddle> <a href=""recycle.asp"">����վ</a>"%>
</td>
</tr>
</table>
</td></tr>
</table>
<%
if Cint(GroupSetting(0))=0 and (instr(lcase(trim(scriptname)),"reg.asp")=0 and instr(lcase(trim(scriptname)),"login.asp")=0) then
	Errmsg=Errmsg+"<br>"+"<li>��û���������̳��Ȩ�ޣ���<a href=login.asp>��½</a>����ͬ����Ա��ϵ��"
	call head_var(2,0,"","")
	call dvbbs_error()
	call footer()
	response.end
end if
end sub

'��ڲ���
'IsBoard=1��̳�ְ��浼����IsBoard=0��̳����ҳ�浼����GetTitle��̳����ҳ���ϼ�ҳ�棬GetUrl��̳����ҳ���ϼ�ҳ��URL
'Depth��̳�ְ��浼������̳��ȣ�����ҳ������Ϊ0
sub head_var(IsBoard,idepth,GetTitle,GetUrl)
%>
<table cellspacing=1 cellpadding=3 align=center border=0 width="<%=Forum_body(12)%>">
<tr>
<%if not founduser then
%><td height=25><BR>>> <%
	if foundboard then
		%><%=BoardReadme%><%
	else
		%>��ӭ���� <B><%=Forum_info(0)%></B><%
	end if
else
	%><td width=65% ></td>
	<td width=35% align=right><%
	if Cint(newincept())>Cint(0) then
		%><bgsound src="<%=Forum_info(7)&Forum_statePic(8)%>" border=0><%
		if Cint(forum_setting(10))=1 then
			%><script language=JavaScript>openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',500,400)</script><%
		end if
		%><img src=<%=Forum_info(7)&Forum_boardpic(9)%>> <a href="usersms.asp?action=inbox">�ҵ��ռ���</a> (<a href="javascript:openScript('messanger.asp?action=read&id=<%=inceptid(1)%>&sender=<%=inceptid(2)%>',500,400)"><font color="<%=Forum_body(8)%>"><%=newincept()%> ��</font></a>)<%
	else
		%><img src=<%=Forum_info(7)&Forum_boardpic(8)%>> <a href="usersms.asp?action=inbox">�ҵ��ռ���</a> (<font color=gray>0 ��</font>)<%
	end if
end if
%></td></tr>
</table>
<table cellspacing=1 cellpadding=3 align=center class=tableBorder2>
<tr>
<td height=25 valign=middle><img src="<%=Forum_info(7)&Forum_pic(12)%>" align=absmiddle> <a href=index.asp onMouseOver='ShowMenu(ALLboardlist,80)'><%=Forum_info(0)%></a> �� <%
if IsBoard=1 then
	if BoardParentID>0 then
	for i=0 to idepth-1
dim showmenu
if i=0 then showmenu="onMouseOver='ShowMenu(boardlist,80)'" else showmenu=""
response.write "<a "&showmenu&" href=list.asp?boardid="&FBoardID(i)&">"&FBoardName(i)&"</a> ��  "
		if i>9 then exit for
	next
	end if
	if request("CatLog")="NN" then
	Response.Cookies("BoardList")(BoardID & "BoardID")= "NNotShow"
	end if
	response.write "<a href=list.asp?boardid="&boardid&">"&boardtype&"</a> ��  " & HTMLEncode(stats)
	if request.cookies("BoardList")(boardid & "BoardID")="NNotShow" then
	response.write "&nbsp;<a href=""?BoardID="&boardid&"&cBoardid="&boardid&"&Catlog=Y"" title=""չ����̳�б�"">[չ��]</a>"
	end if
elseif IsBoard=2 then
	response.write HTMLEncode(stats)
else
	response.write "<a href="&GetUrl&">"&GetTitle&"</a> �� " & HTMLEncode(stats)
end if
%><a name=top></a></td>
</tr>
</table>
<br>
<%
end sub

'ͳ������
function newincept()
rs=conn.execute("Select Count(id) From Message Where flag=0 and issend=1 and delR=0 And incept='"& membername &"'")
    newincept=rs(0)
	set rs=nothing
	if isnull(newincept) then newincept=0
end function
function inceptid(stype)
	set rs=conn.execute("Select top 1 id,sender From Message Where flag=0 and issend=1 and delR=0 And incept='"& membername &"'")
	if stype=1 then
	inceptid=rs(0)
	else
	inceptid=rs(1)
	end if
	set rs=nothing
end function

sub getRe()
	if userid<>"" and isInteger(reID) then
		response.write "<bgsound src="&forum_info(7)&"reply.wav border=0><img src="& forum_info(7)&"Forum_readme.gif> <a href='dispbbs.asp?boardID="&reBoardID&"&ID="&reID&"'>�����������˻ظ���(<font color="&forum_body(8)&">"&reCount&"</font>)</a>&nbsp;[<a href=cookies.asp?action=resetre>ȫ�����</a>] "
	end if
end sub

Rem ��ð����û���Ȩ������
function GetBoardPermission()
	dim pmrs
	GetBoardPermission=false
	set pmrs=conn.execute("select PSetting from BoardPermission where Boardid="&Boardid&" and GroupID="&UserGroupID)
	if not (pmrs.eof and pmrs.bof) then
		GetBoardPermission=true
		GroupSetting=split(pmrs(0),",")
	else
		GetBoardPermission=false
	end if
	if FoundUser then
	set pmrs=conn.execute("select uc_Setting from UserAccess where uc_BoardID="&BoardID&" and uc_UserID="&userID)
	if not(pmrs.eof and pmrs.bof) then
		UserPermission=split(pmrs(0),",")
		GroupSetting=split(pmrs(0),",")
		FoundUserPer=true
	end if
	end if
	set pmrs=nothing
end function


Rem �ж������Ƿ�����
function isInteger(para)
       on error resume next
       dim str
       dim l,i
       if isNUll(para) then 
          isInteger=false
          exit function
       end if
       str=cstr(para)
       if trim(str)="" then
          isInteger=false
          exit function
       end if
       l=len(str)
       for i=1 to l
           if mid(str,i,1)>"9" or mid(str,i,1)<"0" then
              isInteger=false 
              exit function
           end if
       next
       isInteger=true
       if err.number<>0 then err.clear
end function

function allonline()
dim tmprs
tmprs=conn.execute("Select count(*) from online") 
allonline=tmprs(0) 
set tmprs=nothing 
if isnull(allonline) then allonline=0
end function

'**********************************************
'	vbs Cache��
'
'	����valid���Ƿ���ã�ȡֵǰ�ж�
'	����name��cache�����½������ֵ
'	����add(ֵ,����ʱ��)������cache����
'	����value������cache����
'	����blempty���Ƿ�δ����ֵ
'	����makeEmpty���ͷ��ڴ棬������
'	����equal(����1)���ж�cacheֵ�Ƿ�ͱ���1��ͬ
'	����expires(time)���޸Ĺ���ʱ��Ϊtime
'	ľ��  2002.12.24
'	http://www.aspsky.net/
'**********************************************

class Cache
	private obj			'cache����
	private expireTime		'����ʱ��
	private expireTimeName	'����ʱ��application��
	private cacheName		'cache����application��
	private path			'uri
	
	private sub class_initialize()
		path=request.servervariables("url")
		path=left(path,instrRev(path,"/"))
	end sub
	
	private sub class_terminate()
	end sub
	
	public property get blEmpty
		'�Ƿ�Ϊ��
		if isempty(obj) then
			blEmpty=true
		else
			blEmpty=false
		end if
	end property
	
	public property get valid
		'�Ƿ����(����)
		if isempty(obj) or not isDate(expireTime) then
			valid=false
		elseif CDate(expireTime)<now then
				valid=false
		else
			valid=true
		end if
	end property
	
	public property let name(str)
		'����cache��
		cacheName=str & path
		obj=application(cacheName)
		expireTimeName=str & "expires" & path
		expireTime=application(expireTimeName)
	end property
	
	public property let expires(tm)
		'�����ù���ʱ��
		expireTime=tm
		application.lock
		application(expireTimeName)=expireTime
		application.unlock
	end property
	
	public sub add(var,expire)
		'��ֵ
		if isempty(var) or not isDate(expire) then
			exit sub
		end if
		obj=var
		expireTime=expire
		application.lock
		application(cacheName)=obj
		application(expireTimeName)=expireTime
		application.unlock
	end sub
	
	public property get value
		'ȡֵ
		if isempty(obj) or not isDate(expireTime) then
			value=null
		elseif CDate(expireTime)<now then
			value=null
		else
			value=obj
		end if
	end property
	
	public sub makeEmpty()
		'�ͷ�application
		application.lock
		application(cacheName)=empty
		application(expireTimeName)=empty
		application.unlock
		obj=empty
		expireTime=empty
	end sub
	
	public function equal(var2)
		'�Ƚ�
		if typename(obj)<>typename(var2) then
			equal=false
		elseif typename(obj)="Object" then
			if obj is var2 then
				equal=true
			else
				equal=false
			end if
		elseif typename(obj)="Variant()" then
			if join(obj,"^")=join(var2,"^") then
				equal=true
			else
				equal=false
			end if
		else
			if obj=var2 then
				equal=true
			else
				equal=false
			end if
		end if
	end function
end class

sub DispPageNum(CurPage, PageCount, URLPrefix, URLPostfix)
	dim p,ii
	if PageCount=0 then
		response.write "�� "
	else
		p=(CurPage-1) \ 10
		if CurPage=1 then
			response.write "<SPAN class=pagenumstatic><font face=webdings color="&Forum_body(8)&">9</font></SPAN>  "
		else
			response.write "<SPAN class=pagenum><a href="&URLPrefix&"1"&URLPostfix&" title=��ҳ><font face=webdings>9</font></a></SPAN> "
		end if
		if p>0 then response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(p*10)&URLPostfix&" title=��ʮҳ><font face=webdings>7</font></a></SPAN> "
		response.write "<b>"
		for ii=p*10+1 to P*10+10
			if ii=CurPage then
		  	response.write "<SPAN class=pagenumstatic><B><font color="&Forum_body(8)&">"+Cstr(ii)+"</font></B></SPAN> "
			else
				response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(ii)&URLPostfix&">"+Cstr(ii)+"</a></SPAN> "
			end if
			if ii=PageCount then exit for
		next
		if ii>p*10+10 then ii=ii-1
		response.write "</b>"
		if ii<PageCount then response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(ii+1)&URLPostfix&" title=��ʮҳ><font face=webdings>8</font></a></SPAN> "
		if CurPage=PageCount then
			response.write "<SPAN class=pagenumstatic><font face=webdings color="&Forum_body(8)&">:</font></SPAN>"
		else
			response.write "<SPAN class=pagenum><a href="&URLPrefix&Cstr(PageCount)&URLPostfix&" title=βҳ><font face=webdings>:</font></a></SPAN> "
		end if
	end if
end sub

function renzhen(boardid,username)
	dim boarduser,rrs,Board_Setting,BoardMaster
	renzhen=false
	if master then
		renzhen=true
	else
		sql="select boarduser,Board_Setting,BoardMaster from board where boardid="&boardid
		set rrs=server.createobject("adodb.recordset")
		rrs.open sql,conn,1,1
		Board_Setting=split(rrs("board_setting"),",")
		if cint(Board_Setting(1))=1 then
			if Cint(GroupSetting(37))=0 then
				renzhen=false
			else
				renzhen=true
			end if
		elseif cint(Board_Setting(2))=1 then
			if not (isnull(rrs(2)) or rrs(2)="") then
				BoardMaster=split(rrs(2), "|")
				for i = 0 to ubound(BoardMaster)
					if trim(BoardMaster(i))=trim(username) then
						renzhen=true
						exit for
					end if
				next
			end if
			if renzhen=false then
				if isnull(rrs(0)) or rrs(0)="" then
					renzhen=false
				else
					boarduser=split(rrs(0), ",")
					for i = 0 to ubound(boarduser)
						if trim(boarduser(i))=trim(username) then
							renzhen=true
							exit for
						end if
					next
				end if
			end if
		else
			renzhen=true
		end if
		rrs.close
		set rrs=nothing		
	end if
end function

function VipBoard(boardid,username)
	dim boarduser,rrs,Board_Setting,BoardMaster
	VipBoard=false
	if master then
		VipBoard=true
	else
		sql="select Board_Setting,BoardMaster from board where boardid="&boardid
		set rrs=server.createobject("adodb.recordset")
		rrs.open sql,conn,1,1
		Board_Setting=split(rrs("board_setting"),",")
		if cint(Board_Setting(51))=1 or cint(Board_Setting(52))=1 then
			if not (isnull(rrs(1)) or rrs(1)="") then
				BoardMaster=split(rrs(1), "|")
				for i = 0 to ubound(BoardMaster)
					if trim(BoardMaster(i))=trim(username) then
						VipBoard=true
						exit for
					end if
				next
			end if
			if VipBoard=false then
				if myVip<>1 or isnull(MyVip) then
					VipBoard=false
				else
					VipBoard=true
				end if
			end if
		else
			VipBoard=true
		end if
		rrs.close
		set rrs=nothing		
	end if
end function

sub cache_link(bbslinkmode)
	myCache.name="bbslink"
	Dim linkinfo,NoLogoFlag
	i=7
	sql="select boardname,readme,url from bbslink where islogo=0 and not isnull(url) and url<>'' order by id"
	set rs=conn.execute(sql)
	NoLogoFlag = True
	if not rs.eof and not rs.bof then
		do while not rs.eof
			if i>6 then 
				linkinfo = linkinfo & "<tr><td width=""16%"">"
				i=1
			else
				linkinfo = linkinfo & "<td width=""16%"">"
			end if
			linkinfo = linkinfo & "<a href="""&trim(rs(2))&""" target=_blank title="""&trim(rs(1))&""">"&trim(split(rs(0),"|")(0))&"</a>"
			rs.movenext
			linkinfo = linkinfo & "</td>"
			if i=6 then linkinfo = linkinfo & "</tr>"
			i=i+1
		loop
		NoLogoFlag = False
	end if
	sql="select boardname,readme,url,logo from bbslink where islogo=1 and not isnull(url) and url<>'' order by id"
	set rs=conn.execute(sql)
	if not rs.eof and not rs.bof then
		if not NoLogoFlag then
			linkinfo = linkinfo & "<tr><td colspan=6><hr align=center size=1 color="&forum_body(27)&"></td></tr>"
		end if
		if cint(bbslinkmode)>1 then
			linkinfo = linkinfo & "<tr><td colspan=6><MARQUEE id=lmforum onmouseover=""this.stop()"" onmouseout=""if(stopflag==0) this.start();"" scrollAmount=2 scrollDelay=1 BEHAVIOR=SCROLL>"
			do while not rs.eof
				linkinfo = linkinfo & "<img onclick=window.open("""&trim(rs(2))&""",""_blank"") src="""&trim(rs(3))&""" border=0 alt="""&trim(split(rs(0),"|")(0))&"��"&trim(rs(1))&""" height=31 width=88"
				if cint(bbslinkmode)=3 then
					linkinfo = linkinfo & " style=""CURSOR:hand;FILTER: alpha(opacity=45)"" onMouseOut=nereidFade(this,45,10,10) onMouseOver=nereidFade(this,100,0,10)"
				else
					linkinfo = linkinfo & " style=""CURSOR:hand"""
				end if
				linkinfo = linkinfo & ">"
				rs.movenext
				linkinfo = linkinfo & "&nbsp;&nbsp;&nbsp;&nbsp;"
			loop
			linkinfo = linkinfo & "</MARQUEE></td></tr>"
		else
			linkinfo = linkinfo & "<tr><td colspan=6><hr align=center size=1 color="&forum_body(27)&"></td></tr>"
			do while not rs.eof
				if i>6 then 
					linkinfo = linkinfo & "<tr><td width=""16%"">"
					i=1
				else
					linkinfo = linkinfo & "<td width=""16%"">"
				end if
				linkinfo = linkinfo & "<img onclick=window.open("""&trim(rs(2))&""",""_blank"") src="""&trim(rs(3))&""" border=0 alt="""&trim(split(rs(0),"|")(0))&"��"&trim(rs(1))&""" height=31 width=88"
				if cint(bbslinkmode)=1 then
					linkinfo = linkinfo & " style=""CURSOR:hand;FILTER: alpha(opacity=45)"" onMouseOut=nereidFade(this,45,10,10) onMouseOver=nereidFade(this,100,0,10)"
				else
					linkinfo = linkinfo & " style=""CURSOR:hand"""
				end if
				linkinfo = linkinfo & ">"
				rs.movenext
				linkinfo = linkinfo & "</td>"
				if i=6 then linkinfo = linkinfo & "</tr>"
				i=i+1
			loop
			if isnull(linkinfo) or linkinfo="" then linkinfo=" "
		end if
	else
		if cint(bbslinkmode)>1 then
			if not NoLogoFlag then
				linkinfo = linkinfo & "<tr><td colspan=6><hr align=center size=1 color="&forum_body(27)&"></td></tr>"
			end if
			linkinfo = linkinfo & "<tr><td colspan=6><MARQUEE id=lmforum onmouseover=""this.stop()"" onmouseout=""if(stopflag==0) this.start();"" scrollAmount=2 scrollDelay=1 BEHAVIOR=SCROLL>"
			linkinfo = linkinfo & "&nbsp;&nbsp;&nbsp;&nbsp;"
			linkinfo = linkinfo & "</MARQUEE></td></tr>"
		end if
	end if
	set rs=nothing
	if isnull(linkinfo) or linkinfo="" then linkinfo=" "
	myCache.add linkinfo,dateadd("n",9999,now)
end sub

function codecookie(str)
	dim i
	dim strrtn
	for i=len(str) to 1 step -1
		strrtn=strrtn & ascw(mid(str,i,1))
		if (i<>1) then strrtn=strrtn & "a"
	next
	codecookie=strrtn
end function

function decodecookie(str)
	dim i
	dim strarr,strrtn
	strarr=split(str,"a")
	for i=ubound(strarr) to lbound(strarr) step -1
	strrtn=strrtn & chrw(strarr(i))
	next
	decodecookie=strrtn
end function

Function FilterJS(v)
	if not isnull(v) then
		dim t
		dim re
		dim reContent
		Set re=new RegExp
		re.IgnoreCase =true
		re.Global=True
		re.Pattern="(&#)"
		t=re.Replace(v,"<I>&#</I>")
		re.Pattern="(javascript)"
		t=re.Replace(t,"<I>&#106avascript</I>")
		re.Pattern="(jscript:)"
		t=re.Replace(t,"<I>&#106script:</I>")
		re.Pattern="(js:)"
		t=re.Replace(t,"<I>&#106s:</I>")
		re.Pattern="(value)"
		t=re.Replace(t,"<I>&#118alue</I>")
		re.Pattern="(about:)"
		t=re.Replace(t,"<I>about&#58</I>")
		re.Pattern="(file:)"
		t=re.Replace(t,"<I>file&#58</I>")
		re.Pattern="(document.cookie)"
		t=re.Replace(t,"<I>documents&#46cookie</I>")
		re.Pattern="(vbscript:)"
		t=re.Replace(t,"<I>&#118bscript:</I>")
		re.Pattern="(vbs:)"
		t=re.Replace(t,"<I>&#118bs:</I>")
		re.Pattern="(on(mouse|exit|error|click|key))"
		t=re.Replace(t,"<I>&#111n$2</I>")
		FilterJS=t
		set re=nothing
	end if
End Function

Function TopicIcon(BoardID,TopicID,LastPostTime,TopFlag,VoteFlag,LockFlag,HotFlag)
	dim NewFlag,LastViewBoard,LastViewTopic
	dim ImgSrc,ImgAlt

	if request.cookies("LastView")("Board_"&BoardID)="" then
		if isnull(LastViewTime) then
			LastViewBoard=CDate("2003-1-1")
		else
			LastViewBoard=LastViewTime
		end if
	else
		if isnull(LastViewTime) then
			LastViewBoard=CDate(request.cookies("LastView")("Board_"&BoardID))
		elseif LastViewTime<CDate(request.cookies("LastView")("Board_"&BoardID)) then
			LastViewBoard=CDate(request.cookies("LastView")("Board_"&BoardID))
		else
			LastViewBoard=LastViewTime
		end if
	end if
	if request.cookies("LastView")("Topic_"&TopicID&"_"&BoardID)="" then
		LastViewTopic=LastViewBoard
	else
		if CDate(request.cookies("LastView")("Topic_"&TopicID&"_"&BoardID))>LastViewBoard then
			LastViewTopic=CDate(request.cookies("LastView")("Topic_"&TopicID&"_"&BoardID))
		else
			LastViewTopic=LastViewBoard
		end if
	end if
	if LastViewTopic>=LastPostTime then
		NewFlag=0
	else
		NewFlag=1
	end if
	ImgSrc=Forum_info(7)&"TopicIcon/"&NewFlag&TopFlag&VoteFlag&LockFlag&HotFlag&".GIF"
	ImgAlt=""
	if LockFlag=1 then
		ImgAlt=ImgAlt&"������"
	end if
	if TopFlag=2 then
		ImgAlt=ImgAlt&"�̶ܹ�"
	elseif TopFlag=1 then
		ImgAlt=ImgAlt&"�̶�"
	end if
	if ImgAlt<>"" and instr(1,ImgAlt,"��")<=0 then
		ImgAlt=ImgAlt&"��"
	end if
	if HotFlag=1 then
		ImgAlt=ImgAlt&"����"
	end if
	if ImgAlt<>"" and instr(1,ImgAlt,"��")<=0 then
		ImgAlt=ImgAlt&"��"
	end if
	if NewFlag=1 then
		ImgAlt=ImgAlt&"��"
	else
		ImgAlt=ImgAlt&"��"
	end if
	if VoteFlag=1 then
		ImgAlt=ImgAlt&"ͶƱ"
	else
		ImgAlt=ImgAlt&"����"
	end if
	TopicIcon="<img src="""&ImgSrc&""" alt="""&ImgAlt&""" width=13 height=16 boarder=0>"
End Function

Function isCanScore(ScoreUser,PostUser)
	dim ScoreUserGroupID
	if not founduser then
		isCanScore=1
	elseif not (Master or superboardmaster or boardmaster) then
		isCanScore=2
	elseif lcase(trim(membername))=lcase(trim(PostUser)) then
		isCanScore=3
	elseif isnull(ScoreUser) or not (ScoreUser<>"") then
		isCanScore=0
	elseif lcase(trim(ScoreUser))=lcase(trim(membername)) then
		isCanScore=0
	else
		isCanScore=0
		ScoreUserGroupID=conn.execute("select usergroupid from [user] where username='"&ScoreUser&"'")(0)
		if UserGroupID<3 and UserGroupID=ScoreUserGroupID then
			isCanScore=4
		elseif UserGroupID=3 and UserGroupID=ScoreUserGroupID then
			if boardmaster then
				dim rss,userboardmaster
				sql="select BoardMaster from board where boardid="&boardid
				set rss=server.createobject("adodb.recordset")
				rss.open sql,conn,1,1
				userboardmaster=rss(0)
				rss.close
				set rss=nothing
				if instr(1,"|"&lcase(trim(userboardmaster))&"|","|"&lcase(trim(ScoreUser))&"|")>0 then
					isCanScore=4
				end if
			end if
		elseif UserGroupID<4 and UserGroupID>ScoreUserGroupID then
			isCanScore=5
		end if
	end if
End Function
%>