<!--#include file="conn.asp"-->
<!-- #include file="inc/const.asp" -->
<!--#include file="inc/ubbcode.asp"-->
<!--#include file="inc/chkinput.asp"-->
<!-- #include file="inc/z_school_char.asp" -->
<%
'=========================================================
' File: z_school_sms.asp
' Version:5.0
' Date: 2003-1-13
' Script Written by wxzlxl http://www.zmcn.com QQ:628122
'=========================================================
' Copyright (C) 2001,2002 AspSky.Net. All rights reserved.
' Web: http://www.aspsky.net,http://www.dvbbs.net
' Email: info@aspsky.net,eway@aspsky.net
'=========================================================
%>
<html>
<head>
<title>ȫ�����Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="inc/Forum_css.asp"-->
</head>
<body <%=Forum_body(11)%>" onkeydown="if(event.keyCode==13 && event.ctrlKey)messager.submit()">
<%
dim msg
dim abgcolor
dim username
if not founduser then
  	errmsg=errmsg+"<br>"+"<li>��û��<a href=login.asp target=_blank>��¼</a>��"
	founderr=true
elseif BoardID="" or (not isInteger(BoardID)) or BoardID="0" then
Errmsg=Errmsg+"<br><li> �����ͬѧ¼��������ȷ�����Ǵ���Ч�����ӽ��롣"
	founderr=true
elseif cint(Board_Setting(2))=1 then
	if not founduser then
		Errmsg=Errmsg+"<br>"+"<li>��ͬѧ¼Ϊ��֤ͬѧ¼����<a href=login.asp>��¼</a>��ȷ�����Ǳ���ͬѧ��"
		founderr=true
	else
		if chkschoollogin(boardid,membername)=false then
		Errmsg=Errmsg+"<br>"+"<li>�㲻�Ǳ���ͬѧ������<a href=z_school_inclass.asp?boardid="&boardid&">���뱾��</a>��"
		founderr=true
		end if
	end if
end if
stats="ȫ�����Ϣ"
set rs=conn.execute("select boarduser from board where boardid="&boardid)
dim boarduser,sss
sss=""
if rs(0)<>"" then
boarduser=split(replace(rs(0)," ",""),",")
for i = 0 to ubound(boarduser)
sss=sss&"<OPTION value='"&boarduser(i)&"'>"&boarduser(i)&"</OPTION>"
next
else
Errmsg=Errmsg+"<br>"+"<li>��ͬѧ¼��û��ͬѧ��"
founderr=true
end if
boarduser=rs(0)
set rs=nothing
if founderr then
	call dvbbs_error()
else
	select case request("action")
	case "send"
		call savemsg()
	case else
call sendmsg()
	end select
end if
call activeonline()
call footer()

'������Ϣ
sub sendmsg()
stats="���Ͷ���"
dim sendtime,title,content
if request("id")<>"" and isNumeric(request("id")) then
set rs=server.createobject("adodb.recordset")
sql="select sendtime,title,content from message where incept='"&membername&"' and id="&request("id")
rs.open sql,conn,1,1
if not(rs.eof and rs.bof) then
sendtime=rs("sendtime")
title="RE " & rs("title")
content=rs("content")
end if
rs.close
set rs=nothing
end if
%>
<form action="z_school_sms.asp?BoardID=<%=BoardID%>" method=post name=messager>
<input type=hidden name="action" value="send">
<table cellpadding=3 cellspacing=1 align=center class=tableborder1>
          <tr> 
            <th colspan=3>���Ͷ���Ϣ��������������Ϣ��</td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=middle><b>�ռ��ˣ�</b></td>
            <td class=tablebody1 valign=middle>
              <input type=text name="touser" value="<%=request("touser")%>" size=50>
              <SELECT name=font onchange=DoTitle(this.options[this.selectedIndex].value)>
              <OPTION selected value="">ȫ��</OPTION><%=sss%>"></SELECT>
            </td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=top width=15%><b>���⣺</b></td>
            <td class=tablebody1 valign=middle>
              <input type=text name="title" size=70 maxlength=80 value="<%=title%>">
            </td>
          </tr>
          <tr> 
            <td class=tablebody1 valign=top width=15%><b>���ݣ�</b></td>
            <td  class=tablebody1 valign=middle>
              <textarea cols=70 rows=6 name="message" title="Ctrl+Enter����"><%if request("id")<>"" then%>
====== �� <%=sendtime%> ��������д���� ======
<%=server.htmlencode(content)%>
=====================================
<%end if%></textarea>
            </td>
          </tr>
          <tr> 
            <td  class=tablebody1 colspan=2>
<b>˵��</b>��<br>
�� ������ʹ��<b>Ctrl+Enter</b>����ݷ��Ͷ���<br>
�� ������Ӣ��״̬�µĶ��Ž��û�������ʵ��Ⱥ�������<b><%=GroupSetting(33)%></b>���û�<br>
�� �������<b>50</b>���ַ����������<b><%=GroupSetting(34)%></b>���ַ�<br>
            </td>
          </tr>
          <tr> 
            <td  class=tablebody2 valign=middle colspan=2 align=center> 
              <input type=Submit value="����" name=Submit>
              &nbsp; 
              <input type=Submit value="����" name=Submit>
              &nbsp; 
              <input type="reset" name="Clear" value="���">
              &nbsp; 
<%if request("reaction")="chatlog" then%>
              <input type=button value="�ر������¼" name="chatlog" onclick="location.href='?action=new&id=<%=request("id")%>&touser=<%=request("touser")%>'">
<%else%>
              <input type=button value="�鿴�����¼" name="chatlog" onclick="location.href='?action=new&id=<%=request("id")%>&touser=<%=request("touser")%>&reaction=chatlog'">
<%end if%>
              &nbsp; 
              <input type="button" name="close" value="�ر�" onclick="window.close()">
            </td>
          </tr>
<%if request("reaction")="chatlog" then%>
          <tr> 
            <th colspan=3>����<%=request("touser")%>�������¼</th>
          </tr>
<%if membername=request("touser") then%>
          <tr> 
            <td class=tablebody1 colspan=3>�Լ����Լ��������¼ûʲô�ÿ��ģ���</td>
          </tr>
<%else%>
<%
	set rs=server.createobject("adodb.recordset")
	sql="select * from message where ((sender='"&trim(membername)&"' and incept='"&replace(request("touser"),"'","")&"') or (sender='"&replace(request("touser"),"'","")&"' and incept='"&membername&"')) and delS=0 order by sendtime desc"
	rs.open sql,conn,1,1
	if rs.eof and rs.bof then
%>
          <tr> 
            <td class=tablebody1 colspan=3>��û���κ������¼��</td>
          </tr>
<%
	else
	do while not rs.eof
%>
                <tr>
                    <td class=tablebody2 height=25 colspan=3>
<%if rs("sender")=membername then%>
                    ��<b><%=rs("sendtime")%></b>�������ʹ���Ϣ��<b><%=htmlencode(rs("incept"))%></b>��
<%else%>
		    ��<b><%=rs("sendtime")%></b>��<b><%=htmlencode(rs("sender"))%></b>�������͵���Ϣ��
<%end if%></td>
                </tr>
                <tr>
                    <td  class=tablebody1 valign=top align=left colspan=3>
                    <b>��Ϣ���⣺<%=htmlencode(rs("title"))%></b><hr size=1>
                    <%=dvbcode(rs("content"),4,3)%>
		    </td>
                </tr>
<%
	rs.movenext
	loop
	end if
	rs.close
	set rs=nothing
%>
<%end if%>
<%end if%>
        </table>
</form>
<%
end sub


'1111111
sub savemsg()
stats="���Ͷ��ųɹ�"
dim incept,title,message,subtype
if request("touser")="" then
incept=boarduser
else
incept=request("touser")
end if
incept=CheckStr(incept)
incept=split(incept,",")
if request("title")="" then
	errmsg=errmsg+"<br>"+"<li>����û����д����ѽ��"
	founderr=true
	exit sub
elseif strlength(request("title"))>50 then
	errmsg=errmsg+"<br>"+"<li>�����޶����50���ַ���"
	founderr=true
	exit sub
else
	title=CheckStr(request("title"))
end if
if request("message")="" then
	errmsg=errmsg+"<br>"+"<li>�����Ǳ���Ҫ��д���ޡ�"
	founderr=true
	exit sub
elseif strlength(request("message"))>Cint(GroupSetting(34)) then
	errmsg=errmsg+"<br>"+"<li>�����޶����"&GroupSetting(34)&"���ַ���"
	founderr=true
	exit sub
else
	message=CheckStr(request("message"))
end if
for i=0 to ubound(incept)
sql="select username from [user] where username='"&replace(incept(i),"'","")&"'"
set rs=conn.execute(sql)
if rs.eof and rs.bof then
	errmsg=errmsg+"<br>"+"<li>��̳û������û���������ķ��Ͷ���д�����"
	founderr=true
	exit sub
end if
set rs=nothing
if request("Submit")="����" then
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,1)"
	subtype="�ѷ�����Ϣ"
elseif request("Submit")="����" then
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,0)"
	subtype="������"
else
	sql="insert into message (incept,sender,title,content,sendtime,flag,issend) values ('"&incept(i)&"','"&membername&"','"&title&"','"&message&"',Now(),0,1)"
	subtype="�ѷ�����Ϣ"
end if
conn.execute(sql)
next
sucmsg=sucmsg+"<br>"+"<li><b>��ϲ�������Ͷ���Ϣ�ɹ���</b><br>���͵���Ϣͬʱ����������"&subtype&"�С�"
call dvbbs_suc()
end sub
%>
<script language="javascript"> 
function DoTitle(addTitle) {  
 var revisedTitle;  
 var currentTitle = document.messager.touser.value; 

 if(currentTitle=="") revisedTitle = addTitle; 
 else { 
  var arr = currentTitle.split(","); 
  for (var i=0; i < arr.length; i++) { 
   if( addTitle.indexOf(arr[i]) >=0 && arr[i].length==addTitle.length ) return; 
  } 
  revisedTitle = currentTitle+","+addTitle; 
 } 

 document.messager.touser.value=revisedTitle;  
 document.messager.touser.focus(); 
 return; 
} 
</script>