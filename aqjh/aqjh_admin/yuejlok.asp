<%@ LANGUAGE=VBScript codepage ="936" %>
<!--#include file="../showchat.asp"-->
<!--#include file="../const2.asp"-->
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
dd=day(now())
jl_name=trim(request("jl_name"))
if dd<>jl_dd then 
	Response.Write "<script Language=Javascript>alert('��ʾ�����ڻ�����"&jl_dd&"�ţ���');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
%>
<html>
<head>
<title><%=Application("aqjh_chatroomname")%>�����½���</title>
<LINK href="css/css.css" type=text/css rel=stylesheet>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "Select * from �û� Where ����='"&jl_name&"'",conn
if rs.eof then
  	Response.Write "<script Language=Javascript>alert('��ʾ���Ҳ����û�����');location.href = 'javascript:history.go(-1)';</script>"
	response.end
else
'�����»��ֽ���
rs.close
rs.open "Select top "&jl_yjf_top&" * from �û� where ����<>'�ٸ�' order by mvalue desc",conn,1,3
top=1
do while not rs.eof
   if rs("����")=jl_name and top<=jl_yjf_top then
        top1=jl_yjf_top-top+1
        conn.Execute("update �û� set ��Ա��=��Ա��+"&top1*jl_yjf_jk&" where ����='"&jl_name&"'")
        call showchat("<img src=../chat/img/jk.gif><font color=ff0000>�������»��ֽ�����</font>"&jl_name&"Ϊ�����»������е�["&top&"]��,������Ա��["&top1*jl_yjf_jk&"]Ԫ��")
   end if
   top=top+1
   rs.movenext
loop
'������������
rs.close
sql="SELECT ������,count(id) as num FROM �û� where allvalue>"&jl_allvalue&" and ������='"&jl_name&"' group by ������ order by count(id) desc "
rs.open sql,conn,1,1
if not rs.eof then
   num=rs("num")
   conn.Execute("update �û� set ��Ա��=��Ա��+"&num*jl_lr_jk&",����=����+"&num*jl_lr_yl&" where ����='"&jl_name&"'")
   call showchat("<img src=../chat/img/menoy.gif><font color=ff0000>���������˽�����</font>"&jl_name&"Ϊ����һ������["&num&"]��(�ﵽ����)�������ؽ�����Ա��["&num*jl_lr_jk&"]Ԫ,����["&num*jl_lr_yl&"]����")
end if
  	Response.Write "<script Language=Javascript>alert('��ʾ�����½���["&jl_name&"]�û�������ϣ���');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
%>
</body></html>