<%@ LANGUAGE=VBScript%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../../error.asp?id=440"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "select id,����,���,����,����,�书,����,����,�ȼ�,����,����,���� from �û� where ����='"&aqjh_name&"'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=440"
	response.end
end if
aqjh_id=rs("id")
if rs("�ȼ�")<35 or rs("����")<100000 or rs("����")<500000 or rs("�书")<50000 or rs("����")<40600 or rs("����")<36700 or rs("����")<40000 or rs("����")<50000 or rs("����")<100000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=460"
	response.end
end if
if rs("���")="����" then
	Response.Write "<script Language=Javascript>alert('��ʾ���������Ѿ��������ˣ�����Ҫ��������ѽ��');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
if rs("����")="�ٸ�" or (rs("����")<>"��" and rs("����")<>"����") then
	Response.Write "<script Language=Javascript>alert('��ʾ��Ҫ���Դ����ɾͱ��������뿪���ڵ����ɣ�');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
tmprs=conn.execute("Select count(*) As ���� from �û� where �ȼ�>=18 and ������='"& aqjh_name &"'")
lr=tmprs("����")
set tmprs=nothing
if lr<5 then
	Response.Write "<script Language=Javascript>alert('��ʾ��������˼�¼����5�ˣ������������˵ĵȼ���û�е�18�����ϣ�');window.close();</script>"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
rs.close
function chuser(u)
  dim filter,xx,usernameenable,su
  for i=1 to len(u)
    su=mid(u,i,1)
    xx=asc(su)
    zhengchu = -1*xx \ 256
    yushu = -1*xx mod 256
    if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
      chuser=true
      exit function
    end if
  next
  chuser=false
end function
if trim(request.form("subject"))="" or trim(request.form("comment"))=""  then%>
	<script language=vbscript>
	  MsgBox "���󣺰������ƺͼ�������д��"
	  location.href = "javascript:history.back()"
	</script>
	<%
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
mypai1=trim(request.form("subject"))
jjjj=request.form("comment")
namelen=0
for i=1 to len(mypai1)
	zh=mid(mypai1,i,1)
	zhasc=asc(zh)
	if zhasc<0 then
		namelen=namelen+2
	else
		namelen=namelen+1
	end if
next
if namelen>10 or namelen<4 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
if zhasc>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
namelen=0
if len(mypai1)>10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.redirect"../../error.asp?id=027"
end if
if len(jjjj)>50 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=028"
end if
if mypai1="��" or mypai1="������" or mypai1=" ������" or mypai1="����" or mypai1="�ٸ�" or mypail=" �ٸ�" or mypai1="����" or mypai1="����" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
if InStr(mypai1,"��")<>0 or InStr(mypai1,"��")<>0 or InStr(mypai1,"��")<>0 or InStr(mypai1,"��")<>0 or InStr(mypai1,"��")<>0 or InStr(mypai1,"��")<>0  or  InStr(mypai1,"����")<>0  or Instr(mypai1,"ү")<>0  or Instr(mypai1,"��")<>0  or Instr(mypai1,"��")<>0 or Instr(mypai1,"��")<>0 or Instr(mypai1,"��")<>0 or Instr(mypai1,"��")<>0 or Instr(mypai1,"��")<>0 or Instr(mypai1,"��")<>0 or Instr(mypai1,"��")<>0 or Instr(mypai1,"��")<>0 or Instr(mypai1,"��")<>0 or Instr(mypai1,"ɫ")<>0 or Instr(mypai1,"����")<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
if left(mypai1,3)="%20" or InStr(mypai1,"=")<>0 or InStr(mypai1,"`")<>0 or InStr(mypai1,"'")<>0 or InStr(mypai1," ")<>0 or InStr(mypai1,"��")<>0 or InStr(mypai1,"'")<>0 or InStr(mypai1,chr(34))<>0 or InStr(mypai1,"\")<>0 or InStr(mypai1,",")<>0 or InStr(mypai1,"<")<>0 or InStr(mypai1,">")<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
if chuser(mypai1) then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
if left(jjjj,3)="%20" or instr(jjjj,"or")<>0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../../error.asp?id=027"
end if
sql="Select a from p where a='"& mypail &"'"
set rs=conn.execute(sql)
if not(rs.bof or rs.eof) then %>
	<script language=vbscript>
	  MsgBox "���󣺰��������Ѿ����ڣ�"
	  location.href = "javascript:history.back()"
	</script>
	<%
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
'e��������
conn.execute "Insert Into p (a,b,c,d,e,g,h,i) Values('"&mypai1&"','"&aqjh_name&"',0,'" & jjjj & "','����������','��',50000000,'"&now()&"')"
conn.execute "update �û� set ����='"&mypai1&"',���='����',����=����-100000000,grade=5 where ����='"&aqjh_name&"'"
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body {font-size:9pt;SCROLLBAR-FACE-COLOR: #c2cea8;SCROLLBAR-HIGHLIGHT-COLOR: #ffffff;SCROLLBAR-SHADOW-COLOR: #000000;SCROLLBAR-3DLIGHT-COLOR: #ffffff;SCROLLBAR-ARROW-COLOR: #ffffff;SCROLLBAR-TRACK-COLOR: #ffffff;SCROLLBAR-DARKSHADOW-COLOR: #ffffff;SCROLLBAR-BASE-COLOR: #c2cea8}
table {font-size: 9pt}
a:link {CURSOR:crosshair;color: yellow; text-decoration: none}
a:visited {CURSOR:crosshair;color: yellow; text-decoration: none}
a:active {CURSOR:crosshair;color: #white; text-decoration: none}
a:hover { CURSOR:crosshair;color: #white; text-decoration: underline overline}
-->
</style>
<script language="javascript">
function opench(url,win){
        controlWindow=window.open(url,win,"fullscreen=no,toolbar=yes,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=20,left=2,width=450,height=300");
        controlWindow.focus();
        }
function ch(){
	document.forms[0].userpassword.value=document.forms[0].pwd.value;
	document.forms[0].pwd.value="";
//	document.forms[0].submit();
	}
</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=#000000 text=#ffffff link="#000080" alink="#800000" vlink="#000080">
<div align="center"> 
  <p><font size="3"><b>�����ѳɹ���ӣ�</b></font></p>
</div>
</body>
</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>