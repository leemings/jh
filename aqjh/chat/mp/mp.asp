<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("aqjh_usermdb")
conn.open connstr
sql="SELECT ����,���� FROM �û� where ����='" & aqjh_name &"'"
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	response.write "�㲻�ǽ�������"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       response.end
else
	pai=rs("����")
	tl=rs("����")
rs.close
sql="SELECT q,s FROM p where a='"&pai&"'"
rs.open sql,conn,1,1
if rs.bof or rs.eof then
	shengming="û������ս����"
	shidi="û�е���"
	else
	shengming=rs("s")
	shidi=rs("q")
	end if
rs.close
%>
<html><head>
<meta http-equiv="Content-Language" content="zh-tw">
<META http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>���ɴ�ս</title>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../BG.gif" bgproperties="fixed">
<div align="right">
  <table border="1" width="100%" cellspacing="0" cellpadding="2" bgcolor="#008000" bordercolordark="#FFFFFF" bordercolorlight="#000000" background="../BG.gif" align="right">
    <tr>
      <td width="100%" align="center"><img border="0" src="../img/ki17.gif" width="18" height="18">�����ɴ�ս��<img border="0" src="../img/ki17.gif"></td>
    </tr>
    <tr>
      <td width="100%"><a href="mpbao.asp">���뱣��</a>:<font color="#FFFFCC">�������ɱ���</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="mpbao2.asp">�رձ���</a>:<font color="#FFFFCC">�ر����ɱ���</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="javascript:s('/������ս$')">������ս</a>:<font color="#FFFFCC">����������ս</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="javascript:s('/�������$')">�������</a>:<font color="#FFFFCC">������������</font></td>
    </tr>
    <tr>
      <td width="100%"><a href="mp2.asp">����ս��</a>:<font color="#FFFFCC">��������ս��</font></td>
    </tr>
    <tr>
      <td width="100%">����ս��:<font color="#FFFFCC" size="2"><%=tl%></font></td>
    </tr>
    <tr>
      <td width="100%">����ս��:<font color="#FFFFCC" size="2"><%=shengming%></font></td>
    </tr>
    <tr>
      <td width="100%">����:<input type="text" size="12" value="<%=shidi%>" readonly></td>
    </tr>
    <tr>
      <td width="100%" align="center"> <table cellpadding="3" cellspacing="0" border="0" align="center" width="100%">
    <tr valign="middle"> 
      <td width="43%"> 
        <div align="center">��ʽ</div>
      </td>
      <td width="29%"> 
        <div align="center">�书</font></div>
      </td>
      <td width="28%"> 
        <div align="center">����</font></div>
      </td>
    </tr>
    <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
rs.open "SELECT * FROM y where b='" & pai & "' order by e",conn
do while not rs.eof and not rs.bof
dj=int(sqr((rs("e")/100)))+1
%>
    <tr valign="middle"> 
      <td width="43%"> 
        <div align="center"> <font color="#ff0000" size="2"> <a href="javascript:s('/���ɴ�ս$ <%=rs("a")%>')" title="������Ϊ��<%=dj%>��!"><%=rs("a")%></a></font></div>
      </td>
      <td width="29%"> 
        <div align="center"><font color="#FFFFCC" size="2"><%=rs("c")%></font></div>
      </td>
      <td width="28%"> 
        <div align="center"><font color="#FFFFCC" size="2"><%=rs("d")%></font></div>
      </td>
    </tr>
 <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
end if%>
</table>
</td>
    </tr>
    <tr><td width="100%" align="center">�����齭����վ��</td>
    </tr>
  </table></div></body></html>