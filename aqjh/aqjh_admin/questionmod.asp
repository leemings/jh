<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.Buffer=true
aqjh_name=Session("aqjh_name")
aqjh_grade=Session("aqjh_grade")
aqjh_jhdj=Session("aqjh_jhdj")
if aqjh_name="" then Response.Redirect "../error.asp?id=440"
if session("aqjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if aqjh_grade<>10 or instr(Application("aqjh_admin"),aqjh_name)=0  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
ID=request("ID")
if ID<>"" then
  sql="select * from questlib where ID="&ID
  set rs=conn.Execute(sql)
  if rs.eof or rs.bof then ID=""
  
end if
if request("cmdYes")="ȷ��" then
  stype=left(request("stype"),10)
  squestion=left(request("squestion"),100)
  sanswer=left(request("sanswer"),50)
  addexp=clng(request("addexp"))
  if sanswer="" or squestion="" or stype="" then
    errtext="<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666><font color=green>����</font>�����ࡢ���ʡ��𰸲���Ϊ�գ�"
  elseif id<>"" then
    sql="update questlib set ����='"&stype&"',����='"&squestion&"',�ش�='"&sanswer&"',����="&addexp&" where id="&id
    conn.Execute(sql)
    errText="<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666><font color=green>�����ɹ�</font>����ɹ����޸��˴������ݣ�"
  else
    sql="insert into questlib (����,����,�ش�,����) values ('"&stype&"','"&squestion&"','"&sanswer&"',"&addexp&")"
    conn.Execute(sql)
    errText="<LINK href=css/css.css type=text/css rel=stylesheet><body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666><font color=green>�����ɹ�</font>����ɹ���������µĴ��⣡"

  end if
  response.write errText&"<BR>"&"<center>[<a href='javascript:history.go(-2)'>����</a>]"
  
else%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<LINK href=css/css.css type=text/css rel=stylesheet>
</head>
<body bgColor=#dcdec6 text=#000000 link=#003366 vLink=#006666>
  <table width=400 align=center cellspacing=0 cellpadding=0 border=1 BORDERCOLORLIGHT=#000000 BORDERCOLORDARK=#ffffff>
  <tr><td class=navy1>�޸Ļ���Ӵ����</td></tr>
 <form action="questionmod.asp" method="post">
 <input type=hidden name="id" value="<%=id%>">
 <tr><td height=100 style="text-align:left">����:<input type="text" name="stype" maxlength="10" size="10" value="<%if ID<>"" then%><%=rs("����")%><%else%>����<%END IF%>"><font color=990000>[д���������ͣ��Դ���]</font>
 <br>����:<input type="text" name="squestion" maxlength="100" size="40" value="<%if ID<>"" then%><%=rs("����")%><%END IF%>">
 <br>��:<input type="text" name="sanswer" maxlength="50" size="40" value="<%if ID<>"" then%><%=rs("�ش�")%><%END IF%>">
 <br>����:<input type="text" name="addexp" maxlength="10" size="10" value="<%if ID<>"" then%><%=rs("����")%><%else%>9999<%END IF%>"><font color=990000>[����Ϊ������������Ŀ����]</font>
  <tr><td height=40><input class=navy1 type="submit" name="cmdYes" value="ȷ��" >
    <input class=navy1 type="reset" name="cmdClear" value="���" onclick="javascript:document.froms[0].stype.focus();return true">
    </td></tr></form></table><tr<td><br><center>[<a href='javascript:history.go(-1)'>����</a>]</body>
<%
end if
if id<>"" then rs.close
set rs=nothing
conn.close
set conn=nothing%>