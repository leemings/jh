<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
cmtype=Request.QueryString("cmtype")
if instr(cmtype,"'")<>0 or instr(cmtype," ")<>0 then Response.End
username=session("yx8_mhjh_username")
if username="" then Response.Redirect "chaterror.asp?id=000"
set conn=server.CreateObject("adodb.connection")
conn.Open Application("yx8_mhjh_connstr")
set rst=server.CreateObject("adodb.recordset")
sqlstr="select ID,����,����,����,��Ч from ��Ʒ where ����='"&cmtype&"' and ������='"&username&"' and װ��=True"
rst.Open sqlstr,conn
if rst.EOF or rst.BOF then
msg="������û��װ��������Ʒ�������޷�ж�أ�"
else
cmid=rst("id")
cmname=rst("����")
cmattack=rst("����")
cmdefence=rst("����")
cmespecial=rst("��Ч")
rst.Close
rst.Open "select * from �û� where ����='"&username&"' and ����>"&cmattack&" and ����>"&cmdefence,conn
if rst.EOF or rst.BOF then
msg="�����������㣬����ж��"&cmname
else
if cmespecial="��" then
especial=rst("�ؼ�")
else
especial=replace(rst("�ؼ�"),";"&cmespecial&";",";",1,1)
end if
conn.BeginTrans
conn.Execute "update �û� set ����=����-"&cmattack&",����=����-"&cmdefence&",�ؼ�='"&especial&"' where ����='"&username&"'"
conn.Execute "update ��Ʒ set ����=����+1,װ��=False where id="&cmid
if conn.Errors.Count>0 then
conn.RollbackTrans
Response.Redirect "../error.asp?id=104&errormsg="&conn.Errors(0).Description
else
conn.CommitTrans
msg="��ж����"&cmname&",������Щ������"&cmattack&",������˶�����"&cmdefence
end if
end if
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
%>
<head><title>�̵�</title><link rel="stylesheet" href="css.css"><script language=javascript>setTimeout("location.replace('seeequip.asp');",3000);</script></head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#FFDDF2" background='bg1.gif'>
<p align=center><font color=0000FF >ж��װ��</font></p>
<p align="center">
<font color=FF0000><%=msg%></font>
<br><br><br>3���Ӻ��Զ�����<br><a href="javascript:location.replace('seeequip.asp');" onmouseover="window.status='����';return true;" onmouseout="window.status='';return true;">����</a></body>
