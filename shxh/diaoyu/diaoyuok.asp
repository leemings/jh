<%
username=session("Ba_jxqy_username")
if username="" then Response.Redirect "../error.asp?id=440"
Set Cn=Server.CreateObject("ADODB.Connection")
set rst=server.createobject("adodb.recordset")
diaoyu=Application("Ba_jxqy_connstr")
Cn.Open diaoyu
randomize timer
r=int(rnd*11)
rst.open"select * from ���� where ����='"& username &"'",cn,1,1
if rst.bof or rst.eof then 
Response.Redirect "diaoyu.asp"
cn.close                                                                 
end if
if rst("ʱ��")>now()-1/720 then 
Response.Redirect "diaoyu.asp"
cn.close                                                                 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("Ba_jxqy_connstr")
conn.open connstr
if r=2 or r=1 then
mess="��ϲ�����һ�������㣬����Ϊ����ʹ�ã�ɱ������1500������1300"
sql="select * from ��Ʒ where ����='������' and ������='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof or rs.bof then
			sql="insert into ��Ʒ(����,������,����,����,����,�۸�) values ('������','"& username &"','����',-1300,-1500,2500)"
			rs=conn.execute(sql)
                        else 
				sql="update ��Ʒ set �۸�=2500,����=����+1  where ����='������' and ������='" & username & "'"
				rs=conn.execute(sql)
		        end if

else
if r=3 then
mess="��ϲ�����һ�������㣬����������2������"
conn.execute "update �û� set ����=����+20000 where ����='"& username &"'"
else
if r=4 or r=5 then
mess="��ϲ�����һֻ����㣬���¿�����������1500������1500"
sql="select * from ��Ʒ where ����='�����' and ������='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof or rs.bof then
			sql="insert into ��Ʒ(����,������,����,����,����,�۸�) values ('�����','"& username &"','ҩƷ',1500,1500,3000)"
			rs=conn.execute(sql)
                        else
				sql="update ��Ʒ set �۸�=3000,����=����+1  where ����='�����' and ������='" & username & "'"
				rs=conn.execute(sql)
		        end if
else
if r=6 or r=8 then
mess="��ϲ�����һ��С���㣬���¿�����������500������500"
sql="select * from ��Ʒ where ����='С����' and ������='" & username & "'"
			set rs=conn.execute(sql)
			if rs.eof or rs.bof then
			sql="insert into ��Ʒ(����,������,����,����,����,�۸�) values ('С����','"& username &"','ҩƷ',500,500,800)"
			rs=conn.execute(sql)
                        else
				sql="update ��Ʒ set �۸�=800,����=����+1  where ����='С����' and ������='" & username & "'"
				rs=conn.execute(sql)
		        end if
else
if r=9 or r=10 then
mess="��������û������ˤ��������������900����������700"
conn.execute "update �û� set ����=����-900,����=����-700 where ����='" & username & "'"
else
mess="�ҵ��찡����ô����������ֻ��Ь�����㹳���۶��ˣ�����10000���������㹳����......"
conn.execute "update �û� set ����=����-10000 where ����='" & username & "'"
end if
end if
end if
end if
end if
cn.execute "Delete * From ���� Where ����='" & username & "'"
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<link href="../chat/style.css" rel=stylesheet type="text/css">
<title></title>
</head>

<body oncontextmenu=self.event.returnValue=false bgcolor="#ffffff" text="#000000">
<div align="center">
  <p>��</p>
  <table border=1 align=center cellpadding="10" cellspacing="13" height="100">
    <tr>
      <td> 
        <table>
          <tr> 
            <td valign="top"> 
              <div align="center"> 
                <p><%=mess%><br><a href="diaoyu.asp">���ؼ���</a></p>
              </div>
        </table>
      </td>
    </tr>
  </table>
</div>
</body>
</html>
<%                                                                 
set rs=nothing                                                                                                                                 
set conn=nothing  
%>