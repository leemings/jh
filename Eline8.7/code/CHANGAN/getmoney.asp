
<%money = clng(request.form("money"))
my=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
sql="select * from �û� where ����='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
%><script language=vbscript>
            MsgBox "�㲻�ǽ�������!"
          location.href = "../welcome.asp"
       </script><%
	conn.close
	response.end
else%>
<!--#include file="data1.asp"--><%
randomize()
         waittime=(rnd()*10000000)
         for i=0 to waittime
         next
sql="select * from ���� where ����='" & my & "' or  ����='" & my & "'"
set rs=conn.execute(sql)
yin=rs("����")
if money>yin then%>
<script language=vbscript>
                         MsgBox "�������Ľ����û����ô��������"
                         location.href = "xiaowu5.asp"
                    </script>
	<%elseif money>50000 or money<1 then%>
       <script language=vbscript>
                         MsgBox "������ÿ�����ֻ��ȡ5��"
                         location.href = "xiaowu5.asp"
                    </script><%
		else
       sql="update ���� set ����=����-'" & money & "' where ����='" & my & "' or  ����='" & my & "'"
       rs=conn.execute(sql)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
       sql="update �û� set ����=����+'" & money & "' where ����='" & my & "'"
       rs=conn.execute(sql)
       conn.close
       if yin<0 then
       sql="update �û� set ����=����+'" & yin & "' where ����='" & my & "'"
       rs=conn.execute(sql)
	conn.close
       end if
	Response.Redirect "xiaowu5.asp"
			conn.close
			response.end
		end if
   rs.close
   set rs=nothing
end if
%>
