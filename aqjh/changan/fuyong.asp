<%
name=session("myname")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
sql="SELECT * FROM �û� WHERE id="&aqjh_jhid
              Set Rs=conn.Execute(sql)
              if rs("����")>=rs("grade")*500000 or rs("����")>=rs("grade")*500000 then %>
<script language=vbscript>
MsgBox "��Ե�̫���ˣ�С�ĳ���"
location.href = "xiaowu3.asp"
</script>
<%else
              sql="select * from ��Ʒ where ����='ʳƷ' and ӵ����='" & name & "'"
		set rs=conn.execute(sql)			
                     id=rs("ID")
			if rs("����")<=0 then 
			sql="delete * from ��Ʒ where id=" & id
			set rs=conn.execute(sql)
                    Response.Redirect "xiaowu3.asp"
			else
			ti=rs("����")
			sql="update ��Ʒ set ����=����-1 where id=" & id
			set rs=conn.execute(sql)
			sql="update �û� set ����=����+" & ti & " where id="&aqjh_jhid
			set rs=conn.execute(sql)
                     conn.close
                     Response.Redirect "xiaowu3.asp"
                     response.end
    conn.close
    set rs=nothing
		end if
end if 
%>
<html><script language="JavaScript">                                                                  </script></html>