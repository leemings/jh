<%
id=request("id")
my=session("Ba_jxqy_username")
%><!--#include file="dadata.asp"-->
<%
sql="select * from �û� where ����='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	response.write "�㲻�ǽ������ˣ�������ԭ��"
	conn.close
	response.end
else

sql="SELECT * FROM �˳� where ID=" & id 
	Set Rs=connt.Execute(sql)
        wu=rs("��Ʒ��")
	yin=rs("����")
	lx=rs("����")
       sm=rs("˵��")
       gn=rs("����")  
        mw=rs("��ζ��")%>

<%
		sql="select * from �û� where ����='" & my & "'"
		rs=conn.execute(sql)
		if yin<=rs("����") then
			sql="update �û� set ����=����- " & yin & "  where ����='" & my & "'"
			rs=conn.execute(sql)%>
			
<%
			sql="select * from ��Ʒ where ��Ʒ��='" & wu & "' and ӵ����='" & my & "'"
			set rs=connt.execute(sql)
			if rs.eof or rs.bof then
			sql="insert into ��Ʒ(��Ʒ��,ӵ����,����,˵��,����,��ζ��,����) values ('"& wu &"','"& my &"','"& lx &"','" & sm &"','"& gn &"',"& mw &","& yin &")"
			rs=connt.execute(sql)
                  
			connt.close
			Response.Redirect "caichang.asp"
			else
            
				
					response.write "�������ѹ����������Ʒ�����Բ�������"
					connt.close
					response.end
				
			end if
		else
			response.write "����������ԭ���������������"
			connt.close
			response.end
		end if
   rs.close
   set rs=nothing
end if
%>