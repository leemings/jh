<%@ LANGUAGE=VBScript%>
<!--#include file="check_user.asp"-->

<% 
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

if Session("aqjh_name")="" or Session("aqjh_id")="" then response.end
my=Session("aqjh_name")
grade=Session("aqjh_grade")
myid=session("aqjh_id")
id=request("id")

if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"��")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=120"

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
	sql="SELECT * FROM ��Ʒ�� where ����='����' and ID=" & id
	rs.open sql,conn,1,1
if rs.eof or rs.bof then
       Response.Write "<script Language=Javascript>alert('û����������Ʒ���ǲ���������ѽ!');window.close();</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       response.end
end if
       wu=rs("��Ʒ��")
	yin=rs("����")
       yinw=rs("������")
	lx=rs("����")
       gj=rs("����")
       fy=rs("����")
	zt=rs("״̬")
       nl=rs("����")
       tl=rs("����")
       sm=rs("˵��")
rs.close

sql="select ��� from �û� where id=" & myid
rs.open sql,conn,1,1
       jinyan=int(rs("���"))
	if yin>jinyan then
                 Response.Write "<script Language=Javascript>alert('���򲻳ɹ�����û��ô����ѽ');window.close();</script>"
                 rs.close
	          set rs=nothing
	          conn.close
	          set conn=nothing
                 response.end
       end if
             
                    conn.execute"update �û� set ���=���-" & yin & " where ����='" & my & "'"
			rs.close
                    sql="select ���� from ��Ʒ where ��Ʒ��='" & wu & "' and ӵ����='" & my & "'"
			rs.open sql,conn,1,1
                    
			if rs.eof or rs.bof then
			   conn.execute"insert into ��Ʒ(��Ʒ��,ӵ����,����,����,����) values ('"&wu&"','"&my&"','"&lx&"','"&tl&"','"&yin&"')"
                       conn.execute"update ��Ʒ set ����=1 where ��Ʒ��='" & wu & "' and ӵ����='" & my & "'"
                       Response.Write "<script Language=Javascript>alert('����" & wu & "�ɹ�!');window.close();</script>"

  rs.close
	                 set rs=nothing
	                 conn.close
	                 set conn=nothing

			   Response.Redirect "longeat.asp"
		        end if
                      shu=rs("����")
                    if shu>10 then 
                       Response.Write "<script Language=Javascript>alert('�����ϵ���Ʒ̫�࣬��������!');window.close();</script>"
                       rs.close
	                set rs=nothing
	                conn.close
	                set conn=nothing
                       response.end
                     else	
			   conn.execute"update ��Ʒ set ����=����+1  where ��Ʒ��='" & wu & "' and ӵ����='" & my & "'"
			    Response.Write "<script Language=Javascript>alert('����" & wu & "�ɹ�!');window.close();</script>"

		         rs.close
	                set rs=nothing
	                conn.close
	                set conn=nothing
			  Response.Redirect "longeat.asp"
                     end if
	%>


</body>
</html>
