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

if InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id,"��")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=120"

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("aqjh_usermdb")
	
sql="SELECT ��Ʒ��,����,����,����,״̬,˵��,������ FROM ��Ʒ�� where ����='��' and ID=" & id
rs.open sql,conn,1,1

if rs.eof or rs.bof then
       Response.Write "<script Language=Javascript>alert('û���������������ǲ���������ѽ!');window.close();</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       response.end
end if
       wu=rs("��Ʒ��")
       gj=rs("����")
	lx=rs("״̬")
       jin=int(rs("����"))
       dian=int(rs("����"))
       allgj=rs("������")
       sm=rs("˵��")
rs.close

sql="select * from �û� where id=" & myid
rs.open sql,conn,1,1
       hy=rs("��Ա�ȼ�")
       jinyan=int(rs("allvalue"))
       dianjuan=int(rs("���"))

   if lx="��" and hy<2 then
       Response.Write "<script Language=Javascript>alert('�㲻��2�����ϻ�Ա�����ܹ����������͵�����');window.close();</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       response.end
   end if

      if jin>jinyan then
                 Response.Write "<script Language=Javascript>alert('���򲻳ɹ�����û��ô����ѽ');window.close();</script>"
                 rs.close
	          set rs=nothing
	          conn.close
	          set conn=nothing
                 response.end
       end if
	

       if dian>dianjuan then
                 Response.Write "<script Language=Javascript>alert('���򲻳ɹ���������û��ô���ѽ');window.close();</script>"
                 rs.close
	          set rs=nothing
	          conn.close
	          set conn=nothing
                 response.end
       end if
   
         conn.execute"update �û� set allvalue=allvalue-" & jin & ",���=���-"& dian &"  where ����='" & my & "'"
         conn.Execute "insert into myanimal(animalname,username,attack,lei,allattack,sm)values('"&wu&"','"&my&"',"&gj&",'"&wu&"',"&allgj&",'"&sm&"')"
	   Response.Write "<script Language=Javascript>alert('����ɹ���');window.close();</script>"
                 rs.close
	          set rs=nothing
	          conn.close
	          set conn=nothing
                 response.end
                  
	%>


</body>
</html>
