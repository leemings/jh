<!--#include file="conn.asp"-->
<!--#include file="const.asp"-->
<% 
if not master then
	errmess="��û��ִ�д˲�����Ȩ�ޣ�"
	call endinfo(1)
else
	dim mess
	if request("action")="UpPrice" then
		sql= "select ��ǰ�۸�,���̼۸�,sid from ��Ʊ order by sid"
		set rs=conn.execute(sql) 
		DO while not rs.EOF 
			if (rs(0)-rs(1))/rs(1)<0.15 then 
				conn.execute("update ��Ʊ set ��ǰ�۸�=��ǰ�۸�*1.04 where sid="&rs(2))
			end if
			rs.movenext
		loop
		mess="<font color=#3399FF>�������и�Ԥ�����Ʊ�г�Ͷ��޶��ʽ����й�Ʊ�۸�<font color=#FF6699>����</font> 4 %</font>"
		conn.execute("insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )")
	elseif request("action")="DownPrice" then
		sql= "select ��ǰ�۸�,���̼۸�,sid from ��Ʊ order by sid"
		set rs=conn.execute(sql)
		do while not rs.eof 
			if (rs(0)-rs(1))/rs(1)>-0.15 then 
				sql="update ��Ʊ set ��ǰ�۸�=��ǰ�۸�*0.97 where sid="&rs(2)
				conn.execute sql 
			end if
			rs.movenext
		loop
		mess="<font color=#3399FF>���÷�չ���ȣ�ͨ���������أ���Ʊ�г����й�Ʊ�۸�<font color=#FF6699>�´�</font> 3 %</font>"
		conn.execute("insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )")
	end if
        CloseDatabase		'�ر����ݿ�
	Response.Redirect "stock.asp"
end if
'CloseDatabase		'�ر����ݿ�	
%>