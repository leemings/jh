<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%if not master then
	errmess="��û��ִ�д˲�����Ȩ�ޣ�"
	call endinfo(1)
else
	dim mess
	if request("action")="UpPrice" then
		sql= "select DangQianJiaGe,KaiPanJiaGe,sid from [GuPiao] order by sid"
		set rs=gp_conn.execute(sql) 
		DO while not rs.EOF 
			if (rs(0)-rs(1))/rs(1)<0.15 then 
				gp_conn.execute("update [GuPiao] set DangQianJiaGe=DangQianJiaGe*1.04 where sid="&rs(2))
			end if
			rs.movenext
		loop
		mess="<font color=#3399FF>�������и�Ԥ�����Ʊ�г�Ͷ��޶��ʽ����й�Ʊ�۸�<font color=#FF6699>����</font> 4 %</font>"
		gp_conn.execute("insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )")
	elseif request("action")="DownPrice" then
		sql= "select DangQianJiaGe,KaiPanJiaGe,sid from [GuPiao] order by sid"
		set rs=gp_conn.execute(sql)
		do while not rs.eof 
			if (rs(0)-rs(1))/rs(1)>-0.15 then 
				sql="update [GuPiao] set DangQianJiaGe=DangQianJiaGe*0.97 where sid="&rs(2)
				gp_conn.execute sql 
			end if
			rs.movenext
		loop
		mess="<font color=#3399FF>���÷�չ���ȣ�ͨ���������أ���Ʊ�г����й�Ʊ�۸�<font color=#FF6699>�´�</font> 3 %</font>"
		gp_conn.execute("insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )")
	end if
	Response.Redirect "z_gp_GuPiao.asp"
end if%>