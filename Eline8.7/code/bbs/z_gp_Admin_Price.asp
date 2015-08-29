<!--#include file="conn.asp"-->
<!--#include file="inc/const.asp"-->
<!--#include file="z_gp_conn.asp"-->
<!--#include file="z_gp_Const.asp"-->
<%if not master then
	errmess="您没有执行此操作的权限！"
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
		mess="<font color=#3399FF>政府入市干预，向股票市场投入巨额资金，所有股票价格<font color=#FF6699>上扬</font> 4 %</font>"
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
		mess="<font color=#3399FF>经济发展过热，通货膨胀严重，股票市场所有股票价格<font color=#FF6699>下挫</font> 3 %</font>"
		gp_conn.execute("insert into RndEvent(content,addtime) values('"&mess&"','"&now()&"' )")
	end if
	Response.Redirect "z_gp_GuPiao.asp"
end if%>