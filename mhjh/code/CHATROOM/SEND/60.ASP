<%
function newaberrant(st,un,especial,especialtime)
	nowtime=now()
	aberrantlist=Application("yx8_mhjh_aberrantlist")
	aberrantlistubd=ubound(aberrantlist)
	dim newaberrantlist()
	newaberrantname=";"
	j=1
	addsucf=false
	for i=1 to aberrantlistubd step 4
		if addsucf=false and aberrantlist(i+2)=especial and aberrantlist(i)=st then
			redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
			newaberrantlist(j)=st
			newaberrantlist(j+1)=un
			newaberrantlist(j+2)=especial
			newaberrantlist(j+3)=dateadd("s",especialtime,aberrantlist(i+3))
			newaberrantname=newaberrantname&aberrantlist(i)&";"
			j=j+4
			addsucf=true
			if especial<>"����" then newaberrant=newaberrant&""&especial&"�̶ȼ���"&especialtime
		elseif datediff("s",aberrantlist(i+3),timetmp)<0 then
			redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
			newaberrantlist(j)=aberrantlist(i)
			newaberrantlist(j+1)=aberrantlist(i+1)
			newaberrantlist(j+2)=aberrantlist(i+2)
			newaberrantlist(j+3)=aberrantlist(i+3)
			newaberrantname=newaberrantname&aberrantlist(i)&";"
			j=j+4
		end if
	next
	if addsucf=false then
		redim preserve newaberrantlist(j),newaberrantlist(j+1),newaberrantlist(j+2),newaberrantlist(j+3)
		newaberrantlist(j)=st
		newaberrantlist(j+1)=un
		newaberrantlist(j+2)=especial
		newaberrantlist(j+3)=dateadd("s",especialtime,nowtime)
		newaberrantname=newaberrantname&st&";"
		if especial<>"����" then newaberrant=newaberrant&""&especial&especialtime&"��"
	end if
	Application.Lock
	Application("yx8_mhjh_aberrantlist")=newaberrantlist
	Application("yx8_mhjh_aberrantname")=newaberrantname
	Application.UnLock
	erase newaberrantlist
	erase aberrantlist
	if especial="���" then
		newaberrant=newaberrant&"<bgsound src='../mid/dianxu.wav' loop=1>"
	elseif especial="����" then
		lastlogintime=dateadd("s",especialtime,nowtime)
		lastlogintimetype="#"&month(lastlogintime)&"/"&day(lastlogintime)&"/"&year(lastlogintime)&" "&hour(lastlogintime)&":"&minute(lastlogintime)&":"&second(lastlogintime)&"#"
		conn.execute "update �û� set ״̬='����',����¼ʱ��="&lastlogintimetype&",��ɱ=��ɱ+1,ɱ����='"&un&"' where ����='"&st&"'"
		newaberrant=newaberrant&"<bgsound src='../mid/oh_no.wav' loop=1>"
	end if
end function
%>
