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
			if especial<>"死亡" then newaberrant=newaberrant&""&especial&"程度加深"&especialtime
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
		if especial<>"死亡" then newaberrant=newaberrant&""&especial&especialtime&"秒"
	end if
	Application.Lock
	Application("yx8_mhjh_aberrantlist")=newaberrantlist
	Application("yx8_mhjh_aberrantname")=newaberrantname
	Application.UnLock
	erase newaberrantlist
	erase aberrantlist
	if especial="麻痹" then
		newaberrant=newaberrant&"<bgsound src='../mid/dianxu.wav' loop=1>"
	elseif especial="死亡" then
		lastlogintime=dateadd("s",especialtime,nowtime)
		lastlogintimetype="#"&month(lastlogintime)&"/"&day(lastlogintime)&"/"&year(lastlogintime)&" "&hour(lastlogintime)&":"&minute(lastlogintime)&":"&second(lastlogintime)&"#"
		conn.execute "update 用户 set 状态='死亡',最后登录时间="&lastlogintimetype&",被杀=被杀+1,杀人者='"&un&"' where 姓名='"&st&"'"
		newaberrant=newaberrant&"<bgsound src='../mid/oh_no.wav' loop=1>"
	end if
end function
%>
