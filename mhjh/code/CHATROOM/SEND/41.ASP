<%
function chyan(co)
if chatroomsn=1 or chatroomsn=3 or Application("yx8_mhjh_systemname"&chatroomsn)=co or co="官府"  then
chyan="<marquee>"&msg&"</marquee>"
else
chyan="<font color=FF0000>【传言】</font>##你不是本派弟子，不能在这里发传言！"
end if
end function
%>