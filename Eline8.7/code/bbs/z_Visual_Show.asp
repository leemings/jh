<%
if cint(request("width"))>0 then
	response.write "<table border=0 cellspacing=0 cellpadding=0 width="""&request("width")&""" height="""&(request("width")*226/140)&""" align=center>"
	response.write "<tr>"
	response.write "<td align=center width="""&request("width")&""" height="""&(request("width")*226/140)&""" ImagePath="""&request("PicPath")&""" baseName=""images/img_visual/blank.gif"" visual="""&request("visual")&""" posX=""0"" posY=""0"" style=behavior:url(inc/z_show3.htc) localpic="""&request("localpic")&"""></td>"
	response.write "</tr>"
	response.write "</table>"
else
	response.write "<table border=0 cellspacing=0 cellpadding=0 width=""80"" height=""120"" align=center>"
	response.write "<tr>"
	response.write "<td align=center width=""80"" height=""120"" ImagePath="""&request("PicPath")&""" baseName=""images/img_visual/blank.gif"" visual="""&request("visual")&""" posX=""-30"" posY=""-30"" style=behavior:url(inc/z_show3.htc) localpic="""&request("localpic")&"""></td>"
	response.write "</tr>"
	response.write "</table>"
end if
%>