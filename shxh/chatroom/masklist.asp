
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=-1
bgcolor=Application("Ba_jxqy_backgroundcolor")
bgimage=Application("Ba_jxqy_backgroundimage")
username=session("Ba_jxqy_username")
chatroomsn=Session("Ba_jxqy_userchatroomsn")
msg="<html><head><link rel='stylesheet' href='style1.css'><script language=javascript>var n=0;function findstr(str){var textrange,i,found;if(str==''){return false;}textrange=window.document.body.createTextRange();for(i=0;i<=n&&(found=textrange.findText(str))!=false;i++){textrange.moveStart('character',1);textrange.moveEnd('textedit');}if(found){textrange.moveStart('character',-1);textrange.findText(str);textrange.select();textrange.scrollIntoView();n++;}else{if(n>0){n=0;findstr(str);}else{alert('您要的名字没有找到！');}}return false;}</script></head><body background='"&bgimage&"' bgcolor='"&bgcolor&"' oncontextmenu=self.event.returnValue=false leftmargin=0 topmargin=0 marginwidth=3 marginheight=0><div align=center><font size='4' color='#CC0000' face='幼圆'><b>"&Application("Ba_jxqy_systemname"&chatroomsn)&"</b></font><br><font color=FF0000>"&Application("Ba_jxqy_onlinenum"&chatroomsn)&"</font>/<font color=008800>"&Application("Ba_jxqy_allonlinenum")&"</font>人在线<hr></div><form name='search' onSubmit='return findstr(this.searchstr.value);'><input name='searchstr' type='text' size=8 onChange='javascript:n=0;'class='norstyle'><input type='submit' value='查找' class='norstyle' id='submit'1 name='submit'1></form>屏蔽讨厌鬼<br>"
masklist=" "&Request.QueryString("masklist")&" "
onlinelist=Application("Ba_jxqy_onlinelist")
onlinelistubd=ubound(onlinelist)
i=1
do while i<onlinelistubd
	if onlinelist(i)<>username and instr(masklist," "&onlinelist(i)&" ")=0 then
		msg=msg&"<input type=checkbox onclick=javascript:parent.mask('"&onlinelist(i)&"');> "&onlinelist(i)&"<br>"
	elseif onlinelist(i)<>username and instr(masklist," "&onlinelist(i)&" ")<>0 then
		msg=msg&"<input type=checkbox onclick=javascript:parent.mask('"&onlinelist(i)&"'); checked> "&onlinelist(i)&"<br>"
	end if	
	i=i+8
loop
msg=msg&"</body>"
Response.Write msg
%>
