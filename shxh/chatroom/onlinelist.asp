<% 
Response.CacheControl="no-cache" 
Response.AddHeader "pragma","no-cache" 
Response.Expires=-1 
bgcolor=Application("Ba_jxqy_backgroundcolor") 
bgimage=Application("Ba_jxqy_backgroundimage") 
chatroomsn=session("Ba_jxqy_userchatroomsn") 
username=session("Ba_jxqy_username") 
onlinelist=Application("Ba_jxqy_onlinelist") 
onlinelistubd=ubound(onlinelist) 
onlinenum=0 
for i=1 to onlinelistubd step 8 
if chatroomsn=onlinelist(i+5) then 
onlinenum=onlinenum+1 
if onlinelist(i+1)="男" then 
sex="<img src='../images/boy.gif' width='15' height='15'>" 
else 
sex="<img src='../images/girl.gif' width='15' height='15'>" 
end if 
msg=msg&sex&" <a href='javascript:parent.chgsendto("&chr(34)&onlinelist(i)&chr(34)&");' target='talkfrm' onmouseover=""window.status='选择说话或动作对象';return true;"" onmouseout=""window.status='';return true;"" title='选择说话或动作对象'><font size=2 color="&onlinelist(i+6)&">"&onlinelist(i)&"</a> <font color=336600 size=2>"&onlinelist(i+2)&"</font></font><br>" 
end if
next 
Response.Write "<html><head><link rel='stylesheet' href='css.css'><script language=javasctipt >var n=0;function findstr(str){var textrange,i,found;if(str==''){return false;}textrange=window.document.body.createTextRange();for(i=0;i<=n&&(found=textrange.findText(str))!=false;i++){textrange.moveStart('character',1);textrange.moveEnd('textedit');}if(found){textrange.moveStart('character',-1);textrange.findText(str);textrange.select();textrange.scrollIntoView();n++;}else{if(n>0){n=0;findstr(str);}else{alert('您要的名字没有找到！');]return false;}setTimeout('this.location.reload();',180000);</script></head><body background='"&bgimage&"' bgcolor='"&bgcolor&"' oncontextmenu=self.event.returnValue=false leftmargin=0 topmargin=0 marginwidth=3 marginheight=0><div align=center><img src=../images/up.gif border=0>切换帮派房间<img src=../images/up.gif border=0><br><font size='4' color='#CC0000' face='幼圆'><b>"&Application("Ba_jxqy_systemname"&chatroomsn)&"</b></font><br><font color=FF0000>"&onlinenum&"</font>/<font color=008800>"&Application("Ba_jxqy_allonlinenum")&"</font>人在线<hr></div><form name='search' onSubmit='return findstr(this.searchstr.value);'><input name='searchstr' type='text' size=8 onChange='javasctipt :n=0;'class='norstyle'><input type='submit' value='查找' class='norstyle' id='submit'1 name='submit'1></form><img src='../images/words.gif' width='15' height='15'> <a href='javasctipt :parent.chgsendto("&chr(34)&"大家"&chr(34)&");' target='talkfrm' onmouseover=""window.status='选择说话或动作对象';return true;"" onmouseout=""window.status='';return true;"" title='选择说话或动作对象'><font size=2 color=0000ff>大家</font></a><br>"&msg&"</body></html>" 
%> 