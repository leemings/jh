
<%
Response.CacheControl="no-cache"
Response.AddHeader "pragma","no-cache"
Response.Expires=0
chatroomsn=session("yx8_mhjh_userchatroomsn")
onlinelist=Application("yx8_mhjh_onlinelist")
adminname=Application("yx8_mhjh_admin")
yinshen=Application("yx8_mhjh_yinshen")
onlinelistubd=ubound(onlinelist)
onlinenum=0
for i=1 to onlinelistubd step 6
   if chatroomsn=onlinelist(i+5) then
		onlinenum=onlinenum+1
		if onlinelist(i+2)="�ٸ�" then
           color="#FF0000"
        elseif onlinelist(i+1)="��" then
           color="#0000CC"
        else
        color="#008040"
       end if       
     if Application("yx8_mhjh_allonlinenum")>40 then
       if onlinelist(i)=adminname then 
       if yinshen=True then 
       msg=msg
       else  
         msg=msg&" &nbsp;<a href='javascript:parent.chgsendto("&chr(34)&onlinelist(i)&chr(34)&");' target='talkfrm' onmouseover=""window.status='ѡ��˵����������';return true;"" onmouseout=""window.status='';return true;"" title='���ɣ�"&onlinelist(i+2)&" &nbsp;�Ա�"&onlinelist(i+1)&"'><font color="&color&">"&onlinelist(i)&"</font></a>&nbsp;<a href='#' onclick="&chr(34)&"javascript:window.open('show.asp?un="&onlinelist(i)&"','','left=0,top=10,width='+screen.availwidth*2/4+',height='+screen.availheight*3/6+',status=yes,scrollbars=yes,resizable=no')"&chr(34)&" onmouseover="&chr(34)&"window.status='�鿴"&onlinelist(i)&"�ĸ�������';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"�鿴"&onlinelist(i)&"�ĸ�������"&chr(34)&">"&onlinelist(i+2)&"</a><br>"
       end if
       else 
         msg=msg&" &nbsp;<a href='javascript:parent.chgsendto("&chr(34)&onlinelist(i)&chr(34)&");' target='talkfrm' onmouseover=""window.status='ѡ��˵����������';return true;"" onmouseout=""window.status='';return true;"" title='���ɣ�"&onlinelist(i+2)&" &nbsp;�Ա�"&onlinelist(i+1)&"'><font color="&color&">"&onlinelist(i)&"</font></a>&nbsp;<a href='#' onclick="&chr(34)&"javascript:window.open('show.asp?un="&onlinelist(i)&"','','left=0,top=10,width='+screen.availwidth*2/4+',height='+screen.availheight*3/6+',status=yes,scrollbars=yes,resizable=no')"&chr(34)&" onmouseover="&chr(34)&"window.status='�鿴"&onlinelist(i)&"�ĸ�������';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"�鿴"&onlinelist(i)&"�ĸ�������"&chr(34)&">"&onlinelist(i+2)&"</a><br>"
       end if 
     else 
       if onlinelist(i)=adminname then 
       if yinshen=True then 
       msg=msg
       else  
         msg=msg&" <img border='0' src='"&onlinelist(i+3)&"' width='16' height='16'>&nbsp;<a href='javascript:parent.chgsendto("&chr(34)&onlinelist(i)&chr(34)&");' target='talkfrm' onmouseover=""window.status='ѡ��˵����������';return true;"" onmouseout=""window.status='';return true;"" title='���ɣ�"&onlinelist(i+2)&" &nbsp;�Ա�"&onlinelist(i+1)&"'><font color="&color&">"&onlinelist(i)&"</font></a>&nbsp;<a href='#' onclick="&chr(34)&"javascript:window.open('show.asp?un="&onlinelist(i)&"','','left=0,top=10,width='+screen.availwidth*2/4+',height='+screen.availheight*3/6+',status=yes,scrollbars=yes,resizable=no')"&chr(34)&" onmouseover="&chr(34)&"window.status='�鿴"&onlinelist(i)&"�ĸ�������';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"�鿴"&onlinelist(i)&"�ĸ�������"&chr(34)&"><font color=#FF0000>"&onlinelist(i+2)&"</font></a><br>"
       end if
       else 
         msg=msg&" <img border='0' src='"&onlinelist(i+3)&"' width='16' height='16'>&nbsp;<a href='javascript:parent.chgsendto("&chr(34)&onlinelist(i)&chr(34)&");' target='talkfrm' onmouseover=""window.status='ѡ��˵����������';return true;"" onmouseout=""window.status='';return true;"" title='���ɣ�"&onlinelist(i+2)&" &nbsp;�Ա�"&onlinelist(i+1)&"'><font color="&color&">"&onlinelist(i)&"</font></a>&nbsp;<a href='#' onclick="&chr(34)&"javascript:window.open('show.asp?un="&onlinelist(i)&"','','left=0,top=10,width='+screen.availwidth*2/4+',height='+screen.availheight*3/6+',status=yes,scrollbars=yes,resizable=no')"&chr(34)&" onmouseover="&chr(34)&"window.status='�鿴"&onlinelist(i)&"�ĸ�������';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"�鿴"&onlinelist(i)&"�ĸ�������"&chr(34)&"><font color=#FF0000>"&onlinelist(i+2)&"</font></a><br>"
       end if 
     end if
   end if
next
Response.Write "<html><head><link rel='stylesheet' href='css.css'><script language=javascript>var n=0;function findstr(str){var textrange,i,found;if(str==''){return false;}textrange=window.document.body.createTextRange();for(i=0;i<=n&&(found=textrange.findText(str))!=false;i++){textrange.moveStart('character',1);textrange.moveEnd('textedit');}if(found){textrange.moveStart('character',-1);textrange.findText(str);textrange.select();textrange.scrollIntoView();n++;}else{if(n>0){n=0;findstr(str);}else{alert('��Ҫ������û���ҵ���');}}return false;}function aboutuser(corp,identity){alert('���ɣ�'+corp+'\r\n�ȼ���'+identity);}setTimeout('this.location.reload();',180000);</script></head><body style='CURSOR: url(../image/banana.ani)' bgcolor='FFDDF2' background='bg1.gif' oncontextmenu=self.event.returnValue=false leftmargin=0 topmargin=0 marginwidth=3 marginheight=0 ><div align=center><br><font size='4' color='#CC0000' face='��Բ'><b>"&Application("yx8_mhjh_systemname"&chatroomsn)&"</b></font><br><font color=FF0000>"&onlinenum&"</font>/<font color=008800>"&Application("yx8_mhjh_allonlinenum")&"</font>�����ߡ�<font color='#CC0000'>"&time()&"</font>��</div><form name='search' onSubmit='return findstr(this.searchstr.value);'><input name='searchstr' type='text' size=6 onChange='javascript:n=0;'class='norstyle'><input type='submit' value='����' class='norstyle' id='submit'1 name='submit'1></form><img src='../chatroom/image/words.gif' width='15' height='15'> <a href='javascript:parent.chgsendto("&chr(34)&"���"&chr(34)&");' target='talkfrm' onmouseover=""window.status='ѡ��˵����������';return true;"" onmouseout=""window.status='';return true;"" title='ѡ��˵����������'><font color=#CC0000>���</font></a><br><img src='../face/nan024.gif' width='18' height='18'>&nbsp;<a href='javascript:parent.chgsendto("&chr(34)&"������"&chr(34)&");' target='talkfrm' onmouseover=""window.status='������Լ���״̬,��������������,С�ı��Ҵ�����!';return true;"" onmouseout=""window.status='';return true;"" title='������Լ���״̬,��������������,С�ı��Ҵ�����!'><font color=#CC0000>������</font></a>&nbsp;<a href='#' onclick="&chr(34)&"javascript:window.open('show1.asp','','left=0,top=10,width='+screen.availwidth*2/4+',height='+screen.availheight*3/6+',status=yes,scrollbars=yes,resizable=no')"&chr(34)&" onmouseover="&chr(34)&"window.status='�鿴[������]�����ú�ʹ�÷���';return true;"&chr(34)&" onmouseout="&chr(34)&"window.status='';return true;"&chr(34)&" title="&chr(34)&"�鿴[������]���ú�ʹ�÷���"&chr(34)&"><font color=#FF00FF>������</font></a><br>"&msg&"</body></html>"
%>
