<SELECT onchange="if(this.options[this.selectedIndex].value!=''){showfont(this.options[this.selectedIndex].value);this.options[0].selected=true;}else {this.selectedIndex=0;}" name=font>
<option value="����" selected>��������</option>
<OPTION>========</OPTION>
<option value="����">����</option>
<option value="����_GB2312">����</option>
<option value="����">����</option>
<option value="����">����</option>
<OPTION value=Arial>Arial</OPTION> 
<OPTION value="Century Gothic">Century Gothic</OPTION> 
<OPTION value="Comic Sans MS">Comic Sans MS</OPTION>
<OPTION value="Courier New">Courier New</OPTION>
<OPTION value=Impact>Impact</OPTION>
<OPTION value=Tahoma>Tahoma</OPTION>
<OPTION value="Times New Roman">Times New Roman</OPTION>
<OPTION value="Script MT Bold">Script MT Bold</OPTION>
</SELECT>&nbsp;
<select name="size" onChange="if(this.options[this.selectedIndex].value!=''){showsize(this.options[this.selectedIndex].value);this.options[0].selected=true;}else {this.selectedIndex=0;}">
<option value="3" selected>�����С</option>
<OPTION>=====</OPTION>
<option value="1">1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
<option value="7">7</option>
</select><!--&nbsp;
<select name="dongzuo" onChange="insertsmilie(this.options[this.selectedIndex].value)">
<OPTION value="" selected>����</OPTION>
<OPTION>===</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/�к�">�к�</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/�ȴ�">�ȴ�</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/��Ҫ">��Ҫ</OPTION>
<OPTION value="/ȥ��">ȥ��</OPTION>
<OPTION value="/��Ц">��Ц</OPTION>
<OPTION value="/ʹ��">ʹ��</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/ʧ��">ʧ��</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/�ȿ�">�ȿ�</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/ȭ��">ȭ��</OPTION>
<OPTION value="/�Ҳ�">�Ҳ�</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/ɵЦ">ɵЦ</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/��ˮ">��ˮ</OPTION>
<OPTION value="/��̬">��̬</OPTION>
<OPTION value="/ƴ��">ƴ��</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/ͬ��">ͬ��</OPTION>
<OPTION value="/KISS">KISS</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/�ε�">�ε�</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/��ϲ">��ϲ</OPTION>
<OPTION value="/��">��</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/����">����</OPTION>
<OPTION value="/ԭ��">ԭ��</OPTION>
<OPTION value="/�ж�">�ж�</OPTION>
<OPTION value="/��л">��л</OPTION>
<OPTION value="/ҡͷ">ҡͷ</OPTION>
<OPTION value="/��ͷ">��ͷ</OPTION>
<OPTION value="/ӵ��">ӵ��</OPTION>
<OPTION value="/�ϣ�">�ϣ�</OPTION>
<OPTION value="/��ӭ">��ӭ</OPTION>
</select>--><%
if boardid>0 or ucase(right(Request.ServerVariables("SCRIPT_NAME"),14))="Z_FASTPOST.ASP" then
if boardid=0 then
	redim Board_Setting(63)
	dim Referer
	Referer=(ucase(right(Request.ServerVariables("SCRIPT_NAME"),14))="Z_FASTPOST.ASP")
	Board_Setting(10)=Referer:Board_Setting(11)=Referer
	Board_Setting(12)=Referer:Board_Setting(13)=Referer
	Board_Setting(14)=Referer:Board_Setting(15)=Referer
	Board_Setting(23)=Referer:Board_Setting(53)=Referer
	Board_Setting(54)=Referer:Board_Setting(55)=Referer
	Board_Setting(56)=Referer:Board_Setting(57)=Referer
	Board_Setting(58)=Referer:Board_Setting(59)=Referer
	Board_Setting(60)=Referer:Board_Setting(61)=Referer
	Board_Setting(62)=Referer:Board_Setting(63)=Referer
end if%>&nbsp;
<select onchange="if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;this.options[0].selected=true;}" name="teshutie">
<OPTION selected>��������</OPTION>
<OPTION>=====</OPTION>
<%if Board_Setting(54) then%>
	<OPTION value="javascript:usermem()">��Ա��</OPTION>
<%end if
if Board_Setting(59) then%>
	<OPTION value="javascript:usergroup()">������</OPTION>
<%end if
if Board_Setting(58) then%>
	<OPTION value="javascript:userage()">������</OPTION>
<%end if
if Board_Setting(56) then%>
	<OPTION value="javascript:sex()">�Ա���</OPTION>
<%end if
if Board_Setting(57) then%>
	<OPTION value="javascript:uservip()">�߼���</OPTION>
<%end if
if Board_Setting(53) then%>
	<OPTION value="javascript:sec()">������</OPTION>
<%end if
if Board_Setting(55) then%>
	<OPTION value="javascript:spec()">������</OPTION>
<%end if
if Board_Setting(10) then%>
	<OPTION value="javascript:money()">��Ǯ��</OPTION>
<%end if
if Board_Setting(11) then%>
	<OPTION value="javascript:point()">������</OPTION>
<%end if
if Board_Setting(12) then%>
	<OPTION value="javascript:usercp()">������</OPTION>
<%end if
if Board_Setting(13) then%>
	<OPTION value="javascript:power()">������</OPTION>
<%end if
if Board_Setting(14) then%>
	<OPTION value="javascript:article()">������</OPTION>
<%end if
if Board_Setting(60) then%>
	<OPTION value="javascript:timel()">��ʾ��</OPTION>
<%end if
if Board_Setting(61) then%>
	<OPTION value="javascript:timex()">������</OPTION>
<%end if
if Board_Setting(15) then%>
	<OPTION value="javascript:replyview()">�ظ���</OPTION>
<%end if
if Board_Setting(23) then%>
	<OPTION value="javascript:usemoney()">������</OPTION>
<%end if%>
</select>&nbsp;
<select name=fastpost onchange=insertsmilie(this.options[this.selectedIndex].value)>
<OPTION selected>���ٷ���ѡ��</OPTION>
<OPTION>========</OPTION>
<%dim rsfastpost
set rsfastpost=connfastpost.execute("select LabelName,LabelContent from fastpost where username='"&trim(membername)&"'")
if rsfastpost.eof or rsfastpost.bof then%>
	<OPTION>���޿��ٷ���</OPTION>
<%else
	do while not rsfastpost.eof%>
		<OPTION value="<%=rsfastpost(1)%>"><%=rsfastpost(0)%></OPTION>
		<%rsfastpost.movenext
	loop
end if
rsfastpost.close
set rsfastpost=nothing%>
</select>
<%end if%>
<br>
<img onclick=Cbold() src="<%=Forum_info(7)&Forum_ubb(0)%>" width="22" height="22" alt="����" border="0"><img onclick=Citalic() src="<%=Forum_info(7)&Forum_ubb(1)%>" width="23" height="22" alt="б��" border="0"><img onclick=Cunder() src="<%=Forum_info(7)&Forum_ubb(2)%>" width="23" height="22" alt="�»���" border="0"><IMG onclick=Csup() height=22 alt=�ϱ��� src="<%=Forum_info(7)%>sup.gif" width=23 border=0><IMG onclick=Csub() height=22 alt=�±��� src="<%=Forum_info(7)%>sub.gif" width=23 border=0><img onclick=Ccenter() src="<%=Forum_info(7)&Forum_ubb(3)%>" width="22" height="22" alt="����" border="0"><img onclick=Curl() src="<%=Forum_info(7)&Forum_ubb(4)%>" width="22" height="22" alt="��������" border="0"><img onclick=Cemail() src="<%=Forum_info(7)&Forum_ubb(5)%>" width="23" height="22" alt="Email����" border="0"><img onclick=Cimage() src="<%=Forum_info(7)&Forum_ubb(6)%>" width="23" height="22" alt="ͼƬ" border="0"><img onclick=Cswf() src="<%=Forum_info(7)&Forum_ubb(7)%>" width="23" height="22" alt="FlashͼƬ" border="0"><img onclick=Cdir() src="<%=Forum_info(7)&Forum_ubb(8)%>" width="23" height="22" alt="Shockwave�ļ�" border="0"><img onclick=Crm() src="<%=Forum_info(7)&Forum_ubb(9)%>" width="23" height="22" alt="realplay��Ƶ�ļ�" border="0"><img onclick=Cwmv() src="<%=Forum_info(7)&Forum_ubb(10)%>" width="23" height="22" alt="Media Player��Ƶ�ļ�" border="0"><img onclick=Cmov() src="<%=Forum_info(7)&Forum_ubb(11)%>" width="23" height="22" alt="QuickTime��Ƶ�ļ�" border="0"><img onclick=Cquote() src="<%=Forum_info(7)&Forum_ubb(12)%>" width="23" height="22" alt="����" border="0"><IMG onclick=Cmark() height=22 alt=������ˮӡ src="<%=Forum_info(7)%>water.gif" width=23 border=0><IMG onclick=Cfly() height=22 alt=������ src="<%=Forum_info(7)&Forum_ubb(13)%>" width=23 border=0><IMG onclick=Cmarquee() height=22 alt=�ƶ��� src="<%=Forum_info(7)&Forum_ubb(14)%>" width=23 border=0><IMG onclick=Cguang() height=22 alt=������ src="<%=Forum_info(7)&Forum_ubb(15)%>" width=23 border=0><IMG onclick=Cying() height=22 alt=��Ӱ�� src="<%=Forum_info(7)&Forum_ubb(16)%>" width=23 border=0><IMG onclick=openScript('smiles.asp',500,600) src="<%=Forum_info(7)&Forum_ubb(17)%>" border=0 alt=�鿴���������ͼ��><img onclick=Csound() src="<%=Forum_info(7)&Forum_ubb(18)%>" width="23" height="22" alt="��������" border="0">
<br>
<SCRIPT language=JavaScript src="inc/z_color.js"></SCRIPT>
<SCRIPT language=JavaScript>
	var height1 = 10;	// define the height of the color bar
	var pas = 36;			// define the number of color in the color bar
	var width1=2;
										//define the width of the color bar here (for forum with little width put 1)
</SCRIPT>
<TABLE id=ColorPanel cellSpacing=0 cellPadding=0 border=0 >
	<TBODY>
	<TR><TD id=ColorUsed onclick="if(this.bgColor.length > 0) insertTag(this.bgColor)" vAlign=center align=middle><SCRIPT language=JavaScript>document.write('<IMG height='+height1+' width=20 border=1>');</SCRIPT></TD><SCRIPT language=JavaScript>rgb(pas,width1,height1)</SCRIPT></TR>
	</TBODY>
</TABLE>
<script language=JavaScript>
	function checkbbs(flag) {
		if(flag) {
			if(frmAnnounce.subject) {
				if(frmAnnounce.subject.value=="") {
					alert("���ӵ����ⲻӦΪ�ա�");
					return false;
				} else if(strLength(frmAnnounce.subject.value)>50) {
					alert("�������ⳤ�Ȳ��ܳ���50��");
					return false;
				}
			}
		}
		if(frmAnnounce.Content.value=="") {
			alert("û����д���ݡ�");
			return false;
		} else if(strLength(frmAnnounce.Content.value)><%
			if boardid>0 then
				response.write Clng(Board_Setting(16))
			else
				response.write "16240"
			end if%>) {
			alert("�������ݲ��ô���<%
			if boardid>0 then
				response.write Clng(Board_Setting(16))
			else
				response.write "16240"
			end if%>bytes��");
			return false;
		}
		return true;
	}
</script>