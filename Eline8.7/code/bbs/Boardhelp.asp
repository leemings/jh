<!--#include file=conn.asp-->
<!-- #include file=inc/const.asp -->
<%if boardid=0 then
	stats="��̳�ܰ���"
	call nav()
	call head_var(2,0,"","")
else
	stats="�������"
	call nav()
	call head_var(1,BoardDepth,0,0)
end if
call main()
call footer()

sub main()%>
	<table cellspacing=1 cellpadding=3 align=center class=tableborder1>
		<tr> 
			<td class=tablebody2 width=100% align=center><A HREF=#point>��������</A> | <A HREF=#grade>�ȼ�����</A> | <A HREF=#ubb>UBB�﷨</A> | <A HREF=#ubb1>UBB����</A> | <A HREF=#price>�۸�����</A></td>
		</tr>
	  <tr> 
	    <th align=left>&nbsp;&nbsp;A. <A name=point>��������</A></th>
	  </tr> 
		<tr> 
			<td width=100% class=tablebody1>&nbsp;&nbsp;����̳ע�ᡢ��¼�����������������뾫����ɾ�����ӵȲ������û���ֵ��Ӱ�죬�����ɸ����û���������������������Ĭ��ֵ���ܰ����ɶ��û���������ֱ�Ӳ�����<BR>
				<BR>&nbsp;&nbsp;��һ��<B>��Ǯ</B><BR>&nbsp;&nbsp;ע���Ǯ����<font color=<%=Forum_body(8)%>><%=Forum_user(0)%></font>&nbsp;��¼���ӽ�Ǯ��<font color=<%=Forum_body(8)%>><%=Forum_user(4)%></font>&nbsp;�������ӽ�Ǯ��<font color=<%=Forum_body(8)%>><%=Forum_user(1)%></font>&nbsp;�������ӽ�Ǯ��<font color=<%=Forum_body(8)%>><%=Forum_user(2)%></font>&nbsp;�������ӽ�Ǯ��<font color=<%=Forum_body(8)%>><%=Forum_user(15)%></font>&nbsp;ɾ�����ٽ�Ǯ��<font color=<%=Forum_body(8)%>><%=Forum_user(3)%></font>&nbsp;ͶƱ���ӽ�Ǯ��<font color=<%=Forum_body(8)%>><%=Forum_user(40)%></font><BR>
				<BR>&nbsp;&nbsp;������<B>����</B><BR>&nbsp;&nbsp;ע�ᾭ������<font color=<%=Forum_body(8)%>><%=Forum_user(5)%></font>&nbsp;��¼���Ӿ��飺<font color=<%=Forum_body(8)%>><%=Forum_user(9)%></font>&nbsp;�������Ӿ��飺<font color=<%=Forum_body(8)%>><%=Forum_user(6)%></font>&nbsp;�������Ӿ��飺<font color=<%=Forum_body(8)%>><%=Forum_user(7)%></font>&nbsp;�������Ӿ��飺<font color=<%=Forum_body(8)%>><%=Forum_user(17)%></font>&nbsp;ɾ�����پ��飺<font color=<%=Forum_body(8)%>><%=Forum_user(8)%></font>&nbsp;ͶƱ���Ӿ���ֵ��<font color=<%=Forum_body(8)%>><%=Forum_user(41)%></font><BR>
				<BR>&nbsp;&nbsp;������<B>����</B><BR>&nbsp;&nbsp;ע����������<font color=<%=Forum_body(8)%>><%=Forum_user(10)%></font>&nbsp;��¼����������<font color=<%=Forum_body(8)%>><%=Forum_user(14)%></font>&nbsp;��������������<font color=<%=Forum_body(8)%>><%=Forum_user(11)%></font>&nbsp;��������������<font color=<%=Forum_body(8)%>><%=Forum_user(12)%></font>&nbsp;��������������<font color=<%=Forum_body(8)%>><%=Forum_user(16)%></font>&nbsp;ɾ������������<font color=<%=Forum_body(8)%>><%=Forum_user(13)%></font>&nbsp;ͶƱ��������ֵ��<font color=<%=Forum_body(8)%>><%=Forum_user(42)%></font><BR>
			</td>
		</tr>
	  <tr> 
	    <th align=left>&nbsp;&nbsp;B. <A name=grade>�ȼ�����</A></th>
	  </tr> 
	  <tr>
	  	<td class=tablebody1>
	  		<table cellspacing=1 cellpadding=6 border=0>
				  <tr> 
				    <td colspan=2>����Ϊ����̳��Ӧ�ȼ��������£��Լ���Ӧ�ĵȼ�ͼƬ��</td>
				  </tr>
					<%set rs=conn.execute("select * from usertitle where not Minarticle=-1 order by MinArticle")
					do while not rs.eof
						response.write "<tr><td valign=middle>�������������� " & rs("usertitle") & " ��Ҫ " & rs("MinArticle") & " ������ �ȼ���־Ϊ </td><td valign=middle><img src=pic/"&rs("titlepic")&"></td></tr>" 
						rs.movenext
					loop
					rs.close
					set rs=nothing%>
					<tr><td valign=middle>�����������  Ϊ����Ա�趨�����Խ����ض����档</td><td valign=middle><img src=pic/<%=conn.execute("select top 1 TitlePic from usertitle where usertitleid=18")(0)%>></td></tr>
					<tr><td valign=middle>������������  Ϊ����Ա�趨�����Զ���������̳���ӽ��й���</td><td valign=middle><img src=pic/<%=conn.execute("select top 1 TitlePic from usertitle where usertitleid=19")(0)%>></td></tr>
					<tr><td valign=middle>����������������  Ϊ����Ա�趨�����Զ�������̳���ӽ��й���</td><td valign=middle><img src=pic/<%=conn.execute("select top 1 TitlePic from usertitle where usertitleid=20")(0)%>></td></tr>
					<tr><td valign=middle>������������Ա  ��ӵ����̳ȫ��Ȩ�ޡ�</td><td valign=middle><img src=pic/<%=conn.execute("select top 1 TitlePic from usertitle where usertitleid=21")(0)%>></td></tr>
				</table>
			</td>
	  </tr> 
	  <tr> 
	    <th align=left>&nbsp;&nbsp;C. <A name=ubb>UBB�﷨</A></th>
	  </tr> 
	  <tr> 
	    <td class=tablebody1>
				<p>&nbsp;&nbsp;��̳�����ɹ���Ա�����Ƿ�֧��UBB��ǩ��UBB��ǩ���ǲ�����ʹ��HTML�﷨������£�ͨ����̳������ת��������������֧���������õġ���Σ���Ե�HTMLЧ����ʾ������Ϊ����ʹ��˵����
				<p>&nbsp;&nbsp;<font color=red>[B]</font><b>����</b><font color=red>[/B]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪ����Ч����
				<p>&nbsp;&nbsp;<font color=red>[I]</font><i>����</i><font color=red>[/I]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪб��Ч����
				<p>&nbsp;&nbsp;<font color=red>[U]</font><u>����</u><font color=red>[/U]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ�����ʾΪ�»���Ч����
				<p>&nbsp;&nbsp;<font color=red>[align=center]</font>����<font color=red>[/align]</font>�������ֵ�λ�ÿ��������������Ҫ���ַ���centerλ��center��ʾ���У�left��ʾ����right��ʾ���ҡ�
				<p>&nbsp;&nbsp;<font color=red>[URL]</font><A HREF=HTTP://BBS.51ELINE.COM>HTTP://BBS.51ELINE.COM</A><font color=red>[/URL]</font>
				<P>&nbsp;&nbsp;<font color=red>[URL=HTTP://BBS.51ELINE.COM]</font><A HREF=HTTP://BBS.51ELINE.COM>E����̳</A><font color=red>[/URL]</font>�������ַ������Լ��볬�����ӣ��������Ӿ����ַ�����������ӡ�
				<p>&nbsp;&nbsp;<font color=red>[EMAIL]</font><A HREF=mailto:Eline_Email@etang.com>Eline_Email@etang.com</A><font color=red>[/EMAIL]</font>
				<P>&nbsp;&nbsp;<font color=red>[EMAIL=MAILTO:Eline_Email@etang.com]</font><A HREF=mailto:Eline_Email@etang.com>һ����</A><font color=red>[/EMAIL]</font>�������ַ������Լ����ʼ����ӣ��������Ӿ����ַ�����������ӡ�
				<P>&nbsp;&nbsp;<font color=red>[img]</font><img src=http://zhzx.jjedu.org/logo.gif><font color=red>[/img]</font>���ڱ�ǩ���м����ͼƬ��ַ����ʵ�ֲ�ͼЧ����
				<P>&nbsp;&nbsp;<font color=red>[flash]</font>Flash���ӵ�ַ<font color=red>[/Flash]</font>���ڱ�ǩ���м����FlashͼƬ��ַ����ʵ�ֲ���Flash��
				<P>&nbsp;&nbsp;<font color=red>[code]</font>����<font color=red>[/code]</font>���ڱ�ǩ��д�����ֿ�ʵ��html�б��Ч����
				<P>&nbsp;&nbsp;<font color=red>[quote]</font>����<font color=red>[/quote]</font>���ڱ�ǩ���м�������ֿ���ʵ��HTMl����������Ч����
				<P>&nbsp;&nbsp;<font color=red>[list]</font>����<font color=red>[/list]</font> <font color=red>[list=a]</font>����<font color=red>[/list]</font>  <font color=red>[list=1]</font>����<font color=red>[/list]</font>������list���Ա�ǩ��ʵ��HTMLĿ¼Ч����
				<P>&nbsp;&nbsp;<font color=red>[fly]</font>����<font color=red>[/fly]</font>���ڱ�ǩ���м�������ֿ���ʵ�����ַ���Ч������������ơ�
				<P>&nbsp;&nbsp;<font color=red>[move]</font>����<font color=red>[/move]</font>���ڱ�ǩ���м�������ֿ���ʵ�������ƶ�Ч����Ϊ����Ʈ����
				<P>&nbsp;&nbsp;<font color=red>[glow=255,red,2]</font>����<font color=red>[/glow]</font>���ڱ�ǩ���м�������ֿ���ʵ�����ַ�����Ч��glow����������Ϊ��ȡ���ɫ�ͱ߽��С��
				<P>&nbsp;&nbsp;<font color=red>[shadow=255,red,2]</font>����<font color=red>[/shadow]</font>���ڱ�ǩ���м�������ֿ���ʵ��������Ӱ��Ч��shadow����������Ϊ��ȡ���ɫ�ͱ߽��С��
				<P>&nbsp;&nbsp;<font color=red>[color=��ɫ����]</font>����<font color=red>[/color]</font>������������ɫ���룬�ڱ�ǩ���м�������ֿ���ʵ��������ɫ�ı䡣
				<P>&nbsp;&nbsp;<font color=red>[size=����]</font>����<font color=red>[/size]</font>���������������С���ڱ�ǩ���м�������ֿ���ʵ�����ִ�С�ı䡣
				<P>&nbsp;&nbsp;<font color=red>[face=����]</font>����<font color=red>[/face]</font>����������Ҫ�����壬�ڱ�ǩ���м�������ֿ���ʵ����������ת����
				<P>&nbsp;&nbsp;<font color=red>[DIR=500,350]</font>http://<font color=red>[/DIR]</font>��Ϊ����shockwave��ʽ�ļ����м������Ϊ��Ⱥͳ���
				<P>&nbsp;&nbsp;<font color=red>[RM=500,350]</font>http://<font color=red>[/RM]</font>��Ϊ����realplayer��ʽ��rm�ļ����м������Ϊ��Ⱥͳ���
				<P>&nbsp;&nbsp;<font color=red>[MP=500,350]</font>http://<font color=red>[/MP]</font>��Ϊ����Ϊmidia player��ʽ���ļ����м������Ϊ��Ⱥͳ���
				<P>&nbsp;&nbsp;<font color=red>[QT=500,350]</font>http://<font color=red>[/QT]</font>��Ϊ����ΪQuick time��ʽ���ļ����м������Ϊ��Ⱥͳ���
			</td>
	  </tr> 
	  <tr> 
	    <th align=left>&nbsp;&nbsp;D. <A name=ubb1>UBB����</A></th>
	  </tr> 
	  <tr> 
	    <td class=tablebody1>&nbsp;&nbsp; ����Ϊ����̳��UBB�﷨���ã�ͨ����Щ���ã�������֪���ڱ����淢��������Щ����ǲ���ʹ�õģ����ﻹ�����˿����û�ǩ����ʹ�õ�UBBѡ�<BR><%
	    	if boardid>0 then
	    		%>&nbsp;&nbsp;<B>�û�����</B>��<ul><!--#include file="z_Contentflag.asp"--><BR><BR></ul>&nbsp;&nbsp;<B>��������</B><BR>
	    		<ul>
					<li><%=iif(Board_Setting(54),"��Ա��������","��Ա����������")%>
					<li><%=iif(Board_Setting(59),"������������","��������������")%>
					<li><%=iif(Board_Setting(58),"������������","��������������")%>
					<li><%=iif(Board_Setting(56),"�Ա���������","�Ա�����������")%>
					<li><%=iif(Board_Setting(57),"�߼���������","�߼�����������")%>
					<li><%=iif(Board_Setting(53),"������������","��������������")%>
					<li><%=iif(Board_Setting(55),"������������","��������������")%>
					<li><%=iif(Board_Setting(10),"��Ǯ��������","��Ǯ����������")%>
					<li><%=iif(Board_Setting(11),"������������","��������������")%>
					<li><%=iif(Board_Setting(12),"������������","��������������")%>
					<li><%=iif(Board_Setting(13),"������������","��������������")%>
					<li><%=iif(Board_Setting(14),"������������","��������������")%>
					<li><%=iif(Board_Setting(60),"��ʾ��������","��ʾ����������")%>
					<li><%=iif(Board_Setting(61),"������������","��������������")%>
					<li><%=iif(Board_Setting(15),"�ظ���������","�ظ�����������")%>
					<li><%=iif(Board_Setting(23),"������������","��������������")%>
					</ul>
				<%end if
				%><BR>&nbsp;&nbsp;<B>�û�ǩ��</B>��
				<ul>
				<li>HTML��ǩ�� <%=iif(Forum_Setting(66),"����","������")%>
				<li>UBB��ǩ�� <%=iif(Forum_Setting(65),"����","������")%>
				<li>��ͼ��ǩ(����ͼƬ��flash����ý��)�� <%=iif(Forum_Setting(67),"����","������")%>
				</ul>˵��������html��ǩָ�Ƿ�����ʹ��html�﷨����ͼ��flash�Լ������ַ�ת��������UBB�﷨���ݣ���ʹ�÷����ɲ鿴UBB�﷨
			</td>
	  </tr> 
	  <tr> 
	    <th align=left>&nbsp;&nbsp;E. <A name=price>�۸�����</A></th>
	  </tr> 
	  <tr> 
	    <td class=tablebody1>
				<p>&nbsp;&nbsp;��ͨ��Ա��ȫ���Ա��裺 <font color=<%=forum_body(8)%>><%=forum_user(18)%></font> Ԫ/��
				<p>&nbsp;&nbsp;��ͨ��Ա����һ��Ա��裺 <font color=<%=forum_body(8)%>><%=forum_user(19)%></font> Ԫ/��
				<p>&nbsp;&nbsp;VIP ��Ա��ȫ���Ա��裺 <font color=<%=forum_body(8)%>><%=forum_user(20)%></font> Ԫ/��
				<p>&nbsp;&nbsp;VIP ��Ա����һ��Ա��裺 <font color=<%=forum_body(8)%>><%=forum_user(21)%></font> Ԫ/��
				<p>&nbsp;&nbsp;��ͨ��Ա�޸�ͷ�� <font color=<%=forum_body(8)%>><%=forum_user(22)%></font> Ԫ/��
				<p>&nbsp;&nbsp;VIP ��Ա�޸�ͷ�� <font color=<%=forum_body(8)%>><%=forum_user(23)%></font> Ԫ/��
				<p>&nbsp;&nbsp;��ͨ��Ա�޸�ǩ���� <font color=<%=forum_body(8)%>><%=forum_user(24)%></font> Ԫ/��
				<p>&nbsp;&nbsp;VIP ��Ա�޸�ǩ���� <font color=<%=forum_body(8)%>><%=forum_user(25)%></font> Ԫ/��
				<p>&nbsp;&nbsp;��ͨ��Ա���Ͷ��ţ� <font color=<%=forum_body(8)%>><%=forum_user(26)%></font> Ԫ/��
				<p>&nbsp;&nbsp;VIP ��Ա���Ͷ��ţ� <font color=<%=forum_body(8)%>><%=forum_user(27)%></font> Ԫ/��
				<p>&nbsp;&nbsp;��ͨ��Ա�ظ������֪ͨ�� <font color=<%=forum_body(8)%>><%=forum_user(28)%></font> Ԫ/��
				<p>&nbsp;&nbsp;VIP ��Ա�ظ������֪ͨ�� <font color=<%=forum_body(8)%>><%=forum_user(29)%></font> Ԫ/��
				<p>&nbsp;&nbsp;��ͨ��Ա�ظ����ʼ�֪ͨ�� <font color=<%=forum_body(8)%>><%=forum_user(30)%></font> Ԫ/��
				<p>&nbsp;&nbsp;VIP ��Ա�ظ����ʼ�֪ͨ�� <font color=<%=forum_body(8)%>><%=forum_user(31)%></font> Ԫ/��
				<p>&nbsp;&nbsp;��ͨ��Ա��������ע�⣺ <font color=<%=forum_body(8)%>><%=forum_user(32)%></font> Ԫ/��
				<p>&nbsp;&nbsp;VIP ��Ա��������ע�⣺ <font color=<%=forum_body(8)%>><%=forum_user(33)%></font> Ԫ/��
				<p>&nbsp;&nbsp;��ͨ��Ա�ϴ�ͷ�� <font color=<%=forum_body(8)%>><%=forum_user(34)%></font> Ԫ/KB
				<p>&nbsp;&nbsp;VIP ��Ա�ϴ�ͷ�� <font color=<%=forum_body(8)%>><%=forum_user(35)%></font> Ԫ/KB
				<p>&nbsp;&nbsp;��ͨ��Ա�ϴ��ļ��� <font color=<%=forum_body(8)%>><%=forum_user(36)%></font> Ԫ/KB
				<p>&nbsp;&nbsp;VIP ��Ա�ϴ��ļ��� <font color=<%=forum_body(8)%>><%=forum_user(37)%></font> Ԫ/KB
				<p>&nbsp;&nbsp;��ͨ��Ա�ϴ���Ƭ�� <font color=<%=forum_body(8)%>><%=forum_user(38)%></font> Ԫ/KB
				<p>&nbsp;&nbsp;VIP ��Ա�ϴ���Ƭ�� <font color=<%=forum_body(8)%>><%=forum_user(39)%></font> Ԫ/KB
			</td>
		</tr>
	</table>
	<%call activeonline()
end sub
%>