<%
sub eddulf()
	dim eddi,addmoney,addbank,addstock,sql,addcool,userWealth,userbank,userstock
	dim dbbank,dbStock,connbank,connstock,rsbank,rsstock,rs,msg
	dbbank="Data/e_Bank.asp"
	Set connbank = Server.CreateObject("ADODB.Connection")
	connbank.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbbank)
	dbStock="Data/e_Stock.asp"
	Set connstock = Server.CreateObject("ADODB.Connection")
	connstock.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbStock)
	if membername="" or not founduser then
		userWealth=0
		userbank=0
		userstock=0
	else
		set rs=server.createobject("adodb.recordset")
		sql="select * from [User] where username='"&membername&"'"
		rs.open sql,conn,1,3
		userWealth=rs("userWealth")
		set rsbank=server.createobject("adodb.recordset")
		sql="select * from bank where username='"&membername&"'"
		rsbank.open sql,connbank,1,3
		if not rsbank.eof then
			userbank=rsbank("savemoney")
		else
			userbank=0
		end if
		set rsstock=server.createobject("adodb.recordset")
		sql="select * from KeHu where ZhangHao='"&membername&"'"
		rsstock.open sql,connstock,1,3
		if not rsstock.eof then
			userstock=int(rsstock("ZiJin"))
		else
			userstock=0
		end if
	end if
	response.write "<br><table align=center cellpadding=3 cellspacing=1 class=tableborder1>"
	response.write "<tr align=center><th width=""100%"">"&forum_info(0)&"��������ϵͳ</th></tr>"
	response.write "<tr><td width=100% class=tablebody1>"
	response.write "���ڱ���̳����ʱ�������ļ�����<br><ul>"
	Randomize
	eddi=Fix(100*Rnd)
	addmoney=0
	addcool=0
	addbank=0
	addstock=0
	select case eddi
	case 1
		addmoney=50
		response.write "<li>���׵�������䷢���γ���� <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 2
		addmoney=10
		response.write "<li>���׵����������Ϳ��Ļ�Ծ���� <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 3
		addmoney=10
		addcool=5
		response.write "<li>��Ȼ�������Ϳ��Ļ�Ծ���� <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 4
		addmoney=10
		addcool=5
		response.write "<li>õ���������Ϳ��Ļ�Ծ���� <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 5
		addmoney=-1
		addcool=5
		response.write "<li>��˵��˵����������հ�齨��� <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font>��"
	case 6
		addmoney=-1
		addcool=5
		response.write "<li>�漼��������������հ�齨��� <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font>��"
	case 7
		addmoney=-1
		addcool=5
		response.write "<li>֪�Ļ������������հ�齨��� <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font>��"
	case 8
		addmoney=-1
		addcool=5
		response.write "<li>�������°���������հ�齨��� <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font>��"
	case 9
		addmoney=-1
		addcool=5
		response.write "<li>�����������������հ�齨��� <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font>��"
	case 10
		addmoney=-1
		addcool=5
		response.write "<li>��˵��˵����������հ�齨��� <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font>��"
	case 11
		addmoney=-1
		addcool=5
		response.write "<li>�漼��������������հ�齨��� <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font>��"
	case 12
		addmoney=-1
		addcool=5
		response.write "<li>֪�Ļ������������հ�齨��� <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font>��"
	case 13
		addmoney=-1
		addcool=5
		response.write "<li>�������°���������հ�齨��� <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font>��"
	case 14
		addbank=-userbank/2
		response.write "<li>���е��գ������ʧ <font color=red>50%</font>��"
	case 15
		addbank=userbank/2
		response.write "<li>���г鵽�󽱣�������� <font color=red>50%</font>��"
	case 16
		addbank=-userbank*3/10
		response.write "<li>�����⼷�ң������ʧ <font color=red>30%</font>��"
	case 17
		addbank=userbank*3/10
		response.write "<li>���н���������������� <font color=red>30%</font>��"
	case 18
		addbank=-userbank/10
		response.write "<li>���е�����ڿ͹����������ʧ <font color=red>10%</font>��"
	case 19
		addbank=userbank/10
		response.write "<li>���е��Է������������ <font color=red>10%</font>��"
	case 20
		addbank=-50
		response.write "<li>������Ϣ������󣬴����ʧ <font color=red>50</font> Ԫ��"
	case 21
		addbank=50
		response.write "<li>������Ϣ������󣬴������ <font color=red>50</font> Ԫ��"
	case 22
		addstock=-userstock/2
		response.write "<li>֤ȯ�����գ��ʽ���ʧ <font color=red>50%</font>��"
	case 23
		addstock=userstock/2
		response.write "<li>֤ȯ���鵽�󽱣��ʽ����� <font color=red>50%</font>��"
	case 24
		addstock=-userstock*3/10
		response.write "<li>֤ȯ���⼷�ң��ʽ���ʧ <font color=red>30%</font>��"
	case 25
		addstock=userstock*3/10
		response.write "<li>֤ȯ�������������ʽ����� <font color=red>30%</font>��"
	case 26
		addstock=-userstock/10
		response.write "<li>֤ȯ��������ڿ͹������ʽ���ʧ <font color=red>10%</font>��"
	case 27
		addstock=userstock/10
		response.write "<li>֤ȯ�����Է������ʽ����� <font color=red>10%</font>��"
	case 28
		addstock=-50
		response.write "<li>֤ȯ����Ϣ��������ʽ���ʧ <font color=red>50</font> Ԫ��"
	case 29
		addstock=50
		response.write "<li>֤ȯ����Ϣ��������ʽ����� <font color=red>50</font> Ԫ��"
	case 30
		addmoney=-2
		addcool=3
		response.write "<li>��·����ƽ���ε������������� <font color=red>"&abs(addmoney)&"</font> Ԫ����Ǯ������ֵ�� <font color=red>"&addcool&"</font> ��"
	case 31
		addmoney=-3
		response.write "<li>�����������˴�ܣ������������ <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 32
		addmoney=-4
		response.write "<li>���ڽ�ͨ���ſ���ͣ�������� <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 33
		addmoney=-userWealth/10
		response.write "<li>�����ⲻ����·������ <font color=red>10%</font> �ı�Ǯ��"
	case 34
		addmoney=-userWealth/10
		response.write "<li>������˼��С͵�ֹ����ң��������� <font color=red>10%</font> �ĲƲ����������ٴ�����������ǮҪ�����У�"
	case 35
		addmoney=-20
		response.write "<li>�����羹Ȼ���������Ǯ����ô��û�˶�Ǯ�ء��������ڴ������������� <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 36
		addmoney=-30
		addcool=-1
		response.write "<li>���Ѿ۶ģ������� <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&abs(addcool)&"</font> ��"
	case 37
		addmoney=-10
		response.write "<li>�μ���Ϸ�������ã�һ��Ǯû�̵������� <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 38
		addmoney=-userWealth/5
		response.write "<li>���ҵ����⵽��ͽϴ�٣�°ȥ���� <font color=red>1/4</font> ���ֽ𣡣������ٴ�����������ǮҪ�����У�"
	case 39
		addmoney=-userWealth*0.10
		response.write "<li>���겻�������Ҿ�Ȼ���ֵ����ĲƲ����� <font color=red>10%</font> ��"
	case 40
		addmoney=-3
		response.write "<li>Ϊ�˷���������£������˺ܳ���·����ȥ�˳˳��� <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 41
		addmoney=-5
		response.write "<li>��ӰԺ������������5����ȥ�����볡�� <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 42
		addmoney=-10
		addcool=1
		response.write "<li>·����ؤ��������ʹʩ���� <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font> ��"
	case 43
		addmoney=-20
		addcool=2
		response.write "<li>;������Ժ�����ȥ���ޡ�һ����ȥ <font color=red>"&abs(addmoney)&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font> ��"
	case 44
		addmoney=-15
		response.write "<li>����ܳ�����СС�𻵣���ȥ�� <font color=red>"&abs(addmoney)&"</font> Ԫ������ѣ�"
	case 45
		addmoney=-30
		response.write "<li>������Ҫ����ͣ��ڴ󷹵���һ�ٺ�ȥ�� <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 46
		addmoney=-20
		response.write "<li>�㲻С�Ĳ�������ȥ�� <font color=red>"&abs(addmoney)&"</font> Ԫ�����Ʒѣ�"
	case 47
		addmoney=-userWealth*0.03
		response.write "<li>����������˰����ȡ�� <font color=red>3%</font> �ĸ�������˰��"
	case 48
		addmoney=-userWealth/10
		response.write "<li>��С���������Ľֽ��⵽��٣��������� <font color=red>10%</font> ���ֽ𣡣���������������ǮҪ�����У�"
	case 49
		addmoney=-10
		response.write "<li>�浹ù���㲻С���ھ��ֲ����� <font color=red>"&abs(addmoney)&"</font> Ԫ��"
	case 50
		Randomize
		addmoney=Fix(100*Rnd)+1
		response.write "<li>��ϲ����˴η�������� <font color=red>"&addmoney&"</font> Ԫ������ر����ͽ���"
	case 51
		addmoney=10
		response.write "<li>��ϲ����˴η�������� <font color=red>"&addmoney&"</font> Ԫ���ر�����"
	case 52
		addmoney=userWealth/10
		response.write "<li>����ү���չ��ٱ�̳�����ɲ��ˣ���ĲƲ������� <font color=red>10%</font> ��"
	case 53
		addmoney=6
		response.write "<li>���Ƽҽ��չ��ٱ�̳���������У������� <font color=red>"&addmoney&"</font> Ԫ�ĺ����"
	case 54
		addmoney=3
		response.write "<li>��������ο�ʣ������� <font color=red>"&addmoney&"</font> Ԫ��ο�ʽ�"
	case 55
		addmoney=5
		addcool=-2
		response.write "<li>���ڽ�ͷ�սǼ� <font color=red>"&addmoney&"</font> Ԫ����������Ŀڴ�������ֵ�� <font color=red>"&abs(addcool)&"</font> ��"
	case 56
		addmoney=2
		addcool=2
		response.write "<li>���ڽ�ͷ�սǼ�һ������������������ʧ����ʧ���͸��� <font color=red>"&addmoney&"</font> Ԫ�ĳ������ֵ�� <font color=red>"&addcool&"</font> ��"
	case 57
		addmoney=10
		response.write "<li>ʨ�����Ǵ�����һ�麬����ʯ�����������ɷ�ȫ����ֵ��� <font color=red>"&addmoney&"</font> Ԫ��"
	case 58
		addmoney=8
		response.write "<li>������Ļ������룬�������������� <font color=red>"&addmoney&"</font> Ԫ�Ľ�����"
	case 59
		addcool=2
		response.write "<li>���������Ϊ�������ʻ�������ֵ�� <font color=red>"&addcool&"</font> ��"
	case 60
		addmoney=10
		response.write "<li>�㱻�����ѡΪ��������֮�ǣ��ɷ� <font color=red>"&addmoney&"</font> Ԫ�Ľ���"
	case 61
		addmoney=3
		response.write "<li>��ʹ���٣��㷢�����ý���Ϊ <font color=red>"&addmoney&"</font> Ԫ��"
	case 62
		addmoney=7
		response.write "<li>�����˵صõ������ۻ����� <font color=red>"&addmoney&"</font> Ԫ��"
	case 63
		addmoney=6
		response.write "<li>��μ���������֯�����ԣ������ <font color=red>"&addmoney&"</font> Ԫ��������ֽ�"
	case 64
		addmoney=10
		response.write "<li>�����㳤�ڶԱ����֧�֣������������� <font color=red>"&addmoney&"</font> Ԫ����������"
	case 65
		addmoney=6
		response.write "<li>�����㳤�ڶԱ����֧�֣������������� <font color=red>"&addmoney&"</font> Ԫ����������"
	case 66
		addmoney=4
		response.write "<li>�����㳤�ڶԱ����֧�֣������������� <font color=red>"&addmoney&"</font> Ԫ����������"
	case 67
		addmoney=3
		response.write "<li>�����㳤�ڶԱ����֧�֣������������� <font color=red>"&addmoney&"</font> Ԫ����������"
	case 68
		addmoney=2
		response.write "<li>�����㳤�ڶԱ����֧�֣������������� <font color=red>"&addmoney&"</font> Ԫ����������"
	case 69
		addcool=1
		response.write "<li>����ӰԺ�ɷ��α�ȯ������ֵ�� <font color=red>"&addcool&"</font> ��"
	case 70
		addmoney=3
		response.write "<li>����ӰԺ�ɷ��α����㽫�볡ȯ������õ� <font color=red>"&addmoney&"</font> Ԫ��"
	case 71
		addmoney=10
		addcool=2
		response.write "<li>��Ե�޹�����������Է��������� <font color=red>"&addmoney&"</font> Ԫ������ֵ�� <font color=red>"&addcool&"</font> ��"
	case 72
		addmoney=7
		response.write "<li>��μ�������Ϸ�õ� <font color=red>"&addmoney&"</font> Ԫ�Ľ���"
	case 73
		addmoney=5
		response.write "<li>��μ�������Ϸ�õ� <font color=red>"&addmoney&"</font> Ԫ�Ľ���"
	case 74
		addmoney=3
		response.write "<li>��μ�������Ϸ�õ� <font color=red>"&addmoney&"</font> Ԫ�Ľ���"
	case 75
		addmoney=10
		addcool=-2
		response.write "<li>�������Ѿ۶�Ӯ�� <font color=red>"&addmoney&"</font> Ԫ������ֵ�� <font color=red>"&abs(addcool)&"</font> ��"
	case 76
		addmoney=20
		addcool=-1
		response.write "<li>���˼�Ǯ���㵱��ʧ���͸����㣬�յ� <font color=red>"&addmoney&"</font> Ԫ������ֵ�� <font color=red>"&abs(addcool)&"</font> ��"
	case 77
		addmoney=userWealth/4
		response.write "<li>������������������һ�쾻׬ <font color=red>25%</font> ��"
	case 78
		addmoney=1
		response.write "<li>������·�ߣ�����<font color=red>һ��Ǯ</font>���Ҳ�����������Ͱ����ŵ�����Ŀڴ���ߣ�"
	case 79
		addmoney=2
		response.write "<li>����������Ĺ����Ƶõ� <font color=red>"&addmoney&"</font> Ԫ��Ӷ��"
	case 80
		addmoney=6
		addcool=2
		response.write "<li>��·����ƽ���ε���������� <font color=red>"&addmoney&"</font> Ԫ�ĳ������ֵ�� <font color=red>"&addcool&"</font> ��"
	case else
		response.write "<li>��˴η���δ�����κμ�������ӭ���ٴη�����"
	end select
	response.write "</ul>������������лл��Ĳ��룡<br>"

	if membername="" or not founduser then
		response.write "<a href=login.asp><font color=red>������û����ʽ��¼��̳������������̳�ļ�����ֻ��<b>����һ��</b>��</font></a>"
	else
		if addmoney<>0 then
			rs("userWealth")=Fix(rs("userWealth")+addmoney)
			if rs("userWealth")<0 then
				rs("userWealth")=0
			end if
		end if
		if addcool<>0 then
			rs("userCP")=Fix(rs("userCP")+addcool)
			if rs("userCP")<0 then
				rs("userCP")=0
			end if
		end if
		rs.update
		rs.close
		set rs=nothing
		if addbank<>0 and userbank<>0 then
			rsbank("savemoney")=Fix(rsbank("savemoney")+addbank)
			if rsbank("savemoney")<0 then
				rsbank("savemoney")=0
			end if
			rsbank.update
		end if
		rsbank.close
		set rsbank=nothing
		if addstock<>0 and userstock<>0 then
			rsstock("zijin")=rsstock("zijin")+addstock
			if rsstock("zijin")<0 then
				rsstock("zongzijin")=rsstock("zongzijin")-rsstock("zijin")
				rsstock("zijin")=0
			else
				rsstock("Zongzijin")=rsstock("zongzijin")+addstock
			end if
			rsstock.update
		end if
		rsstock.close
		set rsstock=nothing
		if addmoney<>0 or addcool<>0 or addbank<>0 or addstock<>0 then
			msg=""
   if addmoney>0 then
    if msg<>"" then msg=msg&"��"
    msg=msg&"<font color=Teal>�����</font> <font color=red>"&fix(addmoney)&"</font> Ԫ�ֽ�"
   elseif addmoney<0 then
    if msg<>"" then msg=msg&"��"
    msg=msg&"<font color=Teal>��ʧ��</font> <font color=red>"&fix(abs(addmoney))&"</font> Ԫ�ֽ�"
   else
    msg=msg&"<font color=Teal>�ֽ�û��Ӱ�죡</font>"
   end if
   if addbank>0 then
    if msg<>"" then msg=msg&"��"
    msg=msg&"<font color=Teal>�����</font> <font color=red>"&fix(addbank)&"</font> Ԫ���"
   elseif addbank<0 then
    if msg<>"" then msg=msg&"��"
    msg=msg&"<font color=Teal>��ʧ��</font> <font color=red>"&fix(abs(addbank))&"</font> Ԫ���"
   else
    msg=msg&"<font color=Teal>���û��Ӱ�죡</font>"
   end if
   if addstock>0 then
    if msg<>"" then msg=msg&"��"
    msg=msg&"<font color=Teal>�����</font> <font color=red>"&fix(addstock)&"</font> Ԫ�ʽ�"
   elseif addstock<0 then
    if msg<>"" then msg=msg&"��"
    msg=msg&"<font color=Teal>��ʧ��</font> <font color=red>"&fix(abs(addstock))&"</font> Ԫ�ʽ�"
   else
    msg=msg&"<font color=Teal>�ʽ�û��Ӱ�죡</font>"
   end if
   if addcool>0 then
    if msg<>"" then msg=msg&"��"
    msg=msg&"<font color=Teal>������</font> <font color=red>"&fix(addcool)&"</font> ������ֵ"
   elseif addcool<0 then
    if msg<>"" then msg=msg&"��"
    msg=msg&"<font color=Teal>������</font> <font color=red>"&fix(abs(addcool))&"</font> ������ֵ"
   else
    msg=msg&"<font color=Teal>����ֵû��Ӱ�죡</font>"
   end if
			if msg<>"" then msg=msg&"��"
			response.write "<hr width=90% size=1 color=#AAAAAA><div align=center>��˴μ������й�"&msg&"</div>"
		end if
	end if
	response.write "</td></tr></table><br>"
	connbank.close
	Set connbank = Nothing
	connstock.close
	Set connstock = Nothing
end sub
%>