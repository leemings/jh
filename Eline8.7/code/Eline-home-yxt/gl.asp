<%@ LANGUAGE=VBScript codepage ="936" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

sjjh_name=Session("sjjh_name")
sjjh_grade=Session("sjjh_grade")
sjjh_jhdj=Session("sjjh_jhdj")
sjjh_name=sjjh_name
grade=Int(sjjh_grade)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("sjjh_usermdb")
dim tmprs
tmprs=conn.execute("Select count(*) As ���� from �û� where times<3 and CDate(lasttime)<now()-10")
musers=tmprs("����")
set tmprs=nothing
if isnull(musers) then musers=0
user1=musers
tmprs=conn.execute("Select count(*) As ���� from �û� where allvalue<=5 and CDate(lasttime)<now()-10")
musers=tmprs("����")
set tmprs=nothing
if isnull(musers) then musers=0
user2=musers
tmprs=conn.execute("Select count(*) As ���� from �û� where CDate(lasttime)<now()-30")
musers=tmprs("����")
set tmprs=nothing
if isnull(musers) then musers=0
user3=musers
%><title><%=Application("sjjh_chatroomname")%>���ݿ�ά������</title>
<body background="../JHIMG/Bk_hc3w.gif">
<div align="center">
<p>���û����ݿ⣺</p>
<table width="500" border="1" bordercolor="#000000" cellspacing="0" cellpadding="1">
<tr>
<td width="85%">10��֮ǰ,����½3�ε��û��У�<font color="#0000FF"><b><%=user1%>��</b></font></td>
<td width="15%">
<div align="center"><a href="gl1.asp">ɾ��</a></div>
</td>
</tr>
<tr>
<td width="85%">10��֮ǰ�ݵ���allvalueС��5���û��У�<b><font color="#0000FF"><%=user2%>��</font></b></td>
<td width="15%">
<div align="center"><a href="gl2.asp">ɾ��</a></div>
</td>
</tr>
<tr>
<td width="85%">30���δ��½���û��У�<font color="#0000FF"><b><%=user3%>��</b></font></td>
<td width="15%">
<div align="center"><a href="gl3.asp">ɾ��</a></div>
</td>
</tr>
<tr>
<td width="85%">�Ĳ���������������޷������ṩû�е���ʹ�ô�������</td>
<td width="15%">
<div align="center"><a href="gl5.asp">����</a></div>
</td>
</tr>

</table>
  <p align="center"><BR>
    ���׳�����<%=Application("sjjh_chatroomname")%>СС����ɣ�<br>
    ����������Բ鿴����ǰ���ݿ��ʹ�������<br>
    ѡ��ɾ�������԰���һЩ�û�ɾ������ɾ��֮��ѡ�����ݿ�ѹ��������<br>
    ���ҵĽ������Ҿ�����һЩ�����ɹ���һ��11MB�����ݿ�ѹ����4MB��С��<br>
    �ڴ˲���ǰ���鱸��ԭʼ�ļ��������Ϊ���������ɵĺ�����ǲ����κ����Σ�<br>
    ����:��¼���ݿ⡢BBS���ݿ�����ر��Ŀ���ѡ����ԭʼ�ļ����ǰ����ɣ�<br>
    �������Щ�ļ�����<BR>
  </p>
</div>