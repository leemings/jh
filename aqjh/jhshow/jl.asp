<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
on error resume next
aqjh_name=Session("aqjh_name")
search=cstr(server.HTMLEncode(Request("search")))
if Application("jhshowurl")<>"http://qqshow-item.tencent.com" then
	showgif=Application("jhshowurl") & "gif/pic/[img].gif"
	showgif1="src="& Application("jhshowurl") & "gif/'+j+'/'+showArray[i]+'.gif"
else
	showgif=Application("jhshowurl") & "/[img]/0/00/00.gif"
	showgif1="src="& Application("jhshowurl") &"/'+showArray[i]+'/'+j+'/00/00.gif"
end if
%>
<html><title>����������б�</title><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><LINK href="style.css" rel=stylesheet></head>
<body leftmargin="0" topmargin="0" background=../bg.gif>
<div align="center"><font size="+3" face="����"><strong>���������</strong></font> 
  <br>
  <table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
    <tr>
      <td width="26%" height="201" align="center" valign="top">
<table width="98%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#00215a" bordercolordark="#FFFFFF">
          <tr> 
            <td align="center"> <font color="#0000FF"><strong> 
              <%ndate=Day(date())
if ndate<=20 then%>
              ������������� 
              <%elseif ndate<=27 then%>
              �����������ѡ���� 
              <%else%>
              �����������Ʒ��ȡ 
              <%end if%>
              </strong></font> </td>
          </tr>
          <tr> 
            <td align="center"><font color="#0000FF"><strong> 
<%ndate=Day(date())
'��27����,����һ����ʱ�ж�
if ndate>27 then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("showmdb")
	rs.Open "select * from pool", conn, 1,1
	if (isnull(rs("p_�е�һ��")) or rs("p_�е�һ��")="") and (isnull(rs("p_�еڶ���")) or rs("p_�еڶ���")="") and (isnull(rs("p_Ů��һ��")) or rs("p_Ů��һ��")="") and (isnull(rs("p_Ů�ڶ���")) or rs("p_Ů�ڶ���")="") then
		rs.close
		rs.Open "select * from sai where s_�Ա�='��' Order by s_Ʊ�� DESC,s_����ʱ��", conn, 1,1
		if RS.Eof then
			nan1="��"
			nanp1=0
			nan2="��"
			nanp2=0
		else
			nan1=rs("s_����")
			nanp1=rs("s_Ʊ��")
			RS.MoveNext
			nan2=rs("s_����")
			nanp2=rs("s_Ʊ��")
		end if
		rs.close
		rs.Open "select * from sai where s_�Ա�='Ů' Order by s_Ʊ�� DESC,s_����ʱ��", conn, 1,1
		if RS.Eof then
			nv1="��"
			nvp1=0
			nv2="��"
			nvp2=0
		else
			nv1=rs("s_����")
			nvp1=rs("s_Ʊ��")
			RS.MoveNext
			nv2=rs("s_����")
			nvp2=rs("s_Ʊ��")
		end if
		rs.close
		rs.Open "select * from pool", conn, 1,3
		rs("p_�е�һ��")=nan1
		rs("p_�е�һƱ��")=nanp1
		rs("p_�е�һ�콱")=false
		rs("p_�еڶ���")=nan2
		rs("p_�еڶ�Ʊ��")=nanp2
		rs("p_�еڶ��콱")=false
		rs("p_Ů��һ��")=nv1
		rs("p_Ů��һƱ��")=nvp1
		rs("p_Ů��һ�콱")=false
		rs("p_Ů�ڶ���")=nv2
		rs("p_Ů�ڶ�Ʊ��")=nvp2
		rs("p_Ů�ڶ��콱")=false
		rs.update
	end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
if ndate<=20 then%>
</strong><font color="#FF0000">�μӴ����ô󽱣���Ҫ<a href="#" onClick="window.open('biaoming.asp','biaoming','scrollbars=no,toolbar=no,menubar=no,location=no,status=no,resizable=no,width=320,height=375')" ><strong>����</strong></a>��ͶƱʱ���<strong>21</strong>�ſ�ʼ!</font><strong> 
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from pool", conn, 1,3
if rs("p_�е�һ��")<>"" or rs("p_�еڶ���")<>"" or rs("p_Ů��һ��")<>"" or rs("p_Ů�ڶ���")<>"" then
	rs("p_ͶƱ��")=""
	rs("p_ͶƱ����")=0
	rs("p_���")=0
	rs("p_��������")=0
	rs("p_�е�һ��")=""
	rs("p_�е�һƱ��")=0
	rs("p_�е�һ�콱")=false
	rs("p_�еڶ���")=""
	rs("p_�еڶ�Ʊ��")=0
	rs("p_�еڶ��콱")=false
	rs("p_Ů��һ��")=""
	rs("p_Ů��һƱ��")=0
	rs("p_Ů��һ�콱")=false
	rs("p_Ů�ڶ���")=""
	rs("p_Ů�ڶ�Ʊ��")=0
	rs("p_Ů�ڶ��콱")=false
	rs.update
	conn.execute("delete from sai")
end if

rs.close
rs.Open "select * from pool", conn, 1,3
%>
              <br>
              ����������<font color=red><%=rs("p_��������")%></font>��<br>
              ����ۼƣ�<font color=red><%=rs("p_���")%></font>��<br>
              Ԥ�ƽ���<font color=red><%=(int(rs("p_���")*0.6))%></font>�� 
              <%rs.close
set rs=nothing
conn.close
set conn=nothing%>
              </strong></font> 
              <%elseif ndate<=27 then
%>
              <table width="100%" border="0" cellpadding="1" cellspacing="2" bordercolorlight="#00215a" bordercolordark="#FFFFFF">
                <tr align="center" bgcolor="#FFCC00"> 
                  <td>����ID</td>
                  <td>��Ʊ��</td>
                </tr>
                <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
xx=0
rs.Open "select * from sai where s_�Ա�='��' Order by s_Ʊ�� DESC,s_����ʱ��", conn, 1,1
do while Not RS.Eof
%>
                <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left" bgcolor="#FFCC00"><%=rs("s_����")%>(��)</td>
                  <td align="right" bgcolor="#FFCC00"><%=rs("s_Ʊ��")%></td>
                </tr>
                <%xx=xx+1
if xx>=5 then exit do
RS.MoveNext
loop
rs.close
xx=0
rs.Open "select * from sai where s_�Ա�='Ů' Order by s_Ʊ�� DESC,s_����ʱ��", conn, 1,1
do while Not RS.Eof
%>
                <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left"><%=rs("s_����")%>(Ů)</td>
                  <td align="right" bgcolor="#FFCC00"><%=rs("s_Ʊ��")%></td>
                </tr>
                <%xx=xx+1
if xx>=5 then exit do
RS.MoveNext
loop%>
              </table>
              <%rs.close
              rs.Open "select * from pool", conn, 1,3
              baoren=replace(rs("p_ͶƱ��"),"|","<br>")
              %>
              ����������<font color=red><%=rs("p_��������")%></font>��<br>
              ����ۼƣ�<font color=red><%=rs("p_���")%></font>��<br>
              Ԥ�ƽ���<font color=red><%=(int(rs("p_���")*0.6))%></font>��<br>
              ͶƱ������<font color=red><%=rs("p_ͶƱ����")%></font>�� <br> <br> 
              <%rs.close
set rs=nothing
conn.close
set conn=nothing
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from pool", conn, 1,1
 baoren=replace(rs("p_ͶƱ��"),"|","<br>")
%></strong></font>
              <table width="100%" border="0" cellpadding="2" cellspacing="2" bordercolorlight="#00215a" bordercolordark="#FFFFFF">
                <tr align="center" bgcolor="#FFCC00"> 
                  <td width="49%">����</td>
                  <td colspan="2">����</td>
                </tr>
                <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left"><%=rs("p_�е�һ��")%>(��1)</td>
                  <td width="27%" align="right"><%=rs("p_�е�һƱ��")%></td>
                  <td width="24%"><%if rs("p_�е�һ�콱")=false then%><a href="lingjiang.asp?act=1">�콱</a><%else%>����<%end if%></td>
                </tr>
                 <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left"><%=rs("p_�еڶ���")%>(��2)</td>
                  <td width="27%" align="right"><%=rs("p_�еڶ�Ʊ��")%></td>
                  <td width="24%"><%if rs("p_�еڶ��콱")=false then%><a href="lingjiang.asp?act=2">�콱</a><%else%>����<%end if%></td>
                </tr>
                <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left"><%=rs("p_Ů��һ��")%>(Ů1)</td>
                  <td width="27%" align="right"><%=rs("p_Ů��һƱ��")%></td>
                  <td width="24%"><%if rs("p_Ů��һ�콱")=false then%><a href="lingjiang.asp?act=3">�콱</a><%else%>����<%end if%></td>
                </tr>
                 <tr align="center" bgcolor="#FFCC00"> 
                  <td align="left"><%=rs("p_Ů�ڶ���")%>(Ů2)</td>
                  <td width="27%" align="right"><%=rs("p_Ů�ڶ�Ʊ��")%></td>
                  <td width="24%"><%if rs("p_Ů�ڶ��콱")=false then%><a href="lingjiang.asp?act=4">�콱</a><%else%>����<%end if%></td>
                </tr>
              </table>
              <font color="#0000FF"><strong> ���������ֵ��������ȡ�\��<br>
              ����������<font color=red><%=rs("p_��������")%></font>��<br>
              ����ۼƣ�<font color=red><%=rs("p_���")%></font>��<br>
              Ԥ�ƽ���<font color=red><%=(int(rs("p_���")*0.6))%></font>��<br>
              ͶƱ������<font color=red><%=rs("p_ͶƱ����")%></font>�� 
              <%rs.close
              set rs=nothing
              conn.close
              set onn=nothing
              end if%>
              </strong></font></td>
          </tr>
          <tr> 
            <td align="center"><font color="#0000FF">��������</font></td>
          </tr>
          <tr> 
            <td><strong>1.����������</strong><br>
              ÿ�β�������ȡ������<font color="#FF0000"><strong>10�����</strong></font>��ÿ��ֻ�ޱ���һ�Σ�������Ҫ���ؿ��ǲμӱ�����һ��������ı����ν��������ٸ��ġ� 
              <br> <br> <strong>2.��ѡ�취��</strong><br>
              �����ȼ���<%=Application("aqjh_newplay")%>�����ϵ���ҿ��Բμ���ѡ�񣬲μ�ͶƱ�����ÿ��ϵͳ����<strong><font color="#FF0000">1������</font></strong>��Ϊ������<br>
              ÿ��ÿ�ο���ͶƱһ�Σ�������Ʊ���������һ�����ɾ��ID���֡�<br> <font color="#0000FF">�����ͶŮ�����е�Ʊ</font>��<font color="#FF0000">Ů���Ͷ����ҵ�Ʊ</font>!<br> 
              <br> <strong>3.�����취��</strong><br>
              ÿ�����ǽ���ѡ�С�Ů��2�����н����������Ľ�<br> <font color="#0000FF">��һ��Ϊ�ܱ������õ�%20</font><br> 
              <font color="#0000FF">�ڶ���Ϊ�ܱ������õ�%10</font><br>
              20x2+10x2=60% ϵͳ��Ϊ�ܱ������õ�60%!����˭���Ǳ��ε����˶���<br> <strong><br>
              4.ʱ�䰲�ţ�</strong> <br> <font color="#0000FF">ÿ��1-20��Ϊ����ʱ��</font>��<font color="#FF0000">21-27��Ϊ��ѡʱ��</font>,<font color="#0000FF">28�ŵ��µ�Ϊ�콱Ʒʱ��</font>��<strong>���콱Ʒʱ��û���콱Ʒ���û����ᰴ��Ȩ����!<br>
              <br>
              5.ע�����<br>
              </strong>a.��ͬ���ΰ������Ⱥ����Ρ�<br>
              b.��Ϊ��ͶƱ����,��ʱ��������ͬ��װ����ͬƱ������һ��������⡣<br>
              c.�������ף���Ʊ��Ϊһ����������ȡ�������ʸ񣬶���������ɾ��ID���ݡ�</td>
          </tr>
          <%if ndate>=21 then%>
          <tr>
            <td><marquee scrollamount="1" scrolldelay="1" direction= "up"  height="350">
            <div align=center><%=baoren%></div>
            </marquee>
            </td>
          </tr>
          <%end if%>
        </table></td>
      <td width="74%" valign="top"> 
        <%
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("showmdb")
rs.Open "select * from sai  Order by s_����ʱ��", conn, 1,1
if rs.eof and rs.bof then%>
        <div align="center">�Բ�����ʱû�з����κμ�¼���� <br>[ <a href=# onclick=history.go(-1)>������һ��</a> ]<br></div>
<%else
RS.PageSize=20'����ÿҳ��¼��,���޸�
Dim TotalPages
TotalPages = RS.PageCount
If CurPage>RS.Pagecount Then
CurPage=RS.Pagecount
end if
RS.AbsolutePage=CurPage
rs.CacheSize = RS.PageSize'��������¼��
Dim Totalcount
Totalcount =INT(RS.recordcount)
StartPageNum=1
do while StartPageNum+10<=CurPage
StartPageNum=StartPageNum+10
Loop
EndPageNum=StartPageNum+9
If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount%>
<table border="1" width="100%" cellpadding="1" cellspacing="0" bordercolordark="#FFFFFF" bordercolorlight="#00215a" align="center">
<tr>
            <td colspan=6 align=middle bordercolorlight="#00215a" bgcolor="#d7ebff">���������ѡ������(<font color="#0000FF"> 
              <%ndate=Day(date())
if ndate<=20 then%>
              ������������� 
              <%elseif ndate<=27 then%>
              ������������� 
              <%else%>
              ����Ʊ���� 
              <%end if%>
              ���в�����Ů</font>)</td>
          </tr>
<%I=0
p=RS.PageSize*(Curpage-1)
Response.Write"<tr>"
do while (Not RS.Eof) and (I<RS.PageSize)
	p=p+1
	myrnd=int(rnd*999)+1000
	Response.Write"<td   width=""25%"" height=""226"" align=center valign=""top"">"
	Response.Write"<div id=show"& myrnd &" style='PADDING-RIGHT: 0px; PADDING-LEFT: 0px; LEFT: 0; PADDING-BOTTOM: 0px; WIDTH: 140px; PADDING-TOP: 0px; POSITION: relative; TOP: 0; HEIGHT: 226px;'></div>"
	Response.Write "<script Language=Javascript>var jhshow"& rs("s_id") &"='"& rs("s_����") &"';"
	Response.Write "var showArray = jhshow"& rs("s_id") &".split('|');"
	Response.Write "var s='';"
	Response.Write "for (var i=0; i<25; i++){if(showArray[i] != ''){j = i + 1;"
	Response.Write "s+='<IMG id=s'+j+' "& showgif1 &" style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:'+j+';"& chr(34) &">';"
	Response.Write "}}"
	Response.Write "s+='<IMG src=face/blank.gif style="& chr(34) &"padding:0;position:absolute;top:0;left:0;width:140;height:226;z-index:50;"& chr(34) &" title=""ID:"& rs("s_id") &" ��Ʊ��"& rs("s_Ʊ��") &"&#13&#10˵����"& rs("s_˵��") &""">';"
	Response.Write "show"& myrnd &".innerHTML=s;</script>"
	if ndate<=27 and ndate>20 then
		Response.Write "<a href=""toupiao.asp?id="& rs("s_id") &""" ><font color=red><b>ͶƱ</b></font></a> "
		Response.Write "��Ʊ:"& rs("s_Ʊ��") &"<br>"
	end if
	Response.Write "<font color=blue>"& rs("s_˵��") &"</font>"
	Response.Write"</td>"
	I=I+1
	if i/4=int(i/4) then Response.Write"</tr><tr>"
	RS.MoveNext
Loop
if i/4<>int(i/4) then Response.Write"</tr>"
%>
<tr>
            <td colspan=6 align=middle bordercolorlight="#00215a" bgcolor="#d7ebff"> 
              ��<font color=blue><%=Totalcount%></font>�� ��<%=int(Totalcount/25)+1%>ҳ 
              [ 
              <% if StartPageNum>1 then %>
              <a href="jl.asp?CurPage=<%=StartPageNum-1%>&fs=<%=fs%>"><<</a> 
              <%end if%>
              <% For I=StartPageNum to EndPageNum
if I<>CurPage then %>
              <a href="jl.asp?CurPage=<%=I%>&fs=<%=fs%>"><%=I%></a> 
              <%else %>
              <b><font color=#ff0000><%=I%></font></b> 
              <%end if %>
              <%Next %>
              <%if EndPageNum<RS.Pagecount then %>
              <a href="jl.asp?CurPage=<%=EndPageNum+1%>&fs=<%=fs%>">>></a> 
              <%end if%>
              ] 
              <%
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
            </td>
          </tr></table>
      </td>
    </tr></table>
</div></body></html>