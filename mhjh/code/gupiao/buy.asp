<!--#include file="conn.asp"-->
<% 
if session("yx8_mhjh_username")="" then'1
  call endinfo("���ܽ��룬�㻹û�е�¼")
end if

sid=cstr(abs(Request.Form ("sid")))
ushare=abs(Request.Form ("ushare"))
if ushare>10000000 then
  call endinfo("��һ����������һǧ��֧��Ʊ��")
  response.end
end if

sql= "select * from ��Ʊ where sid="&sid        
set rs=conn.execute(sql) 
response.write rs("��ǰ�۸�")
response.end

if rs("��ǰ�۸�")<=1 OR (rs("��ǰ�۸�")-rs("���̼۸�"))/rs("���̼۸�")>=0.15 then'2
  call endinfo("����ͣ�ơ���ͣ����ߵ�ͣ��Ĺ�Ʊ�ǲ��ܽ����������Ŷ���� ")
else
  yx8_mhjh_username=session("yx8_mhjh_username")
  sql="select * from �ͻ� where �ʺ�='"&session("yx8_mhjh_username")&"'"
  set rs=conn.execute(sql)
  username=rs("�ʺ�")
  nowmoney=rs("�ʽ�")

  set rs_s=conn.execute("select ��ͨ��Ʊ,���̼۸�,��ǰ�۸�,��ҵ from ��Ʊ where sid="&sid)
  set rs_u1=conn.execute("select �ֹ���,�ʺ�,����۸� from �� where sid="&sid)
  nowp=rs_s("��ǰ�۸�")
  dsshare=rs_s("��ͨ��Ʊ")-ushare
  tot=rs_s("��ͨ��Ʊ")

  if dsshare<0 then
    call endinfo("û���㹻�Ĺ�Ʊ������")
    rs_s.close
    rs.close
    response.end
  elseif nowmoney<ushare*rs_s("��ǰ�۸�")*1.01 then
    call endinfo("�����ʽ��㣡")
    rs_s.close
    rs.close
    response.end
  end if

  sql = "select count(*) as num from �¼� "
  set rs=conn.execute(sql)
  num=rs("num")
  if num>=40 then '*
    Set rs = Server.CreateObject("ADODB.recordset")
    sql = "select top 1 * from �¼� order by ԭ��ʱ�� asc"
    rs.open sql,conn,3,2
    rs.delete
  end if '*

  if nowmoney>=ushare*rs_s("��ǰ�۸�")*1.01 and dsshare>=0 then'3
    delmoney=ushare*rs_s("��ǰ�۸�")
    set rs_u=conn.execute ("select �ֹ���,����۸� from �� where sid="&sid&" and �ʺ�='"&username&"'")
    if rs_u.eof then'4
      Randomize
      s=Rnd
      if s<=0.4 then
        s=s+0.5
      end if
      sprice=nowp*(1+ushare*s/(tot*5))
      sql="update ��Ʊ set ��ǰ�۸�="&sprice&", ��ͨ��Ʊ="&dsshare&",������=������+"&ushare&" where sid="&sid
      conn.execute sql
      sql="insert into �� (�ʺ�,sid,����۸�,ƽ���۸�,�ֹ���,��ҵ,ʱ��) values ('"&username&"','"&sid&"','"&nowp&"','"&nowp&"','"&ushare&"','"&rs_s("��ҵ")&"','"&date()&"')"
      conn.execute sql
      sql="update �ͻ� set �ʽ�=�ʽ�-"&delmoney&"*1.01 where  �ʺ�='"&username&"'"
      conn.execute sql
      sql="update ���ʽ� set ���ʽ�=���ʽ�+"&delmoney&"*0.99 where  �ʺ�='"&username&"'"
      conn.execute sql  
      zd=ccur((formatcurrency(sprice,3)-formatcurrency(nowp,3))/formatcurrency(rs_s("��ǰ�۸�"),3))
      mess="<font color=#990000>"&username&"����"&rs_s("��ҵ")&" "&ushare&" �ɣ��ּ����� "&formatpercent(zd,2,-1)&"</font>"
      sql="insert into �¼�(ԭ��,ԭ��ʱ��) values('"&mess&"','"&now()&"' )"
      conn.execute sql  
    else
      yuan=rs_u("�ֹ���")
      Randomize
      s=Rnd

      inprice=rs_u1("����۸�")
      if s<=0.4 then
        s=s+0.5
      end if
      sprice=nowp*(1+ushare*s/(tot*5))
      uprice=(rs_s("��ǰ�۸�")*ushare+rs_u("����۸�")*yuan)/(ushare+yuan)
      sql="update ��Ʊ set ��ǰ�۸�="&sprice&", ��ͨ��Ʊ="&dsshare&",������=������+"&ushare&" where sid="&sid
      conn.execute sql
      sql="update �� set ƽ���۸�="&uprice&",����۸�="&nowp&",�ֹ���="&rs_u("�ֹ���")&"+"&ushare&" where �ʺ�='"&username&"' and sid="&sid
      conn.execute sql
      sql="update �ͻ� set �ʽ�=�ʽ�-"&delmoney&"*1.01 where  �ʺ�='"&username&"'"
      conn.execute sql

      sql="update ���ʽ� set ���ʽ�=���ʽ�+"&delmoney&"*0.99 where  �ʺ�='"&username&"'"
      conn.execute sql
      zd=ccur((formatcurrency(sprice,3)-formatcurrency(nowp,3))/formatcurrency(rs_s("��ǰ�۸�"),3))
      mess="<font color=#990000>"&username&"�ٴ�����"&rs_s("��ҵ")&" "&ushare&" �ɣ��ּ����� "&formatpercent(zd,2,-1)&" </font>"
      sql="insert into �¼�(ԭ��,ԭ��ʱ��) values('"&mess&"','"&now()&"' )"
      conn.execute sql  
    end if'4
  end if'3
dim newtalkarr(600) 
   talkarr=Application("yx8_mhjh_talkarr") 
		talkpoint=Application("yx8_mhjh_talkpoint") 
		j=1 
		for i=11 to 600 step 10 
			newtalkarr(j)=talkarr(i) 
			newtalkarr(j+1)=talkarr(i+1) 
			newtalkarr(j+2)=talkarr(i+2) 
			newtalkarr(j+3)=talkarr(i+3) 
			newtalkarr(j+4)=talkarr(i+4) 
			newtalkarr(j+5)=talkarr(i+5) 
			newtalkarr(j+6)=talkarr(i+6) 
			newtalkarr(j+7)=talkarr(i+7) 
			newtalkarr(j+8)=talkarr(i+8) 
			newtalkarr(j+9)=talkarr(i+9) 
			j=j+10 
		next 
		newtalkarr(591)=talkpoint+1 
		newtalkarr(592)=1 
		newtalkarr(593)=0 
		newtalkarr(594)=username 
		newtalkarr(595)="���" 
		newtalkarr(596)="" 
		newtalkarr(597)="000000" 
		newtalkarr(598)="000000" 
		newtalkarr(599)="<font color=red>����Ʊ�չ���</font><font color=blue>"&mess&"</font>" 
newtalkarr(600)=Session("yx8_mhjh_userchatroomsn")
Application.Lock
Application("yx8_mhjh_talkpoint")=talkpoint+1
Application("yx8_mhjh_talkarr")=newtalkarr
Application.UnLock
erase newtalkarr
erase talkarr
%>
<html>
<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="../style1.css">
<%
  call endinfo("��Ʊ����ɹ���")
end if

%>
<%conn.close
set conn=nothing
sub endinfo(message) 
'-------------------------------��Ϣ��ʾ-------------------------------
%>
<table width="<%=TableWidth%>" border=0 cellspacing=1 cellpadding=3 align=center bgcolor="<%=Tablebackcolor%>"><tr bgcolor="<%=Tabletitlecolor%>" ><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><b>��Ϣ��ʾ</b></td></tr><tr><td align=center height=70 bgcolor="<%=aTabletitlecolor%>"><%=message%><br></td></tr><tr><td align=center height=26 bgcolor="<%=Tabletitlecolor%>"><a href="stock.asp">����</a></td></tr></table>
<%end sub
%>

