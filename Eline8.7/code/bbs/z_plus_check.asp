<!--#include file=z_plus_conn.asp-->
<!--#include file=z_plus_connbbs.asp-->
<%
'�����������
Dim plus_sql,plus_rs,plus_setting,plusname,pgetnum
dim plus_ftime,plus_stime,plusScript,plusid
dim plus_wz,plus_ml,plus_jy,plus_ww,plus_jq
plusScript=lcase(request.ServerVariables("PATH_INFO"))
plus_sql="select * from plus where '/' & trim(lcase(url))=right('"&trim(lcase(plusScript))&"',len(trim(url))+1) or trim(lcase(url))='"&trim(lcase(plusScript))&"'"
set plus_rs= server.createobject ("adodb.recordset")
plus_rs.open plus_sql,connplus,1,3
if plus_rs.eof and plus_rs.bof then
	plus_rs.close
	set plus_rs=nothing
	response.write "����Ĳ�������������Ч���ӽ���"
	response.end
else
	plusid=plus_rs("plusid")
	plus_setting=split(plus_rs("config"),",")
	plusname=plus_rs("plusname")
	pgetnum=plus_rs("getnum")
	if isnull(pgetnum) or pgetnum="" then pgetnum=0
	plus_rs("getnum")=pgetnum+1
	plus_rs.update
end if
plus_rs.close
set plus_rs=nothing
set plus_sql=nothing

if plus_setting(0)="1" then
	response.write "�ǳ���Ǹ��"&plusname&" �����ͣ����ʹ�ã����Ժ����ԣ�"
	response.end
end if
if plus_setting(1)="1" then
   plus_ftime=split(plus_setting(2),"|")
   plus_stime=split(plus_setting(3),"|")
   if (Hour(Now)>=Cint(plus_ftime(0)) and Hour(Now)<Cint(plus_ftime(1))) or (Hour(Now)>=Cint(plus_stime(0)) and Hour(Now)<Cint(plus_stime(1))) then
       response.write "�ǳ���Ǹ��"&plusname&" ������ʱ����ڲ���ʹ��<br>"
       response.write "<B>"&plus_fTime(0)&"</B>��<B>"&plus_fTime(1)&"</B>��,<B>"&plus_stime(0)&"</B>��<B>"&plus_stime(1)&"</B>��<br>"
       response.write "���ڹ涨ʱ���ڷ��ʣ�лл��"
       response.end
   end if
end if

if plus_setting(6)="1" or plus_setting(12)="0" or plus_setting(13)="1" then
   bbs_sql="select article,usercp,userep,userpower,userwealth,lockuser from [User] where username='"&membername&"'"
   set bbs_rs=bbs_conn.execute(bbs_sql)
   if bbs_rs.eof and bbs_rs.bof then
       response.write "�ǳ���Ǹ����û��ʹ��"&plusname&"��Ȩ�ޣ���������cookies��ʧЧ"
       response.end
   end if
   if plus_setting(12)="0" then
       if bbs_rs("lockuser")<>0 then
          response.write "�ǳ���Ǹ����û��ʹ��"&plusname&" ��Ȩ��"
          response.end
       end if
   end if

   if plus_setting(6)="1" then
      plus_wz=split(plus_setting(7),"|")
      plus_ml=split(plus_setting(8),"|")
      plus_jy=split(plus_setting(9),"|")
      plus_ww=split(plus_setting(10),"|")
      plus_jq=split(plus_setting(11),"|")
      
      if Clng(plus_wz(0))>=0 and Clng(bbs_rs("article"))<Clng(plus_wz(0)) then
          response.write "�ǳ���Ǹ����ʹ��"&plusname&" ���������������������Ҫ��<br>"
          response.end
      end if
      if Clng(plus_wz(1))<>-1 and Clng(bbs_rs("article"))>Clng(plus_wz(1)) then
          response.write "�ǳ���Ǹ����ʹ��"&plusname&" ���������������������Ҫ��<br>"
          response.end
      end if
      if Clng(plus_ml(0))>=0 and Clng(bbs_rs("usercp"))<Clng(plus_ml(0)) then
          response.write "�ǳ���Ǹ����ʹ��"&plusname&" �����������ֵ������Ҫ��<br>"
          response.end
      end if
      if Clng(plus_ml(1))<>-1 and Clng(bbs_rs("usercp"))>Clng(plus_ml(1)) then
          response.write "�ǳ���Ǹ����ʹ��"&plusname&" �����������ֵ������Ҫ��<br>"
          response.end
      end if
      if Clng(plus_jy(0))>=0 and Clng(bbs_rs("userep"))<Clng(plus_jy(0)) then
          response.write "�ǳ���Ǹ����ʹ��"&plusname&" ������ľ���ֵ������Ҫ��<br>"
          response.end
      end if
      if Clng(plus_jy(1))<>-1 and Clng(bbs_rs("userep"))>Clng(plus_jy(1)) then
          response.write "�ǳ���Ǹ����ʹ��"&plusname&" ������ľ���ֵ������Ҫ��<br>"
          response.end
      end if
      if Clng(plus_ww(0))>=0 and Clng(bbs_rs("userpower"))<Clng(plus_ww(0)) then
          response.write "�ǳ���Ǹ����ʹ��"&plusname&" �����������ֵ������Ҫ��<br>"
          response.end
      end if
      if Clng(plus_ww(1))<>-1 and Clng(bbs_rs("userpower"))>Clng(plus_ww(1)) then
          response.write "�ǳ���Ǹ����ʹ��"&plusname&" �����������ֵ������Ҫ��<br>"
          response.end
      end if
      if Clng(plus_jq(0))>=0 and Clng(bbs_rs("userwealth"))<Clng(plus_jq(0)) then
          response.write "�ǳ���Ǹ����ʹ��"&plusname&" ������ĲƸ�ֵ������Ҫ��<br>"
          response.end
      end if
      if Clng(plus_jq(1))<>-1 and Clng(bbs_rs("userwealth"))>Clng(plus_jq(1)) then
          response.write "�ǳ���Ǹ����ʹ��"&plusname&" ������ĲƸ�ֵ������Ҫ��<br>"
          response.end
      end if
    end if
    bbs_rs.close
    set bbs_rs=nothing
    if plus_setting(13)="1" then
      set bbs_rs= server.createobject ("adodb.recordset")
      bbs_rs.open bbs_sql,bbs_conn,1,3
      bbs_rs("article")=bbs_rs("article")+clng(plus_setting(14))
      bbs_rs("usercp")=bbs_rs("usercp")+clng(plus_setting(15))
      bbs_rs("userep")=bbs_rs("userep")+clng(plus_setting(16))
      bbs_rs("userpower")=bbs_rs("userpower")+clng(plus_setting(17))
      bbs_rs("userwealth")=bbs_rs("userwealth")+clng(plus_setting(18))
      bbs_rs.update
      bbs_rs.close
      set bbs_rs=nothing
    end if
   set bbs_sql=nothing
end if
%>
