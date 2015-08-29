<%
Rem ==========论坛同学录函数=========
Rem 判断认证同学录进入用户
function chkschoollogin(boardid,username)
	dim boarduser
	chkschoollogin=false
	if master then
		chkschoollogin=true
	else
		sql="select boarduser,txlpd from board where boardid="&boardid
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1,1
		if rs.eof and rs.bof then
			chkschoollogin=false
		else
			if isnull(rs(0)) or rs(0)="" then
				chkschoollogin=false
			elseif cint(right(rs("txlpd"),1))=0 then
				chkschoollogin=false
			else
			boarduser=split(rs(0), ",")
			for i = 0 to ubound(boarduser)
				if trim(boarduser(i))=trim(username) then
					chkschoollogin=true
					exit for
				end if
			next
			end if
		end if
		rs.close
		set rs=nothing		
	end if
end function


Rem ==========中文长度=========
   function strLength(str)
       ON ERROR RESUME NEXT
       dim WINNT_CHINESE
       WINNT_CHINESE    = (len("同学录")=3)
       if WINNT_CHINESE then
          dim l,t,c
          dim i
          l=len(str)
          t=l
          for i=1 to l
             c=asc(mid(str,i,1))
             if c<0 then c=c+65536
             if c>255 then
                t=t+1
             end if
          next
          strLength=t
       else 
          strLength=len(str)
       end if
   end function

Rem ==========判断汉字=========
function isChinese(para)
    on error resume next
       if isNUll(para) then 
          isChinese=false
          exit function
       end if
       if trim(para)="" then
          isChinese=false
          exit function
       end if
       dim c
      for i=1 to len(para)
     c=asc(mid(para,i,1))
             if c>=0 then 
    isChinese=false
              exit function
           end if
       next
       isChinese=true
       if err.number<>0 then err.clear
end function
%>
