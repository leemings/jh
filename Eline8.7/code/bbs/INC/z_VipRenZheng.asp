<%function renzhen(boardid,username)
	dim boarduser,rrs,Board_Setting,BoardMaster
	renzhen=false
	if master then
		renzhen=true
	else
		sql="select boarduser,Board_Setting,BoardMaster from board where boardid="&boardid
		set rrs=server.createobject("adodb.recordset")
		rrs.open sql,conn,1,1
		Board_Setting=split(rrs("board_setting"),",")
		if cint(Board_Setting(2))=1 then
			if not (isnull(rrs(2)) or rrs(2)="") then
				BoardMaster=split(rrs(2), "|")
				for i = 0 to ubound(BoardMaster)
					if trim(BoardMaster(i))=trim(username) then
						renzhen=true
						exit for
					end if
				next
			end if
			if renzhen=false then
				if isnull(rrs(0)) or rrs(0)="" then
					renzhen=false
				else
					boarduser=split(rrs(0), ",")
					for i = 0 to ubound(boarduser)
						if trim(boarduser(i))=trim(username) then
							renzhen=true
							exit for
						end if
					next
				end if
			end if
		else
			renzhen=true
		end if
		rrs.close
		set rrs=nothing		
	end if
end function
function VipBoard(boardid,username)
	dim boarduser,rrs,Board_Setting,BoardMaster
	VipBoard=false
	if master then
		VipBoard=true
	else
		sql="select Board_Setting,BoardMaster from board where boardid="&boardid
		set rrs=server.createobject("adodb.recordset")
		rrs.open sql,conn,1,1
		Board_Setting=split(rrs("board_setting"),",")
		if cint(Board_Setting(51))=1 or cint(Board_Setting(52))=1 then
			if not (isnull(rrs(1)) or rrs(1)="") then
				BoardMaster=split(rrs(1), "|")
				for i = 0 to ubound(BoardMaster)
					if trim(BoardMaster(i))=trim(username) then
						VipBoard=true
						exit for
					end if
				next
			end if
			if VipBoard=false then
				if myVip<>1 or isnull(MyVip) then
					VipBoard=false
				else
					VipBoard=true
				end if
			end if
		else
			VipBoard=true
		end if
		rrs.close
		set rrs=nothing		
	end if
end function
%>