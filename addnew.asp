<!--#include file="Inc/conn.asp"-->
<%
Dim Author,Content,sResult,ArticleID
Author =CheckStr(unescape(Request.Form("Author")))
Content =CheckStr(unescape(Request.Form("Content")))
ArticleID =CheckStr(unescape(Request.Form("ArticleID")))
IP	=Request.ServerVariables("REMOTE_ADDR")

if Instr(Content,"url=") or Instr(Content,"http://") or Instr(Content,"ad") or Instr(Content,"共产党") or Instr(Content,"你妈的") or Instr(Content,"他妈的") or Instr(Content,"成人") or Instr(Content,"操你") or Instr(Content,"做爱") or Instr(Content,"B毛") or Instr(Content,"法轮") or Instr(Content,"href") or Instr(Content,"卖身")>0 then
	Response.Write(escape("请不要发布违法及广告信息!"))
	Response.End
end if

sResult = Author + Content
Conn.Execute("INSERT INTO pcook_Pl(memAuthor,memContent,PostTime,ArticleID,IP,yn)  Values ('"&Author&"','"&Content&"','"&now()&"','"&ArticleID&"','"&IP&"',"&pingoff&")")
If Err Then 
   Response.Write(escape("出现错误!"))
Else
	If pingoff=0 then
		Response.Write(escape("发布评论成功,但是需要管理员审核后才能显示！"))
	else
		Response.Write(escape("发布评论成功!"))
	End if
End If
%>
