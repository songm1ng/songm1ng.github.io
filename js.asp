<!--#include file="Inc/conn.asp"-->
<%

dim path,classID,NclassID,NclassID1,showNclass,kind,dateNum,maxLen,listNum,bullet
dim hitColor,new_color,old_color

topType = Request("topType")

If Request("ClassNo") <> "" then
ClassNo = split(Request("ClassNo"),"|")
on error resume next
NClassID = ClassNo(0)
NClassID1 = ClassNo(1)
End if

num = request.querystring("num")
maxlen = Request.querystring("maxlen")
showdate = Request("showdate")
showhits = Request("showhits")
showClass = Request("showClass")

bullet="·"			 '标题前的图片或符号
hitColor="#FF0000"   '点击数的颜色
new_color="#FF0000"  '新文章日期的颜色
old_color="#999999"  '旧文章日期的颜色


dim rs,sql,str,topic
Path="http://"&request.servervariables("server_name")&replace(request.servervariables("script_name"),"js.asp","")

set rs=server.createObject("Adodb.recordset")
sql = "Select top "& num &" ID,Title,TitleFontColor,Author,ClassID,DateAndTime,Hits,IsTop,IsHot from pcook_Article Where yn = 0"

	If NclassID<>"" and NclassID1="" then
		If pcook_MyID(NclassID)="0" then
			SQL=SQL&" and ClassID="&NclassID&""
		else
			MyID = Replace(""&pcook_MyID(NclassID)&"","|",",")
			SQL=SQL&" and ClassID in ("&MyID&")"
		End if
	elseif NclassID<>"" and NclassID1<>"" then
		MyID = Replace(""&Request("ClassNo")&"","|",",")
		SQL=SQL&" and ClassID in ("&MyID&")"
	End if
	
select case topType
	case "new" sql=sql&" order by DateAndTime desc"
	case "hot" sql=sql&" order by hits desc,ID desc"
	case "IsHot" sql=sql&"and IsHot = 1 order by ID desc"
	'case "2" sql=sql&" DATEDIFF('d',intime,Now())<="&dateNum&" order by hits desc,Unid"
end select

set rs = conn.execute(sql)
if rs.bof and rs.eof then 
str=str+"没有符合条件的文章"
else

rs.movefirst
do while not rs.eof
	topic=Left(LoseHtml(rs("title")),maxlen)
	topic=replace(server.HTMLencode(topic)," ","&nbsp;")
	topic=replace(topic,"'","&nbsp;")
	str=str+bullet
	if showClass = 1 then
		str=str+"[<a href='"&path&"Class.asp?ID="&rs("ClassID")&"' target='_blank'>"&Classlist(rs("ClassID"))&"</a>] "
	end if
	str=str+"<a href='"&Path&""
	If seo=1 then
	str=str+"Article_"+Cstr(rs("ID"))+".html"
	elseif seo=2 then
	str=str+"Article/?"+Cstr(rs("ID"))+".html"
	elseif seo=3 then
	str=str+"list.asp?ID="+Cstr(rs("ID"))+""
	elseif seo=4 then
	str=str+"Article/"+Cstr(rs("ID"))+".html"
	End if
	str=str+"' target='_blank'  title='"&replace(replace(server.HTMLencode(rs("title"))," ","&nbsp;"),"'","&nbsp;")&"') >"+Topic+"</a>"
	if showdate <> 0 then
		str=str & "　<font color="
			if rs("DateAndTime")>=date then
				str=str & new_color
		 	else
				str=str & old_color
			end if
			str=str &">" & FormatDate(rs("DateAndTime"),""&showdate&"")&"</font>　"
	end if
	if showhits = 1 then
		str=str&"浏览：<font color="& hitcolor &">"& rs("hits") &"</font>"
	end if
	str=str & "<br />"
	rs.moveNext
loop
end if
rs.close : conn.close
set rs=nothing : set conn=nothing

response.write "document.write ("&Chr(34)&str&Chr(34)&");"
%>