<?xml version="1.0" encoding="UTF-8"?>
		<urlset xmlns="http://www.google.com/schemas/sitemap/0.84">
		<!--#include file="Inc/conn.asp"-->
			<url>
					<loc>http://<%=Siteurl%><%=SitePath%></loc>
					<lastmod>2008-09-26</lastmod>
					<changefreq>always</changefreq>
					<priority>0.9</priority>
			</url>
			<url>
					<loc>http://<%=Siteurl%><%=SitePath%>Guestbook.asp</loc>
					<lastmod>2008-09-26</lastmod>
					<changefreq>daily</changefreq>
					<priority>0.5</priority>
			</url>
			<url>
					<loc>http://<%=Siteurl%><%=SitePath%>Search.asp</loc>
					<lastmod>2008-09-26</lastmod>
					<changefreq>daily</changefreq>
					<priority>0.5</priority>
			</url>
			<url>
					<loc>http://<%=Siteurl%><%=SitePath%>photo/</loc>
					<lastmod>2008-09-26</lastmod>
					<changefreq>daily</changefreq>
					<priority>0.3</priority>
			</url>
				<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select * from pcook_Class Where IsMenu=1 order by num asc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err)

set rsClass=server.createobject("adodb.recordset")
sql = "select * from pcook_Class where ID="&rs1("ClassID")&""
rsClass.open sql,conn,1,1  
ClassName=rsClass("ClassName")
rsClass.close
set rsClass=nothing
%>
			<url>
					<loc>http://<%=Siteurl%><%=SitePath%>class.asp?id=<%=rs1("ID")%></loc>
					<lastmod>2008-09-26</lastmod>
					<changefreq>daily</changefreq>
					<priority>0.5</priority>
			</url>
			<%
  rs1.movenext
  loop
  end if
  rs1.close
  set rs1=nothing
%>

		<%
set rs1=server.createobject("ADODB.Recordset")
sql1="select ID,Title,Content,ClassID,TitleFontColor,DateAndTime from pcook_Article where yn = 0 order by id desc"
rs1.open sql1,conn,1,3
If Not rs1.Eof Then 
do while not (rs1.eof or err)

set rsClass=server.createobject("adodb.recordset")
sql = "select * from pcook_Class where ID="&rs1("ClassID")&""
rsClass.open sql,conn,1,1  
ClassName=rsClass("ClassName")
rsClass.close
set rsClass=nothing
%>

			<url>
					<loc>http://<%=Siteurl%><%=SitePath%><%If seo=1 then%>Article_<%=rs1("ID")%>.html<%elseif seo=2 then%>Article/?<%=rs1("ID")%>.html<%elseif seo=3 then%>list.asp?id=<%=rs1("ID")%><%elseif seo=4 then%>Article/<%=rs1("ID")%>.html<% end if %></loc>
					<lastmod><%=year(rs1("DateAndTime"))%>-<%=month(rs1("DateAndTime"))%>-<%=day(rs1("DateAndTime"))%></lastmod>
					<changefreq>daily</changefreq>
					<priority>0.3</priority>
			</url>
			
<%
  rs1.movenext
  loop
  end if
  rs1.close
  set rs1=nothing
%>
   </urlset>
