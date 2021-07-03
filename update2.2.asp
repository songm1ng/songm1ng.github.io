<!--#include file="Inc/conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Pcook CMS 泡客 V2.1 to V2.2升级程序</title>
<style>
body,td {font-size:14px;line-height:20px;}
</style>
</head>
<body>
<%
setp=request("setp")
If setp=1 or setp="" then
%>
<table width="500" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="300"><p>升级说明:</p>
    <p>　　本升级程序针对Pcook CMS 泡客V2.1升级至V2.2,升级之前请自行备份原数据库,以免造成不可挽回的损失.本站不对此升级程序产生的后果负责。</p>
    <form id="form1" name="form1" method="post" action="?setp=2">
      <input type="submit" name="button" id="button" value="我已备份了数据库,点此开始升级" />
    </form></td>
  </tr>
</table>
<%
elseIf setp=2 then

Conn.Execute("Select top 1 UserID From pcook_Article")
If Err Then
conn.execute="alter table pcook_Article add column UserID Int default 0"
conn.execute="alter table pcook_Article add column UserName varchar(255)"
conn.execute("update pcook_Article set UserID = 0")
End if

Conn.Execute("Select top 1 * From [pcook_User]")
If Err Then
conn.execute="create table [pcook_User](ID autoincrement primary key,UserName varchar(50),[PassWord] varchar(50),sex Int default 0,Birthday DATETIME,Email varchar(50),UserFace varchar(250),province varchar(50),City varchar(50),UserQQ varchar(50),TrueName varchar(50),RegTime DATETIME,IP varchar(50),LastIP varchar(50),LastTime DATETIME,UserMoney Int default 0,dengji varchar(50),dengjipic varchar(50),usergroupid Int default 0,qm memo,yn Int default 0)" 
End if

Conn.Execute("Select top 1 * From [pcook_UserGroup]")
If Err Then
conn.execute="create table pcook_UserGroup(UserGroupID autoincrement primary key,GroupName varchar(50),IsDisp Int default 0,Orders Int default 0,GroupPic varchar(250),ParentGID Int default 0,IsSetting Int default 0,Usermoney Int default 0,color varchar(50))" 
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('管理员',1,1,'31.gif',0,1,-1,'#3F6C3F')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('注册用户',1,6,'01.gif',0,1,-1,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('注册用户',1,6,'01.gif',0,1,-1,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('大财主',0,0,'14.gif',3,0,10000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('小财主',0,0,'13.gif',3,0,8000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('衙役',0,0,'12.gif',3,0,6000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('商人',0,0,'11.gif',3,0,4000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('掌柜',0,0,'10.gif',3,0,2800,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('富农',0,0,'09.gif',3,0,2000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('中农',0,0,'08.gif',3,0,1200,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('猎人',0,0,'07.gif',3,0,800,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('渔夫',0,0,'06.gif',3,0,400,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('贫农',0,0,'05.gif',3,0,200,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('佃户',0,0,'04.gif',3,0,100,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('长工',0,0,'03.gif',3,0,30,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('短工',0,0,'02.gif',3,0,10,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('包身工',0,0,'01.gif',3,0,0,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('小地主',0,0,'15.gif',3,0,12000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('大地主',0,0,'16.gif',3,0,15000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('知县',0,0,'17.gif',3,0,20000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('通判',0,0,'18.gif',3,0,30000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('帝王',0,0,'19.gif',3,0,50000,'')"
End if
Conn.Execute("Delete from [pcook_UserGroup] where UserGroupID = 2")

Response.Write("恭喜你,升级成功!请删除本文件!<br><br>")
Response.Write("<font style='font:bold 16px Microsoft Yahei,sans-serif;line-height:50px;height:50px;'>请仔细看清下面的文字,完成剩下的升级工作.</font><br><br>")
Response.Write("从官方下载新版2.2程序，将原程序中UploadFiles(图片上传目录)和采集数据库(默认位于Admin/cai/Database下)覆盖到新版2.2程序中，上传覆盖原网站即可。<br><br>")
Response.Write("<font color=red>请自行备份原网站数据!</font>")
end if 
%>
</body>
</html>