<!--#include file="Inc/conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Pcook CMS �ݿ� V2.1 to V2.2��������</title>
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
    <td height="300"><p>����˵��:</p>
    <p>�����������������Pcook CMS �ݿ�V2.1������V2.2,����֮ǰ�����б���ԭ���ݿ�,������ɲ�����ص���ʧ.��վ���Դ�������������ĺ������</p>
    <form id="form1" name="form1" method="post" action="?setp=2">
      <input type="submit" name="button" id="button" value="���ѱ��������ݿ�,��˿�ʼ����" />
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
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('����Ա',1,1,'31.gif',0,1,-1,'#3F6C3F')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('ע���û�',1,6,'01.gif',0,1,-1,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('ע���û�',1,6,'01.gif',0,1,-1,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('�����',0,0,'14.gif',3,0,10000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('С����',0,0,'13.gif',3,0,8000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('����',0,0,'12.gif',3,0,6000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('����',0,0,'11.gif',3,0,4000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('�ƹ�',0,0,'10.gif',3,0,2800,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('��ũ',0,0,'09.gif',3,0,2000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('��ũ',0,0,'08.gif',3,0,1200,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('����',0,0,'07.gif',3,0,800,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('���',0,0,'06.gif',3,0,400,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('ƶũ',0,0,'05.gif',3,0,200,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('�軧',0,0,'04.gif',3,0,100,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('����',0,0,'03.gif',3,0,30,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('�̹�',0,0,'02.gif',3,0,10,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('����',0,0,'01.gif',3,0,0,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('С����',0,0,'15.gif',3,0,12000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('�����',0,0,'16.gif',3,0,15000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('֪��',0,0,'17.gif',3,0,20000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('ͨ��',0,0,'18.gif',3,0,30000,'')"
conn.execute="Insert Into pcook_UserGroup ([GroupName],[IsDisp],[Orders],[GroupPic],[ParentGID],[IsSetting],[Usermoney],[color]) Values ('����',0,0,'19.gif',3,0,50000,'')"
End if
Conn.Execute("Delete from [pcook_UserGroup] where UserGroupID = 2")

Response.Write("��ϲ��,�����ɹ�!��ɾ�����ļ�!<br><br>")
Response.Write("<font style='font:bold 16px Microsoft Yahei,sans-serif;line-height:50px;height:50px;'>����ϸ�������������,���ʣ�µ���������.</font><br><br>")
Response.Write("�ӹٷ������°�2.2���򣬽�ԭ������UploadFiles(ͼƬ�ϴ�Ŀ¼)�Ͳɼ����ݿ�(Ĭ��λ��Admin/cai/Database��)���ǵ��°�2.2�����У��ϴ�����ԭ��վ���ɡ�<br><br>")
Response.Write("<font color=red>�����б���ԭ��վ����!</font>")
end if 
%>
</body>
</html>