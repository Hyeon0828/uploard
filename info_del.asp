<html>
<%
db_pw = request("pw")
db_id = Session("user_id")

Dim dbMain
Set db = Server.createObject ("ADODB.Connection")
    dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("ex.accdb")
    db.Open (dbMain)


Dim one_db, sql
sql = "select * from ex1table where db_id ='" & db_id & "'" & " and db_pw ='" & db_pw & "'"
Set one_db = db.execute(sql)

if one_db.EOF then
 response.Write "<script>alert('��й�ȣ�� ��ġ���� �ʽ��ϴ�.');history.go(-1);</script>"
 db.close()

 else
  one_db = "delete from ex1table where db_id = '" & db_id & "'"
  db.execute(one_db)
  db.close()
  Session.Abandon
  response.write "<script language=javascript>alert('ȸ��Ż�� ���� ó���Ǿ����ϴ�.');opener.parent.location.reload();self.close();</script>"
 end if
 %>
  <body>
  </body>
 </html>
