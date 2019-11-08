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
 response.Write "<script>alert('비밀번호가 일치하지 않습니다.');history.go(-1);</script>"
 db.close()

 else
  one_db = "delete from ex1table where db_id = '" & db_id & "'"
  db.execute(one_db)
  db.close()
  Session.Abandon
  response.write "<script language=javascript>alert('회원탈퇴가 정상 처리되었습니다.');opener.parent.location.reload();self.close();</script>"
 end if
 %>
  <body>
  </body>
 </html>
