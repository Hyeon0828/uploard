<html>
  <%
    db_id = request("db_id")

    Dim dbMain, db
    Set db = Server.createObject ("ADODB.Connection")
    dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("ex.accdb")
    db.Open (dbMain)

    Dim dummy_db, sql
    sql = "select * from ex1table where db_id ='" & db_id & "'"
    Set dummy_db = db.execute(sql)

    if dummy_db("db_master") = "m" then
     Response.Write  "<script>alert('삭제가 불가능한 아이디입니다.');history.go(-1);</script>"
     db.Close
    end if

    if dummy_db("db_master") = "" then
      sql = "select * from db_num"
      Set dummy_db = db.execute(sql)

      num_up = int(dummy_db("db_datanum")) - 1

      update_db = "update db_num set db_datanum = '" & num_up & "'"

      db.Execute(update_db)
      dummy_db = "delete from ex1table where db_id = '" & db_id & "'"
      db.Execute(dummy_db)
      db.close()
      Response.Write  "<script language=javascript>alert('아이디를 정상적으로 삭제하였습니다.');window.open('regit_info.asp','_self');</script>"
    End if

  %>
  <body>
  </body>
</html>
