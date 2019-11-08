<html>

  <%
    db_id = request("id")

    Dim dbMain, db
    Set db = Server.createObject ("ADODB.Connection")
    dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("db_sign.accdb")
    db.Open (dbMain)

    if session("ss_m") <> "m" then
      Response.Write  "<script>alert('권한이 없습니다.');history.go(-1);</script>"
      db.Close

    else
      update_db = "update db_num set db_datanum = 0"
      db.Execute(update_db)

      dummy_db = "delete from Db_table"
      db.Execute(dummy_db)

      db.close()
      Response.Write  "<script language=javascript>alert('초기화 성공.');window.open('a1.asp','_self');</script>"
    End if

  %>
  <body>
  </body>
</html>
