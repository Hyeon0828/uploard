<html>
 <%
  Dim co_1, co_2, co_3, co_4, co_5, co_6, co_7, co_8, co_count, co_show

    co_1 = "#6E6E6E"
    co_2 = "#2E2EFE"
    co_3 = "#F5A9A9"
    co_4 = "#F3F781"
    co_5 = "#58FAF4"
    co_6 = "#FA58F4"
    co_7 = "#FACC2E"
    co_8 = "#9AFE2E"
    co_count = 0
    page_count = 0
  %>
 <style type="text/css">
  #div_tyme{
    display: flex;
  }
  .box_txt1{
    text-align: center;
    font-weight: bolder;
    border-right: 0px;
    font-size: 16px;
    width: 100px;
    color: black;
    border-radius: 10px 0px 0px 10px;
  }
  .box_txt2{
    border-left: 0px;
    border-right: 0px;
    font-size: 14px;
    padding: 0px;
  }
  .box_txt3{
    text-align: center;
    border-right: 0px;
    font-size: 12px;
    width: 80px;
    color: black;
    border-radius: 0px 5px 5px 0px;
  }
 </style>
 <body>
   <center>
     <table>
       <%
         Set db = Server.createObject ("ADODB.Connection")
         dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("bang.accdb")
         db.Open (dbMain)

         all_data_num = "select * from db_num"    '게시물 개수 세기
         Set dummy_db = db.execute(all_data_num)

         db_count = int(dummy_db("db_datanum"))

         if int(request("page")) = 0 then '처음 페이지와 쪽수 누른 페이지 구별
            db_page = 0
         else
            db_page = int(request("page")) - 1
         end if

         Dim cnt_id, sql
         sql = "select * from main_table ORDER BY db_date DESC"
         Set search_db = db.execute(sql)

         if search_db.EOF then
          response.Write "작성된 방명록이 없습니다."

         else
           db_count_mod = 0

           if db_count mod 6 = 0 then  '첫 페이지에 보일 레코드 개수 구하기 나머지가 0일경우 6으로 출력
              db_count_mod = db_count mod 6 + 6
           else
              db_count_mod = db_count mod 6
           end if

           if  db_count <> 6 * db_page + db_count_mod then '마지막 쪽 제외하고 레코드 값 넘기기
             for i = 1 to db_count - db_page * 6 - db_count_mod
             search_db.MoveNext
             next
           end if

           if db_page = 0 then '첫 페이지 나타낼 레코드 할당
              Do while Not search_db.EOF

                 in_a = search_db("db_name")
                 in_b = search_db("db_main")
                 in_c = search_db("db_date")
                 co_count = search_db("db_co")
                 if co_count = 1 then co_show = co_1
                 if co_count = 2 then co_show = co_2
                 if co_count = 3 then co_show = co_3
                 if co_count = 4 then co_show = co_4
                 if co_count = 5 then co_show = co_5
                 if co_count = 6 then co_show = co_6
                 if co_count = 7 then co_show = co_7
                 if co_count = 8 then co_show = co_8

                 response.write "<tr><td><div id=div_tyme><input type=text style='background:" & co_show & "; border: " & co_show &" 1.5px solid;' disabled class=box_txt1 value=" & in_a &_
                    "><textarea readonly rows=3 cols=55 style='border: " & co_show & " 1.5px solid;' class=box_txt2>" &_
                    in_b & " </textarea><input type=text disabled  style='background:" & co_show & "; border: " & co_show &" 1.5px solid;'  class=box_txt3 value=" & in_c & "</div></td></tr>"
                 search_db.MoveNext
               Loop
           else
             Do while Not page_count = 6
                 page_count = page_count + 1
                 in_a = search_db("db_name")
                 in_b = search_db("db_main")
                 in_c = search_db("db_date")
                 co_count = search_db("db_co")
                 if co_count = 1 then co_show = co_1
                 if co_count = 2 then co_show = co_2
                 if co_count = 3 then co_show = co_3
                 if co_count = 4 then co_show = co_4
                 if co_count = 5 then co_show = co_5
                 if co_count = 6 then co_show = co_6
                 if co_count = 7 then co_show = co_7
                 if co_count = 8 then co_show = co_8

                 response.write "<tr><td><div id=div_tyme><input type=text style='background:" & co_show & "; border: " & co_show &" 1.5px solid;' disabled class=box_txt1 value=" & in_a &_
                    "><textarea readonly rows=3 cols=55 style='border: " & co_show & " 1.5px solid;' class=box_txt2>" &_
                    in_b & " </textarea><input type=text disabled  style='background:" & co_show & "; border: " & co_show &" 1.5px solid;'  class=box_txt3 value=" & in_c & "</div></td></tr>"
               search_db.MoveNext
             Loop
           end if
         end if
         db.close()
         set db = nothing
         set search_db = nothing
        %>
    </table>
  </center>
  </body>
 </html>
