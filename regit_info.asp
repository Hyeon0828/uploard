<html>
<style type="text/css">
    #div_tyme{
      float: right;
      display: flex;
      align:center;
    }
    #div_main{
      width: 650px;
      height: 400px;
      margin-top: 20px;
      align:left;
    }
   #box_title{
     border: 0px;
   }
   .box_font{
     font-size: 17px;
     margin-right: 10px;
   }
   .table {
     width: 100%;
     height: 450px;
   }
 </style>

 <script language = javascript>
   function sc_re() {
     window.open("regit_r.asp", '', 'status=no, height=500, width=700');
   }
   function sc_many() {
     window.open("regit_m.asp", '', 'status=no, height=500, width=700');
   }
 </script>

 <body onload="meeber.text_sc.focus();" align=center>
    <form name="member" action="regit_info.asp">

      <div id="div_main">
        <input type="button" value="�ֱ� �˻�" class="btn" onclick="sc_re();">
        <input type="button" value="�ִ� �˻�" class="btn" onclick="sc_many();">
      <div id="div_tyme">��������
        <select name="sel_sc" onchange="member.text_sc.focus();" class="form-control" style="width:100px; margin-right:3px;">
          <option value="1">����</option>
          <option value="2">���̵�</option>
        </select>
        <input type="text" name="text_sc" maxlength="16" class="form-control">
        <input type="submit" value="�˻�" class="btn btn-default" style="margin-left:3px; margin-bottom:5px;">
      </div>


      <table class="table">
        <thead style="background: #F6F6F6;">
          <tr>
            <th>���̵�</th>
            <th>����</th>
            <th>����</th>
            <th>�������</th>
          </tr>
        </thead>
        <tbody align=center>
       <%
         Dim cnt_id, sql
         Set db = Server.createObject ("ADODB.Connection")
         dbMain = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.Mappath("ex.accdb")
         db.Open (dbMain)

         text_sc = request("text_sc")
         sel_sc = request("sel_sc")
         page_count = 0

         For sp_i = 1 to len(text_sc)
          if mid(text_sc, sp_i, 1) <> chr(32) then
            text_sc_x = text_sc_x + mid(text_sc, sp_i, 1)
          end if
         next

         all_data_num = "select * from db_num"
         Set dummy_db = db.execute(all_data_num)

         db_count = int(dummy_db("db_datanum"))

         if int(request("page")) = 0 then 'ù �������� ���� ������ ����
            db_page = 0
         else
            db_page = int(request("page")) - 1
         end if

         if text_sc_x = "" then  '�� ���� �˻�

           sql = "select * from ex1table Order by ID ASC"
           Set search_db = db.execute(sql)

           if search_db.EOF then
            response.Write "��ϵ� ȸ�������� �����ϴ�."

           else '�� ���� �˻� �� ������ �ִ� ���
             db_count_mod = 0

             if db_count mod 10 = 0 then  'ó�� ȭ�鿡 ���ڵ�(�ʵ� ����) ���� ���ϱ� mod�� 0�ϰ�� 10���� ���
                db_count_mod = db_count mod 10 + 10
             else
                db_count_mod = db_count mod 10
             end if
             search_db.MoveFirst

             if  db_count <> 10 * db_page + db_count_mod then '������ �� �����ϰ� ���ڵ� �� �ѱ��
               for i = 1 to db_count - db_page * 10 - db_count_mod
               search_db.MoveNext
               next
             end if

             if db_page = 0 then 'ù ������ ��Ÿ�� ���ڵ� �Ҵ�
              Do while Not search_db.EOF

                db_id = search_db("db_id")
                db_name = search_db("db_name")
                db_gender = search_db("db_Gender")
                db_birth = int(search_db("db_birth"))

                response.write "<tr><td>" & db_id & "</td><td>" & db_name & "</td><td>" & db_gender & "</td><td>" & db_birth & "</td><td>"

                 search_db.MoveNext
               Loop

             else
               Do while Not page_count = 10

                  page_count = page_count + 1
                  db_id = search_db("db_id")
                  db_name = search_db("db_name")
                  db_gender = search_db("db_gender")
                  db_birth = int(search_db("db_birth"))

                  response.write "<tr><td>" & db_id & "</td><td>" & db_name & "</td><td>" & db_gender & "</td><td>" & db_birth & "</td><td>"

                  search_db.MoveNext
              Loop
             end if
           end if


          else
            if sel_sc = 1 then sql = "select * from ex1table where db_name like '%" & text_sc_x & "%'"
            if sel_sc = 2 then sql = "select * from ex1table where db_id like '%" & text_sc_x & "%'"

            Set search_db = db.execute(sql)

            if search_db.EOF then
             response.Write "<script>alert('ã�� ������ �����ϴ�.');window.open('regit_info.asp','_self');</script>"

            else

            db_count = 0
            db_count_mod = 0

            Do while Not search_db.EOF 'ã�� �Խñ� ����
              db_count = db_count + 1
              search_db.MoveNext
            Loop

            if db_count mod 10 = 0 then  'ù �������� ���� ���ڵ� ���� ���ϱ� �������� 0�ϰ�� 10���� ���
               db_count_mod = db_count mod 10 + 10
            else
               db_count_mod = db_count mod 10
            end if
            search_db.MoveFirst

            if  db_count <> 10 * db_page + db_count_mod then '������ �� �����ϰ� ���ڵ� �� �ѱ��
              for i = 1 to db_count - db_page * 10 - db_count_mod
              search_db.MoveNext
              next
            end if

            if db_page = 0 then 'ù ������ ��Ÿ�� ���ڵ� �Ҵ�
             Do while Not search_db.EOF

               db_id = search_db("db_id")
               db_name = search_db("db_name")
               db_gender = search_db("Db_Gender")
               db_birth = int(search_db("db_birth"))
               db_m = int(search_db("db_m"))

               if db_m = "" then db_m = 0

               now_date = now()    '�ð� ����� �������Ķ����� ������ ����� �ȵǴ� �� ����
               now_a = left(now_date,10)
               now_b = mid(now_date,12,2)
               space_c = mid(now_date,16,1) '�������� �������� �˾Ƴ��� ���� ����

               if space_c <> ":" then now_c = int(mid(now_date,15,2))
               if space_c = ":" then now_c = int(mid(now_date,15,1))

               now_d = right(now_date,6)

               if now_b = "����" then now_c = now_c + 12
               if now_c = 24 then now_c = 0

               now_cc = Cstr(now_c)  'str�� ��ȯ
               if now_c < 10 then now_cc = "0"+now_cc
               now_all = now_a + " " + now_cc + now_d

               db_m = db_m + 1

               update_m = "update ex1table set db_m = " & db_m & " where db_id ='" & db_id & "'"
               db.Execute(update_m)
               update_r = "update ex1table set db_r = '" & now_all & "' where db_id ='" & db_id & "'"
               db.Execute(update_r)


               response.write "<tr><td>" & db_id & "</td><td>" & db_name & "</td><td>" & db_gender & "</td><td>" & db_birth & "</td><td>"

                search_db.MoveNext
              Loop

            else
              Do while Not page_count = 10

                 page_count = page_count + 1
                 db_id = search_db("db_id")
                 db_name = search_db("db_name")
                 db_gender = search_db("db_gender")
                 db_birth = int(search_db("db_birth"))
                 db_m = int(search_db("db_m"))

                 if db_m = "" then db_m = 0

                 now_date = now()       '�ð� ����� �������Ķ����� ������ ����� �ȵǴ� �� ����
                 now_a = left(now_date,10)
                 now_b = mid(now_date,12,2)
                 space_c = mid(now_date,16,1)

                 if space_c <> ":" then now_c = int(mid(now_date,15,2))
                 if space_c = ":" then now_c = int(mid(now_date,15,1))

                 now_d = right(now_date,6)

                 if now_b = "����" then now_c = now_c + 12
                 if now_c = 24 then now_c = 0

                 now_cc = Cstr(now_c)
                 if now_c < 10 then now_cc = "0"+now_cc
                 now_all = now_a + " " + now_cc + now_d

                 db_m = db_m + 1

                 update_m = "update ex1table set db_m = '" & db_m & "' where db_id ='" & db_id & "'"
                 db.Execute(update_m)
                 update_r = "update ex1table set db_r = '" & now_all & "' where db_id ='" & db_id & "'"
                 db.Execute(update_r)


                 response.write "<tr><td>" & db_id & "</td><td>" & db_name & "</td><td>" & db_gender & "</td><td>" & db_birth & "</td><td>"

                 search_db.MoveNext
             Loop
            end if
            end if
          end if
        %>

          </tbody>
        </table>
          <%
            response.write "<font class=box_font><</font>"
            if db_count mod 10 = 0 then
              for i = 1 to db_count/10
                response.write "<a href='regit_info.asp?page=" & i & "&text_sc=" & text_sc_x & "&sel_sc=" & request("sel_sc") & "' target='_self' class=box_font>" & i & "</a>"
              next
            else
              for i = 1 to db_count/10 + 1
                response.write "<a href='regit_info.asp?page=" & i & "&text_sc=" & text_sc_x & "&sel_sc=" & request("sel_sc") & "' target='_self' class=box_font>" & i & "</a>"
              next
            end if
            response.write "<font class=box_font style='margin:0px;margin-left:2px;'>></font>"

            db.close()
            set db = nothing
            set search_db = nothing
          %>
      </div>
  </form>
  </body>
 </html>
