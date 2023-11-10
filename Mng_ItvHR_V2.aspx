<%@ Page Language="VB"  EnableEventValidation="true"  %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">


<!-- #include file="../../Public/aspxMasV6.aspx" -->

<script runat="server">
    '=============================================
    'Date :  2022/07/19  by Jassy
    'Version:1.0.0 Create
    '2022/07/27		修改外部, 職能. 欄位 有資料才顯示, 
    '2022/08/23     修改網頁名子.."面試管理暨查詢 (HR權限) "
    '2022/11/7      特殊字元的名字可以查詢
    '2022/12/23     top20 寫在最後的query  上
    '=============================================  
    Dim Dr As Data.DataRow
    Dim mas As New Mas, conn2 As Data.SqlClient.SqlConnection, conn3 As Data.SqlClient.SqlConnection

    Dim Funcs As New Funcs("SqlConnectNewEmpV2")
    Dim Hts As New Hashtable()

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)

        PageLoad("SqlConnectNewEmpV2")
        mas.PageLoad(IsPostBack, "SqlConnectNewEmpV2", 0)


        If Not IsPostBack Then
            txt_from.Text = Now().Date() ' Now().AddDays(0).Date()
            txt_to.Text = Now().Date()
            txt_from.Attributes.Add("onfocus", "dtSelect(this,'');")
            txt_to.Attributes.Add("onfocus", "dtSelect(this,'');")

            txt_edit_from.Text = Now().Date() ' Now().AddDays(0).Date()
            txt_edit_to.Text = Now().Date()

            txt_edit_from.Attributes.Add("onfocus", "dtSelect(this,'');")
            txt_edit_to.Attributes.Add("onfocus", "dtSelect(this,'');")

            Funcs.insList(lst_status, "SELECT  Code='',Descr='ALL' union all select 'HC','錄用簽核中' union all select Code,Descr from NewCode where Tp='Status' and Code in ('W','A','S')   ")
            Funcs.insList(lst_place, "SELECT   Code='%',Descr='ALL' union all SELECT   Code,Descr FROM NewCode WHERE Tp='Place' ORDER BY Code ")

            Funcs.insList(ddl_top, "SELECT   Code='top 20',Descr='前20筆' union all SELECT  'top 50','前50筆' UNION ALL SELECT  'top 100 ' , '前100筆' UNION ALL SELECT  'top 1000 ' , '前1000筆' ")

            Funcs.insList(ddl_site, "SELECT   Code='',Descr='ALL' union all SELECT   '桃園','桃園' UNION ALL SELECT  '士林' , '士林' ")
            Dim userSite As String

            Hts.Clear()
            Hts.Add("@uid", uid )
            userSite = Funcs.sqlFunc("select Txt FROM Base.[dbo].[EmpDet] where Tp='Company' AND  Eid=" & "@uid" ,Hts)

            ddl_site.SelectedValue = userSite
            lst_place.SelectedValue= "%"

            insert_lst_hr()
            rb1.Checked = True

            searchcertia()
            ' BindGrid()
            'ddl_site.SelectedValue = ""
        End If
        'searchcertia()
        ' BindGrid()


    End Sub



    Protected Sub insert_lst_hr()

        Dim site As String
        site = ddl_site.SelectedValue


        If site = "" Then
            Funcs.insList(lst_hr, "declare @hr varchar(max)select @hr='' " &
          "   Select @hr=@hr+HR from ItvMas   group by HR " &
          "   select  distinct Code=s, Descr=[dbo].[getEmpBase2](s,'Name')+'('+ [dbo].[getEmpBase2](s,'Ext')+')' from Base.dbo.Split(@hr,';') " &
          " union select Code='' , Descr='ALL'")
        Else


            Hts.Clear()
            Hts.Add("@ddl_siteSelectedValue", ddl_site.SelectedValue )

            Funcs.insList(lst_hr, "declare @hr varchar(max)select @hr='' " &
    "   Select @hr=@hr+HR from ItvMas   group by HR " &
    "   select  distinct Code=s, Descr=[dbo].[getEmpBase2](s,'Name')+'('+ [dbo].[getEmpBase2](s,'Ext')+')'" &
    "  From Base.dbo.Split(@hr,';') a  join Base.dbo.EmpMas b on a.s=b.Eid and b.ExitDt='' and Place=case '" &
    ddl_site.SelectedValue & "' when '士林' then 'S' else 'H'   end  union select Code='' , Descr='ALL'")

            lst_hr.SelectedValue = ""

        End If


    End Sub

    Function getmutiListvalue(obj As ListBox)
        Dim str As String = ""
        If obj.SelectedValue <> "" Then
            'ClientMsg(obj.Items.Count)

            For j As Integer = 0 To obj.Items.Count - 1

                If obj.SelectedValue = "%" Or obj.Items.Item(j).Selected = True Then
                    If str = "" Then
                        str = "'" + obj.Items.Item(j).Value + "'"

                    Else
                        str = str + ",'" + obj.Items.Item(j).Value + "'"

                    End If

                End If

            Next
        End If
        Return str
    End Function

    Protected Sub BindGrid()
        Dim sql As String = "", Msql As String = "", Msql_from As String = "", Msql_where As String = "", Endsql As String = ""  ,topstr As String = ""
        Msql = "declare @hr varchar(max)select @hr=''  Select @hr=@hr+HR from ItvMas   group by HR  "   ' 捉系統中所有hr

        topstr = ddl_top.SelectedValue


        Hts.Clear()
        Hts.Add("@topstr", topstr )


        Msql = Msql + " SELECT distinct " & topstr & " btn_visible='True' , Status=dbo.getItvBase2(a.Rid,'Status_CHI') ," &
"'Name'=c.Name+'<br/>'+c.Tel  ," &
"'DtTm'=a.ItvDt+'<br/>'+a.ItvTm,'EmpTp'=f.Descr, " &
"'FamilyPlace'=a.JobDescr+'<br/>'+h.Descr " &
",'Mng'= dbo.getItvBase2(a.Rid,'ItvMngGrade') " &
",'HR'=  dbo.getItvBase2(a.Rid,'ItvHRGrade')  " &
",Testing " &
",'Remark'=Base.dbo.GetXml(a.Remark,'HRLog')  + Case when dbo.GetXml(a.Remark,'SReasonTp')<>'' Then '&nbsp; 未到談分類：'+dbo.GetXml(a.Remark,'SReasonTp') Else '' End  + Case when dbo.GetXml(a.Remark,'SReason')<> '' Then '/未到談說明：' " &
"+dbo.GetXml(a.Remark,'SReason') Else '' End  + Case when dbo.GetXml(a.Remark,'ContactDt')<> '' Then '(下次聯絡日：' + dbo.GetXml(a.Remark,'ContactDt') + ')' Else '' End  + " &
"Case when dbo.GetXml(a.Remark,'DReasonTp')<> '' Then '&nbsp; 未錄取分類：'+ dbo.GetXml(a.Remark,'DReasonTp')  Else '' End  + Case when dbo.GetXml(a.Remark,'DReason')<> '' Then '/未錄取說明：' + dbo.GetXml(a.Remark,'DReason')  " &
"Else '' End  + Case when dbo.GetXml(a.Remark,'Enter')<> '' Then '<BR/>  錄取說明：' + dbo.GetXml(a.Remark,'Enter')  Else '' End  + Case when dbo.GetXml(a.Remark,'NDt')<> '' Then '<BR/>  報到取消日：' + dbo.GetXml(a.Remark,'NDt') " &
"        Else '' End  + Case when dbo.GetXml(a.Remark,'NReasonTp')<> '' Then '<BR/> 未報到分類：' + dbo.GetXml(a.Remark,'NReasonTp')  Else '' End  + Case when dbo.GetXml(a.Remark,'NReason')<> '' Then '/ ' + dbo.GetXml(a.Remark,'NReason')  " &
"        Else '' End  , " &
"        TalLevel=Isnull(j.Descr,'')+'<Br>'+Isnull(l.Descr,'')+' / '+Isnull(m.Descr,'')" &
"        ,'ItvId'=a.Rid" &
"        ,'TalentId'=c.Rid" &
"        ,TalentFid=Isnull(k.TalentId,0)" &
",Editor=dbo.getEmpBase2(a.Editor,'Mas1') " &
"  , a.ItvDt,a.ItvTm  "

        Msql_from = "FROM ItvMas a LEFT JOIN ItvDet b ON a.Rid=b.ItvId " &
"LEFT JOIN ApplyMas aa on " &
"/*aa.TalentId = a.TalentId and*/ aa.ItvId = a.Rid and aa.Status not in ('D') " &
"LEFT JOIN TalentMas c ON a.TalentId=c.Rid " &
"LEFT JOIN NewCode f ON f.Code=a.ItvTp And f.Tp='ItvTp' " &
"        LEFT JOIN NewCode g ON ISNULL(aa.Status,a.Status)=g.Code AND g.Tp='Status' " &
"LEFT JOIN NewCode h on a.Place=h.Code And h.Tp='Place'  " &
"        left join TalentDet i on c.Rid=i.TalentId " &
"Left Join Base..BaseCode j on i.JobFamily = j.Code And j.Tp='JobFamily' " &
"        Left Join NewCode l on i.JobLvl=l.Code And l.Tp='JobLvl' " &
"        Left Join NewCode m on i.TalentLvl=m.Code And m.Tp= 'TalentLvl' " &
"        Left JOIN dbo.AttachView k on c.Rid=k.Rid  "



        Msql_where = " WHERE      a.Status not in ('D') "

        If ck_date.Checked = True Then

            Hts.Add("@txt_fromText",  txt_from.Text )
            Hts.Add("@txt_toText",  txt_to.Text )


            Msql_where = Msql_where + " and a.ItvDt>=convert(datetime," & "@txt_fromText" & ")  AND a.ItvDt<=convert(datetime," & "@txt_toText" & ") "

        End If
        If ck_edit_date.Checked = True Then

            Hts.Add("@txt_edit_fromText",  txt_edit_from.Text )
            Hts.Add("@txt_edit_toText",  txt_edit_to.Text )

            Msql_where = Msql_where + " and a.Udt>=convert(datetime," & "@txt_edit_fromText" & ")  AND a.Udt<=dateadd(d,1,convert(datetime," & "@txt_edit_toText" & ") )"

        End If

        If ck_interviewer.Checked = True And Txt_interviewer.Text <> "" Then

            Hts.Add("@Txt_interviewerText",  Txt_interviewer.Text )


            Msql_where = Msql_where + " and c.Name  like N'%'+" & "@Txt_interviewerText" & "+'%' "
        End If

        If ck_boss.Checked = True And txt_boss.Text <> "" Then

            Hts.Add("@txt_bossText",  txt_boss.Text )

            Msql_where = Msql_where + " and dbo.getItvBase2(a.Rid,'ItvMngGrade') like '%'+" & "@txt_bossText" & "+'%'"
        End If

        If ck_place.Checked = True And getmutiListvalue(lst_place) <> "" Then


            Hts.Add("@lst_place",  getmutiListvalue(lst_place) )

            Msql_where = Msql_where + " and a.Place in (" & "@lst_place"& ")"
        End If


        If ck_status.Checked = True Then
            If Not (lst_status.SelectedValue = "") Then


                If lst_status.SelectedValue = "HC" Then
                    Msql_where = Msql_where + " and aa.Status IN ('H','C') "
                ElseIf lst_status.SelectedValue = "AS" Then
                    Msql_where = Msql_where + " and dbo.getItvBase2(a.Rid,'Status') IN ('A','S') "

                Else

                    Hts.Add("@lst_statusSelectedValue",  lst_status.SelectedValue )
                    Msql_where = Msql_where + " and charindex( dbo.getItvBase2(a.Rid,'Status')," & "@lst_statusSelectedValue" & " )>0 "
                End If
            End If
        End If
        If ck_hr.Checked = True Then
            If Not (lst_hr.SelectedValue = "") Then

                Hts.Add("@lst_hrSelectedValue",  lst_hr.SelectedValue )

                Msql_where = Msql_where + " and charindex( " & "@lst_hrSelectedValue" & ", a.HR  )>0 "
            End If
        End If



        If ck_site.Checked = True Then

            If Not (ddl_site.SelectedValue = "") Then


                Hts.Add("@ddl_siteSelectedValue",  ddl_site.SelectedValue )

                Msql_where = Msql_where + "  and ( left(a.HR,9) in (	 select s from Base.dbo.Split(@hr,';') a inner  join  Base.dbo.EmpMas b on a.s=b.Eid  and Place=case " &
                "@ddl_siteSelectedValue" & " when '士林' then 'S' else 'H'   end   ) " &
                  " or Right(a.HR,9) in (  select s from Base.dbo.Split(@hr,';') a inner  join  Base.dbo.EmpMas b on a.s=b.Eid and Place=case " &
                  "@ddl_siteSelectedValue" & " when '士林' then 'S' else 'H'   end   ))"

            End If

        End If



        Msql = Msql + Msql_from + Msql_where

        Msql = Msql + " ORDER BY ItvDt,ItvTm "

        'Response.Write(Msql)
        Funcs.gridBind( gd_view, Msql ,Hts)

    End Sub


    Public Sub disable_check()
        ck_edit_date.Checked = False
        ck_date.Checked = False
        ck_place.Checked = False
        ck_interviewer.Checked = False
        ck_hr.Checked = False
        ck_date.Checked = False
        ck_boss.Checked = False
        ck_status.Checked = False
        ck_site.Checked = False

    End Sub


    Public Sub searchcertia()
        If (rb1.Checked = True) Then
            disable_check()
            ck_date.Checked = True
            ck_hr.Checked = True
            ck_status.Checked = True
            lst_status.SelectedValue = ""
            ck_site.Checked = True

            BindGrid()

        ElseIf (rb2.Checked = True) Then
            disable_check()
            ck_status.Checked = True
            lst_status.SelectedValue = "HC"
            BindGrid()

        ElseIf (rb3.Checked = True) Then
            disable_check()
            ck_edit_date.Checked = True
            'ck_hr.Checked = True
            ' ck_status.Checked = True
            ' lst_status.SelectedValue = ""
            ck_site.Checked = True
            BindGrid()

        Else
            disable_check()
        End If



    End Sub


    Protected Sub btn_search_Click(sender As Object, e As EventArgs)
        If ck_boss.Checked = False And ck_date.Checked = False And ck_hr.Checked = False And ck_interviewer.Checked = False And ck_place.Checked = False And ck_status.Checked = False And ck_edit_date.Checked = False Then
            ClientMsg("至少需設定一個查詢條件")
        ElseIf ck_boss.Checked = False And ck_date.Checked = False And ck_hr.Checked = False And ck_interviewer.Checked = False And ck_place.Checked = False And (ck_status.Checked = True Or ck_hr.Checked = True) Then
            If (ck_status.Checked = True And lst_status.SelectedValue = "") Or (ck_hr.Checked = True And lst_hr.SelectedValue = "") Then
                ClientMsg("查詢條件不足")
            ElseIf lst_status.SelectedValue <> "" Or lst_hr.SelectedValue <> "" Then
                '   ClientMsg("wrong")
                BindGrid()
            End If
        Else
            BindGrid()
        End If

    End Sub
    Protected Sub btn_sel_Click(sender As Object, e As EventArgs)
    End Sub
    Protected Sub lbtn_Resume_Click(sender As Object, e As EventArgs)
        Dim btn As LinkButton = TryCast(sender, LinkButton)
        Dim gr As GridViewRow = btn.NamingContainer, ResumeItvId As String
        If gr IsNot Nothing Then
            Dim hf_TalentId As HiddenField = TryCast(gr.FindControl("hf_TalentId"), HiddenField)
            Dim hf_ItvId As HiddenField = TryCast(gr.FindControl("hf_ItvId"), HiddenField)

            If hf_TalentId.Value <> "" AndAlso hf_ItvId.Value <> "" Then


                Hts.Clear()
                Hts.Add("@hf_TalentIdValue", hf_TalentId.Value )

                ResumeItvId = Funcs.sqlFunc("select Isnull(max(b.ItvId),'') from NewEmpV2..ItvMas a join NewEmpV2..ItvEmp b on a.TalentId=b.TalentId And a.Rid=b.ItvId  And BirthY <> '' " _
               & " Where a.TalentId=" & "@hf_TalentIdValue" & "" ,Hts)
                'jassy sql injection			   
                mas.winOpen("rpt_Interview.aspx?val=" & ResumeItvId & " ")
                '"Select Isnull(max(a.ItvId) From ItvEmp a Join ItvMas b on a.TalentId=b.TalentId And a.ItvId=b.Rid Left Join ItvMas c on b.TalentId=c.TalentId And b.JobDescr	=c.JobDescr where a.ItvId = 39"
            End If
        End If
    End Sub






    Protected Sub rb1_CheckedChanged(sender As Object, e As EventArgs)
        searchcertia()

    End Sub

    Protected Sub ddl_site_SelectedIndexChanged(sender As Object, e As EventArgs)
        insert_lst_hr()
    End Sub

    Protected Sub gv_list_RowDataBound(sender As Object, e As GridViewRowEventArgs)
        If e.Row.RowType = DataControlRowType.DataRow Then


            'Dim ItvId As String = DataBinder.Eval(e.Row.DataItem, "ItvId").ToString
            Dim hf As HiddenField = TryCast(e.Row.FindControl("hf_ItvId"), HiddenField)
            Dim hf_tid As HiddenField = TryCast(e.Row.FindControl("hf_TalentId"), HiddenField)
            Dim lbtn_resume As LinkButton = TryCast(e.Row.FindControl("lbtn_Resume"), LinkButton)
            Dim lbtn_jfunc As LinkButton = TryCast(e.Row.FindControl("lbtn_JobFunc"), LinkButton)
            Dim lbtn_olink As LinkButton = TryCast(e.Row.FindControl("lbtn_Outlink"), LinkButton)
            Dim rsmFid As String = "", jFid As String = ""


            Hts.Clear()
            Hts.Add("@hf_tidValue", hf_tid.Value)

            rsmFid = Funcs.sqlFunc("select Isnull(max(b.ItvId),'') from NewEmpV2..ItvMas a join NewEmpV2..ItvEmp b on a.TalentId=b.TalentId And a.Rid=b.ItvId  And BirthY <> '' " _
               & " Where a.TalentId=" & "@hf_tidValue" & "" ,Hts)



            Hts.Clear()
            Hts.Add("@hfValue", hf.Value)

            jFid = Funcs.sqlFunc("declare @job nvarchar(100) = '' select @job = JobDescr from ItvMas where Rid=" & "@hfValue" & " select a.Rid FROM   Attachment.dbo.NewEmp AS a JOIN " _
            & " (SELECT a.TalentId, b.Tp, a.JobDescr, MAX(b.Cdt) AS Cdt FROM    dbo.ItvMas AS a INNER JOIN Attachment.dbo.NewEmp AS b ON a.Rid = b.ItvId AND b.Tp = 'TalJobFunc' and a.JobDescr=@job AND a.Rid=" & "@hfValue" & " Group BY a.TalentId, b.Tp, a.JobDescr) b on a.TpId=b.TalentId and a.Cdt=b.Cdt AND a.Tp=b.Tp" ,Hts)

            Dim talfid As HiddenField = TryCast(e.Row.FindControl("hf_TalentFid"), HiddenField)
            lbtn_resume.Visible = IIf(rsmFid = "0", False, True)
            lbtn_jfunc.Visible = IIf(jFid = "", False, True)
            lbtn_olink.Visible = IIf(talfid.Value = 0, False, True)
            ' If e.Row.Cells(1).Text = "未錄取" Then
            '                ddl_Status.Enabled = False
            '           Else
            '              ddl_Status.Enabled = True
            '         End If

        End If
    End Sub

    Protected Sub lbtn_Outlink_Click(sender As Object, e As EventArgs)
        Dim btn As LinkButton = TryCast(sender, LinkButton)
        Dim gr As GridViewRow = btn.NamingContainer
        If gr IsNot Nothing Then
            Dim hf_TalentId As HiddenField = TryCast(gr.FindControl("hf_TalentId"), HiddenField)
            Dim hf_ItvId As HiddenField = TryCast(gr.FindControl("hf_ItvId"), HiddenField)
            If hf_TalentId.Value <> "" AndAlso hf_ItvId.Value <> "" Then

                'jassy sql injection
                mas.ShowFile("SELECT TOP 1 Data,FileName FROM Attachment..NewEmp WHERE Tp='Talent' AND TpId='" & hf_TalentId.Value & "' ORDER BY Rid DESC ", mas.Conn)
            End If
        End If
    End Sub

    Protected Sub lbtn_JobFunc_Click(sender As Object, e As EventArgs)
        Dim btn As LinkButton = TryCast(sender, LinkButton)
        Dim gr As GridViewRow = btn.NamingContainer
        If gr IsNot Nothing Then
            Dim hf_TalentId As HiddenField = TryCast(gr.FindControl("hf_TalentId"), HiddenField)
            Dim hf_ItvId As HiddenField = TryCast(gr.FindControl("hf_ItvId"), HiddenField)
            If hf_TalentId.Value <> "" AndAlso hf_ItvId.Value <> "" Then
                'jassy sql injection                
                mas.ShowFile("Declare @job nvarchar(100) = '',@rid integer= 0 select @job = JobDescr from ItvMas where Rid=" & hf_ItvId.Value & " select @rid=a.Rid FROM   Attachment.dbo.NewEmp AS a JOIN " _
                    & " (SELECT a.TalentId, b.Tp, a.JobDescr, MAX(b.Cdt) AS Cdt FROM    dbo.ItvMas AS a INNER JOIN Attachment.dbo.NewEmp AS b ON a.Rid = b.ItvId AND b.Tp = 'TalJobFunc' and a.JobDescr=@job AND a.Rid='" & hf_ItvId.Value & "' Group BY a.TalentId, b.Tp, a.JobDescr) b on a.TpId=b.TalentId and a.Cdt=b.Cdt AND a.Tp=b.Tp " _
                    & " Select Data,FileName FROM Attachment..NewEmp WHERE Tp='TalJobFunc' AND Rid=@rid order by Rid desc", mas.Conn)
            End If
        End If
    End Sub

    Protected Sub btn_list_Click(sender As Object, e As EventArgs)
        'PageLoad("SqlConnectNewEmpV2")
        Dim sdt1 As String, blist_sql As String = ""
        If txt_from.Text = "" Then
            sdt1 = DateTime.Today.ToString("yyyy/MM/dd")
        Else
            sdt1 = txt_from.Text
        End If


        Hts.Clear()
        Hts.Add("@sdt1",  sdt1 )

        blist_sql = "select distinct 面試日期=Convert(varchar(16),a.ItvDt,20), 面試地點=c.Descr, 應徵者=b.Name, 承辦人=NewEmpV2.dbo.getEmpBase2(a.HR,'Mas1') " _
           & " from NewEmpV2..ItvMas a join TalentMas b On a.TalentId=b.Rid  left join NewEmpV2.dbo.NewCode c On a.Place=c.Code And c.Tp='Place' " _
           & " where Convert(varchar(10), Convert(datetime,a.ItvDt),111)=" & "@sdt1" & " and a.Status not in ('X','S','R')"

        excel_Click("面試人員名冊", "", blist_sql ,Hts )
    End Sub

    Protected Sub btn_excel_Click(sender As Object, e As EventArgs)
        excel_Click("面試資料", "", sql_Excel("sql"),sql_Excel("Hts"))
    End Sub



    Function sql_Excel( type as String)
        Dim sql As String = "", Msql As String = "", Msql_from As String = "", Msql_where As String = "", Endsql As String = ""  ,topstr As String = ""

        Msql = "declare @hr varchar(500) select @hr=''  Select @hr=@hr+HR from ItvMas   group by HR  "



        Msql = Msql & " SELECT distinct  btn_visible='True' , [狀態]=dbo.getItvBase2(a.Rid,'Status_CHI') ," &
"[人選姓名]=c.Name ,[Email]=c.Email, [電話]=c.Tel  ," &
"[日期]=convert(varchar(10),convert(datetime,a.ItvDt),111),[時間]=a.ItvTm,[面試類別]=f.Descr, " &
"[面試職務]=a.JobDescr,[面試地點]=h.Descr " &
",[面試主管]= dbo.getItvBase2(a.Rid,'ItvMngGrade') " &
",[面試HR]=  dbo.getItvBase2(a.Rid,'ItvHRGrade')  " &
",'Remark'=Base.dbo.GetXml(a.Remark,'HRLog')  + Case when dbo.GetXml(a.Remark,'SReasonTp')<>'' Then '&nbsp; 未到談分類：'+dbo.GetXml(a.Remark,'SReasonTp') Else '' End  + Case when dbo.GetXml(a.Remark,'SReason')<> '' Then '/未到談說明：' " &
"+dbo.GetXml(a.Remark,'SReason') Else '' End  + Case when dbo.GetXml(a.Remark,'ContactDt')<> '' Then '(下次聯絡日：' + dbo.GetXml(a.Remark,'ContactDt') + ')' Else '' End  + " &
"Case when dbo.GetXml(a.Remark,'DReasonTp')<> '' Then '&nbsp; 未錄取分類：'+ dbo.GetXml(a.Remark,'DReasonTp')  Else '' End  + Case when dbo.GetXml(a.Remark,'DReason')<> '' Then '/未錄取說明：' + dbo.GetXml(a.Remark,'DReason')  " &
"Else '' End  + Case when dbo.GetXml(a.Remark,'Enter')<> '' Then '<BR/>  錄取說明：' + dbo.GetXml(a.Remark,'Enter')  Else '' End  + Case when dbo.GetXml(a.Remark,'NDt')<> '' Then '<BR/>  報到取消日：' + dbo.GetXml(a.Remark,'NDt') " &
"        Else '' End  + Case when dbo.GetXml(a.Remark,'NReasonTp')<> '' Then '<BR/> 未報到分類：' + dbo.GetXml(a.Remark,'NReasonTp')  Else '' End  + Case when dbo.GetXml(a.Remark,'NReason')<> '' Then '/ ' + dbo.GetXml(a.Remark,'NReason')  " &
"        Else '' End  , " &
"        TalLevel=Isnull(j.Descr,'')+'<Br>'+Isnull(l.Descr,'')+' / '+Isnull(m.Descr,'')" &
"        ,'ItvId'=a.Rid" &
"        ,'TalentId'=c.Rid" &
"        ,TalentFid=Isnull(k.TalentId,0)" &
",Editor=dbo.getEmpBase2(a.Editor,'Mas1') " &
"  , a.ItvDt,a.ItvTm  " &
" ,[面試主管1]=dbo.[getEmpBase2]( substring(a.Manager,1,9),'Name'),[面試部門1]=dbo.[getEmpBase2]( substring(a.Manager,1,9),'Dept2') " &
" ,[面試處部門1]= dbo.[getEmpBase2]( substring(a.Manager,1,9),'Division')  " &
" ,[面試主管2]=dbo.[getEmpBase2]( substring(rtrim(a.Manager),11,9),'Name'),[面試部門2]=dbo.[getEmpBase2]( substring(rtrim(a.Manager),11,9),'Dept2') " &
" ,[面試主管3]=dbo.[getEmpBase2]( substring(rtrim(a.Manager),21,9),'Name'),[面試部門3]=dbo.[getEmpBase2]( substring(rtrim(a.Manager),21,9),'Dept2') " &
" ,[面試主管4]=dbo.[getEmpBase2]( substring(rtrim(a.Manager),31,9),'Name'),[面試部門4]=dbo.[getEmpBase2]( substring(rtrim(a.Manager),31,9),'Dept2') " &
" ,[面試主管5]=dbo.[getEmpBase2]( substring(rtrim(a.Manager),41,9),'Name'),[面試部門5]=dbo.[getEmpBase2]( substring(rtrim(a.Manager),41,9),'Dept2') "  &
" ,[面試HR1]= dbo.[getEmpBase2]( substring(rtrim(a.HR),1,9),'Name') , [面試HR2]= dbo.[getEmpBase2]( substring(rtrim(a.HR),11,9),'Name') "  &
" ,[測驗成績_英文]=dbo.GetScore(a.Testing,'英文') , [測驗成績_邏輯]=dbo.GetScore(a.Testing,'邏輯') , [測驗成績_專業]=dbo.GetScore(a.Testing,'專業') " &
" ,[Job Family]=j.Descr,	[分類]=i.JobLvl	,[評等]=i.TalentLvl " &
 " , [File]=case when Len(isnull(p.FileName,''))>0 then '職能;' else '' end  " &
" + case when Len(isnull(q.FileName,''))>0 then '外部;' else '' end " &
" + case when Len(isnull(r.BirthY,''))>0 then '履歷;' else '' end "  &
" ,Testing "

        Msql_from = " into #temp FROM ItvMas a LEFT JOIN ItvDet b ON a.Rid=b.ItvId " &
 "LEFT JOIN ApplyMas aa on " &
 "/*aa.TalentId = a.TalentId and*/ aa.ItvId = a.Rid and aa.Status not in ('D') " &
 "LEFT JOIN TalentMas c ON a.TalentId=c.Rid " &
 "LEFT JOIN NewCode f ON f.Code=a.ItvTp And f.Tp='ItvTp' " &
 "        LEFT JOIN NewCode g ON ISNULL(aa.Status,a.Status)=g.Code AND g.Tp='Status' " &
 "LEFT JOIN NewCode h on a.Place=h.Code And h.Tp='Place'  " &
 "        left join TalentDet i on c.Rid=i.TalentId " &
 "Left Join Base..BaseCode j on i.JobFamily = j.Code And j.Tp='JobFamily' " &
 "        Left Join NewCode l on i.JobLvl=l.Code And l.Tp='JobLvl' " &
 "        Left Join NewCode m on i.TalentLvl=m.Code And m.Tp='TalentLvl' " &
 "        Left JOIN dbo.AttachView k on c.Rid=k.Rid  " &
" Left join Attachment..NewEmp p on  a.TalentId=p.TpId and p.Tp='Talent'  " &
" Left join Attachment..NewEmp q on  a.TalentId=p.TpId and p.Tp='TalJobFunc' " &
" Left join NewEmpV2..ItvEmp r on    a.TalentId=r.TalentId and r.ItvId=b.ItvId  And BirthY <> '' "





        Msql_where = " WHERE      a.Status not in ('D') "

        Hts.Clear()

        If ck_date.Checked = True Then


            Hts.Add("@txt_fromText",  txt_from.Text )
            Hts.Add("@txt_toText",  txt_to.Text )

            Msql_where = Msql_where & " and a.ItvDt>=convert(datetime," & "@txt_fromText" & ")  AND a.ItvDt<=convert(datetime," & "@txt_toText" & ") "

        End If
        If ck_edit_date.Checked = True Then

            Hts.Add("@txt_edit_fromText",  txt_edit_from.Text )

            Msql_where = Msql_where & " and a.Udt>=convert(datetime," & "@txt_edit_fromText" & ")  AND a.Udt<=dateadd(d,1,convert(datetime,'" & txt_edit_to.Text & "') )"

        End If

        If ck_interviewer.Checked = True And Txt_interviewer.Text <> "" Then

            Hts.Add("@Txt_interviewerfromText",  Txt_interviewer.Text )

            Msql_where = Msql_where & " and c.Name  like '%'+" & "@Txt_interviewerfromText" & "+'%' "
        End If

        If ck_boss.Checked = True And txt_boss.Text <> "" Then

            Hts.Add("@txt_bossText",  txt_boss.Text )
            Msql_where = Msql_where & " and dbo.getItvBase2(a.Rid,'ItvMngGrade') like '%'+" & "@txt_bossText" & "+'%'"
        End If

        If ck_place.Checked = True And getmutiListvalue(lst_place) <> "" Then

            Hts.Add("@txt_bossText",  getmutiListvalue(lst_place) )
            Msql_where = Msql_where & " and a.Place in (" & "@lst_place" & ")"
        End If


        If ck_status.Checked = True Then
            If Not (lst_status.SelectedValue = "") Then
                If lst_status.SelectedValue = "HC" Then
                    Msql_where = Msql_where + " and aa.Status IN ('H','C') "
                ElseIf lst_status.SelectedValue = "AS" Then
                    Msql_where = Msql_where & " and dbo.getItvBase2(a.Rid,'Status') IN ('A','S') "

                Else

                    Hts.Add("@lst_statusSelectedValue",  lst_status.SelectedValue )
                    Msql_where = Msql_where & " and charindex( dbo.getItvBase2(a.Rid,'Status')," & "@lst_statusSelectedValue" & " )>0 "
                End If
            End If
        End If
        If ck_hr.Checked = True Then
            If Not (lst_hr.SelectedValue = "") Then
                Hts.Add("@lst_hrSelectedValue",  lst_hr.SelectedValue )

                Msql_where = Msql_where & " and charindex( " & "@lst_hrSelectedValue" & ", dbo.getItvBase2(a.Rid,'ItvHR')  )>0 "
            End If
        End If




        If ck_site.Checked = True Then

            If Not (ddl_site.SelectedValue = "") Then


                Hts.Add("@ddl_siteSelectedValue", ddl_site.SelectedValue )

                Msql_where = Msql_where & "  and ( left(a.HR,9) in (	 select s from Base.dbo.Split(@hr,';') a inner  join  Base.dbo.EmpMas b on a.s=b.Eid  and Place=case " &
                "@ddl_siteSelectedValue" & " when '士林' then 'S' else 'H'   end   ) " &
                  " or Right(a.HR,9) in (  select s from Base.dbo.Split(@hr,';') a inner  join  Base.dbo.EmpMas b on a.s=b.Eid and Place=case " &
                  "@ddl_siteSelectedValue" & " when '士林' then 'S' else 'H'   end   ))"

            End If
        End If


        Msql =  Msql & Msql_from & Msql_where

        Msql = Msql & " select [狀態],[人選姓名],[Email],[電話],[面試職務],[面試地點],[日期],[時間],[面試類別],[面試主管1],[面試部門1],[面試處部門1],[面試主管2],[面試部門2],[面試主管3],[面試部門3],[面試主管4],[面試部門4],[面試主管5],[面試部門5],[面試HR1],[面試HR2],[測驗成績_英文],[測驗成績_邏輯],[測驗成績_專業],[Job Family],[分類],[評等] ,[File],[測驗成績]=Testing from  #temp "

        Msql = Msql &  " ORDER BY ItvDt,ItvTm "



        'Response.Write(Msql)
        'return ""

        'return Msql
        'ClientMsg(Msql)

        if type="Hts" then
            return Hts
        else
            return Msql
        end if


    End Function


    Protected Sub excel_Click(ByVal title As String, ByVal titel2 As String, ByVal sql As String , hts As Hashtable)
        Dim dt As Data.DataTable = CType(Application.Item("MyDataTable"), Data.DataTable)
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")
        Response.ContentType = "application/ms-Excel"
        dt = Funcs.sqlTable(sql , hts)
        Response.AddHeader("Content-Disposition", "inline;filename=report.xls")


        If dt.Columns.Count > 0 Then
            Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
            Response.Write("<table border=""1""><tr><td style=""font-size:large;font-weight:bold"" colspan=" & dt.Columns.Count & ">" & title & "</td></tr>")
            Response.Write(ConvertToTable(dt))
            Response.Write("</table>")
        End If

        Response.End()
    End Sub
    Private Function ConvertToTable(ByVal dt As Data.DataTable) As String
        Dim dr As Data.DataRow, ary() As Object, i As Integer
        Dim iCol As Integer
        Response.Write("<tr>")
        For iCol = 0 To dt.Columns.Count - 1
            Response.Write("<td style=background-color:Yellow>" & dt.Columns(iCol).ToString & "</td>")
        Next
        Response.Write("</tr>")
        For Each dr In dt.Rows
            ary = dr.ItemArray
            Response.Write("<tr>")
            For i = 0 To UBound(ary)
                Response.Write("<td>" & ary(i).ToString & if (Funcs.IsNumeric(ary(i).ToString),"&nbsp" ,"") & "</td>")
            Next
            Response.Write("</tr>")
        Next
        Return Nothing
    End Function
</script>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title> 面試管理暨查詢 (HR權限)  </title>
     <link rel="stylesheet" type="text/css" href="StyleSheet1.css" />
 <SCRIPT  type="text/javascript" language="javascript">
     function winopen(data) {
         document.getElementById("Str").value = data;
         url = "Mng_ItvHR_V2_son.aspx";
        var searchopen=window.open (url, 'newwindow', 'height=500, width=700,top=100,left=300,toolbar=no,menubar=no,scrollbars=yes,resizable=no,location=n o, status=no');
　}

 </script>

    </head>
<body>
    <form id="form1" runat="server">
        <div>
            <input type="hidden" Nmae="Str" id="Str"  runat="server"/>
            <table width="100%"   style="text-align:left;vertical-align:central" >
            <tr><td colspan="8"  style="text-align:center;" >
            <h2>面試管理暨查詢 (HR權限)</h2>
                </td></tr>
                  <tr><td colspan="8">
                      <table style="vertical-align: baseline"><tr><td>
                      面試起日
                    
                    <asp:TextBox ID="txt_from" runat="server" Width="150px"></asp:TextBox>
                    迄日
                    <asp:TextBox ID="txt_to" runat="server" Width="150"></asp:TextBox>
                          </td><td>
                                    <label class="switch">
  <input type="checkbox" id="ck_date" runat="server">
  <span class="slider"></span>
</label>
                                 </td></tr></table>
                              
                          
<td>人選<br/>姓名</td><td>
                    <table style="vertical-align:central"><tr><td>
                    <asp:TextBox ID="Txt_interviewer" runat="server" Width="150"></asp:TextBox>
                    <br/><asp:Label ID="Label2" runat="server" Text="可輸入姓名其中1個字"  Font-Bold="true"   Font-Size="X-Small" ForeColor="#FF0066"></asp:Label>
</td><td>
                    <label class="switch">
  <input type="checkbox" id="ck_interviewer" runat="server">
  <span class="slider"></span>
</label>
    </td></tr></table>
                                  </td>
                      
                      <td>面試<br/>主管</td><td>
                          <table style="vertical-align:central"><tr><td>
                    <asp:TextBox ID="txt_boss" runat="server" Width="150px"></asp:TextBox>
                               <br/>    <asp:Label ID="Label1" runat="server" Text="可輸入姓名其中1個字或電話"  Font-Bold="true" Font-Size="X-Small" ForeColor="#FF0066"></asp:Label>
 
                   </td><td>
                    <label class="switch">
  <input type="checkbox" id="ck_boss" runat="server">
  <span class="slider"></span>
</label></td></tr></table>



                                                           </td>
                          <td>
                              <table ><tr><td  style="vertical-align:central">
                            廠區
                           
                                 </td>
                                 <td style="vertical-align: bottom">
                                 <asp:DropDownList ID="ddl_site" runat="server" AutoPostBack  width="100px" OnSelectedIndexChanged="ddl_site_SelectedIndexChanged">
                                 </asp:DropDownList>
                                      <label class="switch">  <input type="checkbox" id="ck_site" runat="server">  <span class="slider"></span></label>
                                 </td>
                                 </table>
								 
								 
                              </td>
							  
				                          <td>
                              <table ><tr> <td>
							  <asp:Label ID="Label88" runat="server" Text="快速查詢" Font-Bold="True"  ForeColor="Red"></asp:Label>
           <asp:DropDownList ID="ddl_top" runat="server"   width="100px" >
          </asp:DropDownList>
</td> </tr>	</table>


                         
								 
								 
                              </td>			  
							  
                                                              </tr></table>

                      </td></tr>
              
                <tr>
                    <td colspan="8"> <table ><tr>
            <td>
                <table  style="vertical-align:central"><tr>
                             <td>
                                 編輯<br>日期<br>  
<label class="switch"> <input type="checkbox" id="ck_edit_date" runat="server"> <span class="slider"></span></label>
                             </td>
                             <td>
                         起日
                    
                    <asp:TextBox ID="txt_edit_from" runat="server" Width="150px"></asp:TextBox>
                                 <br/>
                    迄日
                    <asp:TextBox ID="txt_edit_to" runat="server" Width="150"></asp:TextBox>


                       </td></tr></table>

                </td>
            
            <td >
                         <table  style="vertical-align:central"><tr>
                             <td>
                                 表單<br>狀態<br>  
<label class="switch"> <input type="checkbox" id="ck_status" runat="server"> <span class="slider"></span></label>
                             </td>
                             <td>
                        <asp:ListBox ID="lst_status" runat="server" Width="150px"></asp:ListBox>
                       </td></tr></table>
                        </td>
                     


                        <td>

                             <table style="vertical-align:central"><tr><td>                       
                      
                        <td>面試<br />HR<br/>
<label class="switch">  <input type="checkbox" id="ck_hr" runat="server">  <span class="slider"></span></label>
                        </td> <td>
 <asp:ListBox ID="lst_hr" runat="server" Width="120px"></asp:ListBox>
     </td></tr></table>

                            </td>


                        <td>
                              <table style="vertical-align:central"><tr><td> 
                   面試<br/>地點<br/>
                            <label class="switch">  <input type="checkbox" id="ck_place" runat="server">  <span class="slider"></span></label>

                        </td><td>
 <asp:ListBox ID="lst_place" runat="server" Width="100px" SelectionMode="Multiple"></asp:ListBox>
     </td></tr></table>       
                           </td>       
                    
                <td  style="text-align:left;width:200px;" >
                    <asp:Label ID="Label3" runat="server" Text="查詢條件" Font-Bold="True"  ForeColor="Red"></asp:Label>
                    <br />
                    <asp:RadioButton  ID="rb1" GroupName="certia1" Text="HR維護" runat="server" OnCheckedChanged="rb1_CheckedChanged"  AutoPostBack />
                    <br/><asp:RadioButton  ID="rb2" GroupName="certia1" Text="進行任用簽核名單" runat="server" OnCheckedChanged="rb1_CheckedChanged" AutoPostBack />
                    
                     
                    </td>
                        <td style="text-align:left;width:300px">
                             <asp:RadioButton  ID="rb3" GroupName="certia1" Text="今日作業" runat="server" OnCheckedChanged="rb1_CheckedChanged" AutoPostBack  />
                       <br/>     <asp:RadioButton  ID="rb4" GroupName="certia1" Text="清空條件" runat="server" OnCheckedChanged="rb1_CheckedChanged" AutoPostBack  />
                   <br/><br/>
                            <asp:Button ID="btn_search" runat="server"   Text="查詢" OnClick="btn_search_Click"   />
                     
                    <asp:Button ID="btn_search0" runat="server"   Text="更新" OnClick="btn_search_Click"   />
                            
                    <asp:Button ID="btn_list" runat="server"   Text="列印名冊" OnClick="btn_list_Click"   />
					<asp:Button ID="btn_excel" runat="server"   Text="匯出Excel" OnClick="btn_excel_Click"   />

                            </td>

                                                       </tr>
                
            <tr><td colspan="8" style="text-align:center">
            <asp:GridView ID="gd_view" runat="server" CellPadding="4" AutoGenerateColumns="false" BackColor="White" BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px"  OnRowDataBound="gv_list_RowDataBound" >
                <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
                <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#0033CC" />
                <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
                <RowStyle BackColor="White" ForeColor="#003399" />
                <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                <SortedAscendingCellStyle BackColor="#EDF6F6" />
                <SortedAscendingHeaderStyle BackColor="#0D4AC4" />
                <SortedDescendingCellStyle BackColor="#D6DFDF" />
                <SortedDescendingHeaderStyle BackColor="#002876" />
                <Columns>

                                <asp:TemplateField HeaderText="">
                                    <ItemTemplate>
                                       <asp:Button runat="server" id="btn_sel"  text="修改" OnClick="btn_sel_click" Visible='<%# Eval("btn_visible")%>' OnClientClick=<%# Eval("TalentId", "javascript:winopen('{0}@" & Eval("ItvId", "{0}") & "')")%> />
                                        <asp:HiddenField runat="server" ID="hf_ItvId" Value='<%# Eval("ItvId")%>' />
                                        <asp:HiddenField runat="server" ID="hf_TalentId" Value='<%# Eval("TalentId")%>' />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Status" HeaderText="狀態" />
                                <asp:BoundField DataField="Name" HeaderText="人選姓名<br/>電話" HtmlEncode="false" />
                                <asp:BoundField DataField="FamilyPlace" HeaderText="面試職務<br/>面試地點" HtmlEncode="false" />
                                <asp:BoundField DataField="DtTm" HeaderText="日期<br/>時間" HtmlEncode="false" />
                                <asp:BoundField DataField="EmpTp" HeaderText="面試類別" />
                                <asp:BoundField DataField="Mng" HeaderText="面試主管_評分" HtmlEncode="false" />
                                <asp:BoundField DataField="HR" HeaderText="面試HR_評分" HtmlEncode="false" />
                                <asp:TemplateField HeaderText="File">
                                    <ItemTemplate>
                                        <asp:HiddenField runat="server" ID="hf_TalentFid" Value='<%# Eval("TalentFid")%>' />
                                    <asp:LinkButton runat="server" ID="lbtn_Resume" Text="履歷" OnClick="lbtn_Resume_Click" />
                                        <asp:LinkButton runat="server" ID="lbtn_Outlink"  Text="外部" OnClick="lbtn_Outlink_Click" />
                                        <asp:LinkButton runat="server" ID="lbtn_JobFunc"  Text="職能" OnClick="lbtn_JobFunc_Click" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Testing" HeaderText="測驗成績" />
                                <asp:BoundField DataField="Remark" HeaderText="備註" HtmlEncode="false" />
                                <%--<asp:BoundField DataField="Editor" HeaderText="編輯者<br>分類／評等" HtmlEncode="false" />--%>
                                <asp:BoundField DataField="TalLevel" HeaderText="Job Family<br/>分類/評等" HtmlEncode="false" />
                                <asp:BoundField DataField="Editor" HeaderText="編輯者" HtmlEncode="false" />
                            </Columns>
                            <EmptyDataTemplate>
                                No Data
                            </EmptyDataTemplate>
            </asp:GridView>
                <!--      <asp:Button runat="server" id="btn_sel"  text="修改" OnClientClick='<%# Eval("ItvId", "javascript:winopen({0});")%>'/>-->
                </tr></table>

        </div>
    </form>
</body>
</html>
