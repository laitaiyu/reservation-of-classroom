<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
  <HEAD>
<!-- #include File="OleDbFunction.inc" -->
<meta content=True name=vs_snapToGrid>
<meta content="True" name=vs_showGrid>
<meta content="Microsoft Visual Studio .NET 7.1" name=GENERATOR>
<meta content="Visual Basic .NET 7.1" name=CODE_LANGUAGE>
<meta content=JavaScript name=vs_defaultClientScript>
<meta content=http://schemas.microsoft.com/intellisense/ie5 name=vs_targetSchema>
<Script Language="VB" Runat="Server">

  Dim holidays(12,31) as String

  Sub Load_Holidays()
      holidays(01,01) ="*���� ���إ���}�������"
      holidays(01,11) ="�q�k�`"
      holidays(01,15) ="�Įv�`"
      holidays(01,23) ="�ۥѤ�"
      holidays(02,04) ="�A���`"
      holidays(02,14) ="���H�`"
      holidays(02,15) ="���@�`"
      holidays(02,28) ="*�M��������"
      holidays(03,01) ="�L�и`"
      holidays(03,05) ="���l�x�`"
      holidays(03,08) ="���k�`"
      holidays(03,12) ="�Ӿ�`"
      holidays(03,17) ="����`"
      holidays(03,20) ="�l�F�`"
      holidays(03,21) ="��H�`"
      holidays(03,25) ="���N�`"
      holidays(03,26) ="�s���`"
      holidays(03,29) ="�C�~�`"
      holidays(03,30) ="�X���`"
      holidays(04,01) ="�M�H�` �D�p�`"
      holidays(04,04) ="�����`"
      holidays(04,05) ="���ָ`"
      holidays(04,07) ="�å͸`"
      holidays(04,22) ="�@�ɦa�y��"
      holidays(05,01) ="*�Ұʸ`"
      holidays(05,04) ="�����`"
      holidays(05,05) ="�R�и`"
      holidays(05,10) ="�]��`"
      holidays(05,12) ="�@�h�`"
      holidays(06,03) ="�T�ϸ`"
      holidays(06,06) ="�u�{�v�` ���Q�`"
      holidays(06,09) ="�K���`"
      holidays(06,15) ="ĵ��`"
      holidays(06,30) ="�|�p�v�`"
      holidays(07,01) ="�����` �����` �|�ȸ`"
      holidays(07,06) ="�X�@�`"
      holidays(07,11) ="����`"
      holidays(07,12) ="Ť�׸`"
      holidays(08,08) ="���˸`"
      holidays(08,14) ="�ŭx�`"
      holidays(09,01) ="�O�̸`"
      holidays(09,03) ="�x�H�`"
      holidays(09,09) ="��|�` �߮v�`"
      holidays(09,13) ="�k�ߤ�"
      holidays(09,28) ="�Юv�`"
      holidays(10,06) ="�ѤH�`"
      holidays(10,10) ="*��y������"
      holidays(10,21) ="�ع��`"
      holidays(10,25) ="�x�W���_�`"
      holidays(10,31) ="�U�t�`"
      holidays(11,01) ="�ӤH�`"
      holidays(11,11) ="�u�~�` �a�F�`"
      holidays(11,17) ="�ۨӤ��`"
      holidays(11,12) ="��v�` ���ؤ�ƴ_���`"
      holidays(11,21) ="���Ÿ`"
      holidays(12,05) ="�����` ���H�`"
      holidays(12,10) ="�H�v�`"
      holidays(12,12) ="�˧L�`"
      holidays(12,25) ="��ˬ����� �t�ϸ`"
      holidays(12,27) ="�ؿv�v�`"
      holidays(12,28) ="�q�H�`"
      holidays(12,31) ="���H�`"  
  End Sub

  Sub Page_Load(sender As Object, e As Eventargs)
      If Session("LV") <> "Admin" Then
         Page.Response.Redirect ("./Main.aspx")
         Exit Sub
      End If
      Load_Holidays()
      If Not IsPostBack Then BindList()
  End Sub

  Sub BindList()
					 Dim strSQL As String 
 	  		strSQL = "Select [OrderMenu].[O-ID],[OrderMenu].[O-T_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[PhotoStudio].[P-Adds] From [OrderMenu],[Teacher],[Book],[Class],[PhotoStudio] " & _
				     	 		  "Where [OrderMenu].[O-Date] Like #" & DateString() & "#" & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID]  And [OrderMenu].[O-P_ID] = [PhotoStudio].[P-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"
					 myDataList.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "OrderMenu")
					 myDataList.DataBind()
  End Sub

  Sub BindList2()

				    Dim strSQL As String 
        If myCalendar.SelectedDates.Count = 1 Then
 	  			    strSQL = "Select [OrderMenu].[O-ID],[OrderMenu].[O-T_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[PhotoStudio].[P-Adds] From [OrderMenu],[Teacher],[Book],[Class],[PhotoStudio] " & _
				     	 		       "Where [OrderMenu].[O-Date] Like #" & DateValue(myCalendar.SelectedDate) & "#" & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID]  And [OrderMenu].[O-P_ID] = [PhotoStudio].[P-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"

        Else If myCalendar.SelectedDates.Count > 1 Then
           Dim strFirstDate As String 
           Dim strLastDate As String 
           With myCalendar.SelectedDates
                strFirstDate = .Item(0)
                strLastDate = .Item(.Count-1)
           End With
 	  			    strSQL = "Select [OrderMenu].[O-ID],[OrderMenu].[O-T_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name],[PhotoStudio].[P-Adds] From [OrderMenu],[Teacher],[Book],[Class],[PhotoStudio] " & _
				     	 		       "Where [OrderMenu].[O-Date] Between #" & strFirstDate & "# And #" & strLastDate & "#" & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID]  And [OrderMenu].[O-P_ID] = [PhotoStudio].[P-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"

        End If       
					   myDataList.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "OrderMenu")
					   myDataList.DataBind()

  End Sub

	 Sub DayChange( sender As Object, e As Eventargs )
				  Call BindList2()
	 End Sub
 
  Sub DataList_EditCommand(sender As Object, e As DataListCommandEventArgs)
      myDataList.EditItemIndex = e.Item.ItemIndex
      BindList()
  End Sub

  Sub DataList_CancelCommand(sender As Object, e As DataListCommandEventArgs)
      myDataList.EditItemIndex = -1
      BindList()
  End Sub

  Sub ExecuteSQL(strSQL As String)
      Dim objConn As New OleDbConnection()
      objConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                 "Data Source=" & Server.MapPath("./DB/order.mdb")
      objConn.Open()
      Dim objCmd As New OleDbCommand(strSQL, objConn)
      objCmd.ExecuteNonQuery
      objConn.Close()
  End Sub

  Sub DataList_DeleteCommand(sender As Object, e As DataListCommandEventArgs)
    Dim strSQL As String = "Delete From [OrderMenu] Where [" & _
        myDataList.DataKeyField & "]=" & myDataList.DataKeys(e.Item.ItemIndex)
    ExecuteSQL(strSQL)
    myDataList.EditItemIndex = -1
    BindList()
  End Sub

  Sub DataList_UpdateCommand(sender As Object, e As DataListCommandEventArgs)
    Dim strO_Date As String = Ctype(e.Item.FindControl("txtO_Date"), TextBox).Text
    Dim UpdateDate As DateTime = DateTime.Now.Date()
    Dim strSQL As String = "Update [OrderMenu] Set [O-Date]='" & strO_Date & "' " & _
                           "Where [" & myDataList.DataKeyField & "]=" & myDataList.DataKeys(e.Item.ItemIndex)
    ExecuteSQL(strSQL)
    myDataList.EditItemIndex = -1
    BindList()
  End Sub

  Sub myCalendar_DayRender(sender As Object, e As DayRenderEventArgs)

      Dim d as CalendarDay
      Dim c as TableCell

      d = e.Day
      c = e.Cell

      If d.IsOtherMonth Then
         c.Controls.Clear
      Else
          Try
              Dim Hol As String = holidays(d.Date.Month,d.Date.Day)

              If Hol <> "" Then
                  c.Controls.Add(new LiteralControl("<br>" + Hol))
              End If
          Catch exc as Exception
              Response.Write (exc.ToString())
          End Try
      End If
  End Sub

</SCRIPT>
</HEAD>
  <Body  MS_POSITIONING="GridLayout" background=./image/admin-bg.gif>
    <form id=Form1 method=post runat="server">
      <Asp:DataList Runat="Server" Id="myDataList" CellPadding="3" Width="600px"
       HorizontalAlign="Center" OnEditCommand="DataList_EditCommand"
       OnUpdateCommand="DataList_UpdateCommand" OnDeleteCommand="DataList_DeleteCommand"
       OnCancelCommand="DataList_CancelCommand" DataKeyField="O-ID"
       ExtractTemplateRows="True" BorderColor="#DEBA84" GridLines="Both" RepeatLayout="Flow" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 328px" BackColor="#DEBA84" BorderWidth="1px" BorderStyle="None" CellSpacing="2">
<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#738A9C">
</SelectedItemStyle>

<HeaderTemplate>
          <Asp:Table Runat="Server" Width="600">
            <Asp:TableRow Runat="Server">
              <Asp:TableCell Runat="Server" Width="40" Font-Size ="10pt">
                �۰ʽs��
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                ���
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                �M�~�Ы�
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                �`��
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                �Юv
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                �ϥΤ��e
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                �ɥγ��
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                �\��
              </Asp:TableCell>
            </Asp:TableRow>
          </Asp:Table>
        
</HeaderTemplate>

<EditItemStyle BackColor="Lavender">
</EditItemStyle>

<ItemStyle ForeColor="#8C4510" BackColor="#FFF7E7">
</ItemStyle>

<ItemTemplate>
          <Asp:Table Runat="Server" Width="600">
            <Asp:TableRow Runat="Server" HorizontalAlign="Center">
              <Asp:TableCell Runat="Server" Width="40" Font-Size ="10pt" >
                <%# Container.DataItem("O-ID") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# DateValue(Container.DataItem("O-Date")).ToShortDateString() %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# Container.DataItem("P-Adds") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# Container.DataItem("O-Time") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# Container.DataItem("T-Name_FN_CHT") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# Container.DataItem("B-Name") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <%# Container.DataItem("C-Name") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="70" Font-Size ="10pt">
                <!--<Asp:LinkButton Runat="Server" Text="�s��" CommandName="Edit"/>-->
              </Asp:TableCell>
            </Asp:TableRow>
          </Asp:Table>
        
</ItemTemplate>

<FooterStyle ForeColor="#8C4510" BackColor="#F7DFB5">
</FooterStyle>

<HeaderStyle Font-Bold="True" HorizontalAlign="Center" ForeColor="White" BackColor="#A55129">
</HeaderStyle>

<EditItemTemplate>
          <Asp:Table Runat="Server" Width="400">
            <Asp:TableRow Runat="Server" HorizontalAlign="Center">
              <Asp:TableCell Runat="Server" Width="70">
                <%# Container.DataItem("O-ID") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="200">
                <Asp:TextBox Runat="Server" Id="txtBook" Width="200" Text='<%# Container.DataItem("O-Date") %>' />
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="130">
                <Asp:LinkButton Runat="Server" Width="18" Text="��s" CommandName="Update" /> 
              �@<Asp:LinkButton Runat="Server" Width="18" Text="�R��" CommandName="Delete" /> 
              �@<Asp:LinkButton Runat="Server" Width="18" Text="����" CommandName="Cancel" />
              </Asp:TableCell>
            </Asp:TableRow>
          </Asp:Table>
        
</EditItemTemplate>
      </Asp:DataList><asp:calendar id="myCalendar" style="Z-INDEX: 106; LEFT: 8px; POSITION: absolute; TOP: 60px" runat="server" BorderWidth="1px" BackColor="#FFFFCC" BorderColor="#FFCC66" Width="600px" Height="265px" Font-Size="8pt" ForeColor="#663399" SelectionMode="DayWeekMonth" ShowGridLines="True" OnDayRender="myCalendar_DayRender" OnSelectionChanged="DayChange" Font-Names="Verdana" NextPrevFormat="FullMonth" SelectWeekText="<img src='./image/Week.ico' border='0'></img>" SelectMonthText="<img src='./image/Month.ico' border='0'></img>">
<TodayDayStyle ForeColor="White" BackColor="#FFCC66">
</TodayDayStyle>

<SelectorStyle BackColor="#FFCC66">
</SelectorStyle>

<NextPrevStyle Font-Size="9pt" ForeColor="#FFFFCC">
</NextPrevStyle>

<DayHeaderStyle Height="1px" BackColor="#FFCC66">
</DayHeaderStyle>

<SelectedDayStyle Font-Bold="True" BackColor="#CCCCFF">
</SelectedDayStyle>

<TitleStyle Font-Size="9pt" Font-Bold="True" ForeColor="#FFFFCC" BackColor="#990000">
</TitleStyle>

<OtherMonthDayStyle ForeColor="#CC9966">
</OtherMonthDayStyle>
             </asp:calendar>
      <asp:Label id="Label1" style="Z-INDEX: 105; LEFT: 208px; POSITION: absolute; TOP: 12px" 
                 Width="176px" runat="server" Font-Bold="True" Height="36px" Font-Size="X-Large">�w��������</asp:Label>
    </form>
  </BODY>
</HTML>
