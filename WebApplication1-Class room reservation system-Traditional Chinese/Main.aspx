<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- #include File="OleDbFunction.inc" --><HTML><HEAD>
		<title>�M�~�Ыǹw���t��</title>
		<SCRIPT language="VB" Runat="Server">

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

				Sub Page_Load( sender As Object, e As Eventargs )
        Session("LV")=Nothing
        Load_Holidays()
		  	 		If Not IsPostBack Then BindList_PS()
		  	 		If Not IsPostBack Then BindList()
				End Sub

				Sub BindList_PS()
  				  myCalendar.SelectedDate = DateString()

					   Dim strSQL As String 

			     strSQL = "Select [P-ID],[P-Adds] From [PhotoStudio] " 
					   myRadioButtonList.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "PhotoStudio")
					   myRadioButtonList.DataTextField="P-Adds"
					   myRadioButtonList.DataValueField="P-ID"
					   myRadioButtonList.DataBind()
					   myRadioButtonList.SelectedIndex = 0
    End Sub

				Sub BindList()
					   Dim strSQL As String 
 	  			 strSQL = "Select [OrderMenu].[O-T_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name] From [OrderMenu],[Teacher],[Book],[Class] " & _
				     	 		    "Where [OrderMenu].[O-Date] Like #" & DateString() & "#" & " And [OrderMenu].[O-P_ID] = " & Clng(myRadioButtonList.SelectedItem.Value) & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"
        myDataGrid.CurrentPageIndex = 0
					   myDataGrid.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "OrderMenu")
					   myDataGrid.DataBind()
				End Sub

				Sub BindList2()
				    Dim strSQL As String 
        If myCalendar.SelectedDates.Count = 1 Then
           strSQL = "Select [OrderMenu].[O-T_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name] From [OrderMenu],[Teacher],[Book],[Class] " & _
    			             "Where [OrderMenu].[O-Date] Like #" & DateValue(myCalendar.SelectedDate) & "#" & " And [O-P_ID] = " & Clng(myRadioButtonList.SelectedItem.Value) & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"
           myDataGrid.CurrentPageIndex = 0

        Else If myCalendar.SelectedDates.Count > 1 Then
           Dim strFirstDate As String 
           Dim strLastDate As String 
           With myCalendar.SelectedDates
                strFirstDate = .Item(0)
                strLastDate = .Item(.Count-1)
           End With
           strSQL = "Select [OrderMenu].[O-T_ID],[OrderMenu].[O-B_ID],[OrderMenu].[O-C_ID],[OrderMenu].[O-P_ID],[OrderMenu].[O-Date],[OrderMenu].[O-Time],[Teacher].[T-Name_LN_CHT],[Teacher].[T-Name_FN_CHT],[Book].[B-Name],[Class].[C-Name] From [OrderMenu],[Teacher],[Book],[Class] " & _
    			             "Where [OrderMenu].[O-Date] Between #" & strFirstDate & "# And #" & strLastDate & "# And [O-P_ID] = " & Clng(myRadioButtonList.SelectedItem.Value) & " And ( [OrderMenu].[O-T_ID] = [Teacher].[T-ID] And [OrderMenu].[O-B_ID] = [Book].[B-ID] And [OrderMenu].[O-C_ID] = [Class].[C-ID]) Order By [OrderMenu].[O-Date],[OrderMenu].[O-Time]"

        End If       
				    myDataGrid.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "OrderMenu")
				    myDataGrid.DataBind()
				End Sub

    Sub myDataGrid_PageIndexChanged( sender As Object, e As DataGridPageChangedEventArgs )
        myDataGrid.CurrentPageIndex = e.NewPageIndex
        Call BindList2()
    End Sub

				Sub DayChange( sender As Object, e As Eventargs )
        myDataGrid.CurrentPageIndex = 0 
				    Call BindList2()
				End Sub

				
				Sub PhotoStudioChange ( sender as Object, e As Eventargs )
        myDataGrid.CurrentPageIndex = 0 
				    Call BindList2()
				End Sub

				
				Sub Login ( sender as Object, e As Eventargs )
				    Dim objCnn As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                          "Data Source=" & Server.MapPath("./DB/order.mdb") )
				    objCnn.Open()
				    Dim strSQL As String
				    If rbnLevel1.Checked = True Then
			        strSQL = "Select [A-ID],[A-LN],[A-PW] From [Admin] Where [A-LN] Like '" & txtLN.Text & "'"
				    End If
				    If rbnLevel2.Checked = True Then
			        strSQL = "Select [T-ID],[T-LN],[T-PW] From [Teacher] Where [T-LN] Like '" & txtLN.Text & "'"
				    End If
  				     
				    Dim objCmd As New OleDbCommand()
				    objCmd.Connection = objCnn
				    objCmd.CommandText = strSQL
				    Dim objReader As OleDbDataReader = objCmd.ExecuteReader()
				    If objReader.Read() = True Then
				       If objReader.Item(2) = txtPW.Text Then
				          If rbnLevel1.Checked = True Then
                 Session("LV")="Admin" 
				             Response.Redirect("./Admin-Frame.asp?ID=" & objReader.Item(0))
				          End If
				          If rbnLevel2.Checked = True Then
                 Session("LV")="Teacher" 
				             Response.Redirect("./Teacher-Frame.asp?ID=" & objReader.Item(0))
				          End If
				       End If
				    End If
				    objCnn.Close() 
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
		<meta content="True" name="vs_snapToGrid">
		<meta content="True" name="vs_showGrid">
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 10.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body background="./image/main-bg.gif" MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<FONT face="�s�ө���">
				<asp:label id="LabPS" style="Z-INDEX: 106; LEFT: 168px; POSITION: absolute; TOP: 72px" runat="server"
					Height="20px" Width="80px">�M�~�ЫǡG</asp:label><asp:radiobutton id="rbnLevel2" style="Z-INDEX: 114; LEFT: 72px; POSITION: absolute; TOP: 168px"
					runat="server" Height="24px" Width="76px" Text="�Юv" GroupName="Level" Checked="True"></asp:radiobutton><asp:label id="labLV" style="Z-INDEX: 112; LEFT: 8px; POSITION: absolute; TOP: 140px" runat="server"
					Height="16px" Width="53px">�����G</asp:label><asp:textbox id="txtPW" style="Z-INDEX: 111; LEFT: 68px; POSITION: absolute; TOP: 228px" tabIndex="2"
					runat="server" Height="20px" Width="68px" TextMode="Password"></asp:textbox><asp:label id="labPW" style="Z-INDEX: 110; LEFT: 8px; POSITION: absolute; TOP: 228px" runat="server"
					Height="16px" Width="53px">�K�X�G</asp:label><asp:radiobuttonlist id="myRadioButtonList" style="Z-INDEX: 105; LEFT: 256px; POSITION: absolute; TOP: 72px"
					runat="server" Height="44px" Width="424px" AutoPostBack="True" RepeatLayout="Flow" RepeatDirection="Horizontal" OnSelectedIndexChanged="PhotoStudioChange"
					Font-Size="10pt"></asp:radiobuttonlist>
				<DIV style="DISPLAY: inline; Z-INDEX: 101; LEFT: 12px; WIDTH: 96px; POSITION: absolute; TOP: 104px; HEIGHT: 24px"
					ms_positioning="FlowLayout">�z�O������</DIV>
				<asp:calendar id="myCalendar" style="Z-INDEX: 102; LEFT: 168px; POSITION: absolute; TOP: 120px"
					runat="server" Height="265px" Width="540px" SelectionMode="DayWeekMonth" ShowGridLines="True"
					OnDayRender="myCalendar_DayRender" BorderWidth="1px" OnSelectionChanged="DayChange" ForeColor="#663399"
					Font-Size="8pt" Font-Names="Verdana" BackColor="#FFFFCC" BorderColor="#FFCC66" NextPrevFormat="FullMonth"
					SelectWeekText="<img src='./image/Week.ico' border='0'></img>" SelectMonthText="<img src='./image/Month.ico' border='0'></img>">
					<TodayDayStyle ForeColor="White" BackColor="#FFCC66"></TodayDayStyle>
					<SelectorStyle BackColor="#FFCC66"></SelectorStyle>
					<NextPrevStyle Font-Size="9pt" ForeColor="#FFFFCC"></NextPrevStyle>
					<DayHeaderStyle Height="1px" BackColor="#FFCC66"></DayHeaderStyle>
					<SelectedDayStyle Font-Bold="True" BackColor="#CCCCFF"></SelectedDayStyle>
					<TitleStyle Font-Size="9pt" Font-Bold="True" ForeColor="#FFFFCC" BackColor="#990000"></TitleStyle>
					<OtherMonthDayStyle ForeColor="#CC9966"></OtherMonthDayStyle>
				</asp:calendar></FONT>
			<DIV style="Z-INDEX: 103; LEFT: 252px; WIDTH: 332px; POSITION: absolute; TOP: 16px; HEIGHT: 24px"
				ms_positioning="FlowLayout"><SPAN style="FONT-SIZE: 16pt; FONT-FAMILY: �s�ө���; mso-bidi-font-size: 12.0pt; mso-hansi-font-family: 'Times New Roman'; mso-bidi-font-family: 'Times New Roman'; mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW; mso-bidi-language: AR-SA"><STRONG>�M�~�Ыǹw���t��</STRONG>
				</SPAN>
			</DIV>
			<asp:datagrid id="myDataGrid" style="Z-INDEX: 104; LEFT: 168px; POSITION: absolute; TOP: 384px"
				runat="server" Width="540px" BorderWidth="1px" BackColor="White" BorderColor="#CC9966" AutoGenerateColumns="False"
				PageSize="8" CellPadding="4" BorderStyle="None" AllowPaging="True" OnPageIndexChanged="myDataGrid_PageIndexChanged">
				<FooterStyle ForeColor="#330099" BackColor="#FFFFCC"></FooterStyle>
				<HeaderStyle Font-Size="10pt" Font-Bold="True" ForeColor="#FFFFCC" BackColor="#990000"></HeaderStyle>
				<PagerStyle HorizontalAlign="Center" ForeColor="#330099" BackColor="#FFFFCC" Mode="NumericPages"></PagerStyle>
				<SelectedItemStyle Font-Bold="True" ForeColor="#663399" BackColor="#FFCC66"></SelectedItemStyle>
				<ItemStyle ForeColor="#330099" BackColor="White"></ItemStyle>
				<Columns>
					<asp:BoundColumn DataField="O-Date" HeaderText="���" DataFormatString="{0:d}"></asp:BoundColumn>
					<asp:BoundColumn DataField="O-Time" HeaderText="�`��"></asp:BoundColumn>
					<asp:BoundColumn DataField="T-Name_FN_CHT" HeaderText="�Юv"></asp:BoundColumn>
					<asp:BoundColumn DataField="B-Name" HeaderText="���"></asp:BoundColumn>
					<asp:BoundColumn DataField="C-Name" HeaderText="�Z��"></asp:BoundColumn>
				</Columns>
			</asp:datagrid><asp:label id="labLN" style="Z-INDEX: 109; LEFT: 8px; POSITION: absolute; TOP: 200px" runat="server"
				Height="16px" Width="53px">�b���G</asp:label><asp:textbox id="txtLN" style="Z-INDEX: 107; LEFT: 68px; POSITION: absolute; TOP: 200px" tabIndex="1"
				runat="server" Height="20px" Width="68px"></asp:textbox><asp:button id="btnLogin" style="Z-INDEX: 108; LEFT: 96px; POSITION: absolute; TOP: 256px" onclick="Login"
				tabIndex="3" runat="server" Height="20px" Width="40px" Text="�n�J" BackColor="LightSteelBlue" BorderStyle="Dotted"></asp:button><asp:radiobutton id="rbnLevel1" style="Z-INDEX: 113; LEFT: 72px; POSITION: absolute; TOP: 140px"
				runat="server" Height="24px" Width="76px" Text="�޲z��" GroupName="Level"></asp:radiobutton></form>
	</body>
</HTML>
