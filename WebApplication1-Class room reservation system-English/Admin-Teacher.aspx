<%@ Page CodeBehind="Admin-Teacher.aspx.vb" Language="vb" AutoEventWireup="True" Inherits="WebApplication1.Admin_Teacher" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
  <HEAD>
<!-- #include File="OleDbFunction.inc" -->
<meta content=True name=vs_snapToGrid>
<meta content=True name=vs_showGrid>
<meta content="Microsoft Visual Studio .NET 7.1" name=GENERATOR>
<meta content="Visual Basic .NET 7.1" name=CODE_LANGUAGE>
<meta content=JavaScript name=vs_defaultClientScript>
<meta content=http://schemas.microsoft.com/intellisense/ie5 name=vs_targetSchema>
<SCRIPT language=VB Runat="Server">

  Sub BindList()
      Dim strSQL As String = "Select * From [Teacher] Order By [T-ID] Desc"
      myDataList.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "[Teacher]")
      myDataList.DataBind()
  End Sub
 
  Sub Page_Load(sender As Object, e As Eventargs)
      If Session("LV") <> "Admin" Then
         Page.Response.Redirect ("./Main.aspx")
         Exit Sub
      End If
      If Not IsPostBack Then BindList()
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
    Dim strSQL As String = "Delete From [Teacher] Where [" & _
        myDataList.DataKeyField & "]=" & myDataList.DataKeys(e.Item.ItemIndex)
    ExecuteSQL(strSQL)
    myDataList.EditItemIndex = -1
    BindList()
  End Sub

  Sub DataList_UpdateCommand(sender As Object, e As DataListCommandEventArgs)
    Dim strT_Name As String          = Ctype(e.Item.FindControl("txtET_Name"), TextBox).Text
    Dim strT_Tel_Home As String      = Ctype(e.Item.FindControl("txtET_Tel_Home"), TextBox).Text
    Dim strT_Tel_Cellphone As String = Ctype(e.Item.FindControl("txtET_Tel_Cellphone"), TextBox).Text
    Dim strT_ADD_1 As String         = Ctype(e.Item.FindControl("txtET_ADD_1"), TextBox).Text
    Dim strT_LN As String            = Ctype(e.Item.FindControl("txtET_LN"), TextBox).Text
    Dim strT_PW As String            = Ctype(e.Item.FindControl("txtET_PW"), TextBox).Text
    Dim UpdateDate As DateTime = DateTime.Now.Date()
    Dim strSQL As String = "Update [Teacher] Set [T-Name_FN_CHT]='" & strT_Name & "'," & _
                                                "[T-TEL_Home]='" & strT_Tel_Home & "'," & _
                                                "[T-TEL_Cellphone]='" & strT_Tel_Cellphone & "'," & _
                                                "[T-ADD_1]='" & strT_ADD_1 & "'," & _
                                                "[T-LN]='" & strT_LN & "'," & _
                                                "[T-PW]='" & strT_PW & "' " & _
                           "Where [" & myDataList.DataKeyField & "]=" & myDataList.DataKeys(e.Item.ItemIndex)
    ExecuteSQL(strSQL)
    myDataList.EditItemIndex = -1
    BindList()
  End Sub

  Sub Teacher_Insert (sender As Object, e As Eventargs)
    If txtT_Name.Text = "" Then Exit Sub
    If txtT_LN.Text = "" Then Exit Sub
    If txtT_PW.Text = "" Then Exit Sub
    Dim strSQL As String = "Insert Into [Teacher] ([T-Name_FN_CHT],[T-TEL_Home],[T-TEL_Cellphone],[T-ADD_1],[T-LN],[T-PW]) " & _ 
                           "Values ('" & txtT_Name.Text & "','" & txtT_Tel_Home.Text & "','" & txtT_Tel_Cellphone.Text & "','" & txtT_ADD_1.Text & "','" & txtT_LN.Text & "','" & txtT_PW.Text & "')"
    ExecuteSQL(strSQL)
    txtT_Name.Text = ""
    txtT_Tel_Home.Text = ""
    txtT_Tel_Cellphone.Text = ""
    txtT_ADD_1.Text = ""
    txtT_LN.Text = ""
    txtT_PW.Text = ""
    BindList()
  End Sub
</SCRIPT>
</HEAD>
<BODY MS_POSITIONING="GridLayout" background=./image/admin-bg.gif>
<form id=Form1 method=post runat="server"><ASP:DATALIST id=myDataList style="Z-INDEX: 100; LEFT: 4px; POSITION: absolute; TOP: 296px" Runat="Server" CellSpacing="2" BorderStyle="None" BorderWidth="1px" BackColor="#DEBA84" RepeatLayout="Flow" GridLines="Both" BorderColor="#DEBA84" ExtractTemplateRows="True" DataKeyField="T-ID" OnCancelCommand="DataList_CancelCommand" OnDeleteCommand="DataList_DeleteCommand" OnUpdateCommand="DataList_UpdateCommand" OnEditCommand="DataList_EditCommand" HorizontalAlign="Center" Width="600px" CellPadding="3">
<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#738A9C">
</SelectedItemStyle>

<HeaderTemplate>
          <Asp:Table Runat="Server" Width="600">
            <Asp:TableRow Runat="Server">
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                ID
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                Name
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                Telephone
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                Cellphone
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="100" Font-Size="10pt">
                Address
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                Login name
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                Password
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="80" Font-Size="10pt">
                Function
              </Asp:TableCell>
            </Asp:TableRow>
          </Asp:Table>
        
</HeaderTemplate>

<EditItemStyle BackColor="Lavender">
</EditItemStyle>

<ItemStyle ForeColor="#8C4510" BackColor="#FFF7E7">
</ItemStyle>

<ItemTemplate>
<Asp:Table id="Table1" Runat="server" Width="600">
            <Asp:TableRow Runat="Server" HorizontalAlign="Center">
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <%# Container.DataItem("T-ID") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <%# Container.DataItem("T-Name_FN_CHT") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <%# Container.DataItem("T-Tel_Home") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <%# Container.DataItem("T-Tel_Cellphone") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="100" Font-Size="10pt">
                <%# Container.DataItem("T-ADD_1") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <%# Container.DataItem("T-LN") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <%# Container.DataItem("T-PW") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="80" Font-Size="10pt">
                <Asp:LinkButton Runat="Server" Text="Edit" CommandName="Edit"/>
              </Asp:TableCell>
            </Asp:TableRow>
          </Asp:Table>
</ItemTemplate>

<FooterStyle ForeColor="#8C4510" BackColor="#F7DFB5">
</FooterStyle>

<HeaderStyle Font-Bold="True" HorizontalAlign="Center" ForeColor="White" BackColor="#A55129">
</HeaderStyle>

<EditItemTemplate>
<Asp:Table id="Table2" Runat="server" Width="600">
            <Asp:TableRow Runat="Server" HorizontalAlign="Center">
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <%# Container.DataItem("T-ID") %>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <Asp:TextBox Runat="Server" Id="txtET_Name" Width="50" Text='<%# Container.DataItem("T-Name_FN_CHT") %>' Font-Size="10pt"/>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <Asp:TextBox Runat="Server" Id="txtET_Tel_Home" Width="50" Text='<%# Container.DataItem("T-TEL_Home") %>' Font-Size="10pt"/>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <Asp:TextBox Runat="Server" Id="txtET_Tel_Cellphone" Width="50" Text='<%# Container.DataItem("T-TEL_Cellphone") %>' Font-Size="10pt"/>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="100" Font-Size="10pt">
                <Asp:TextBox Runat="Server" Id="txtET_ADD_1" Width="100" Text='<%# Container.DataItem("T-ADD_1") %>' Font-Size="10pt"/>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <Asp:TextBox Runat="Server" Id="txtET_LN" Width="50" Text='<%# Container.DataItem("T-LN") %>' Font-Size="10pt"/>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="50" Font-Size="10pt">
                <Asp:TextBox Runat="Server" Id="txtET_PW" Width="50" Text='<%# Container.DataItem("T-PW") %>' Font-Size="10pt"/>
              </Asp:TableCell>
              <Asp:TableCell Runat="Server" Width="100" Font-Size="10pt">
                <Asp:LinkButton Runat="Server" Width="18" Text="Update" CommandName="Update" Font-Size="10pt"/>
               , <Asp:LinkButton Runat="Server" Width="18" Text="Delete" CommandName="Delete" Font-Size="10pt"/>
               , <Asp:LinkButton Runat="Server" Width="18" Text="Cancel" CommandName="Cancel" Font-Size="10pt"/>
              </Asp:TableCell>
            </Asp:TableRow>
          </Asp:Table>
</EditItemTemplate>
</ASP:DATALIST>
      <asp:textbox id=txtT_PW style="Z-INDEX: 115; LEFT: 164px; POSITION: absolute; TOP: 232px" 
                   runat="server" Width="308px" Font-Size="10pt"></asp:textbox>
      <asp:textbox id=txtT_LN style="Z-INDEX: 114; LEFT: 164px; POSITION: absolute; TOP: 200px" 
                   runat="server" Width="308px" Font-Size="10pt"></asp:textbox>
      <asp:textbox id=txtT_ADD_1 style="Z-INDEX: 112; LEFT: 164px; POSITION: absolute; TOP: 168px" 
                   runat="server" Width="308px" Font-Size="10pt"></asp:textbox>
      <asp:textbox id=txtT_Tel_Cellphone style="Z-INDEX: 111; LEFT: 164px; POSITION: absolute; TOP: 136px" 
                   runat="server" Width="308px" Font-Size="10pt"></asp:textbox>
      <asp:TextBox id="txtT_Tel_Home" style="Z-INDEX: 110; LEFT: 164px; POSITION: absolute; TOP: 104px" 
                   runat="server" Width="308px" Font-Size="10pt"></asp:TextBox>
      <asp:TextBox id="txtT_Name" style="Z-INDEX: 102; LEFT: 164px; POSITION: absolute; TOP: 72px" 
                   Width="308px" runat="server" Font-Size="10pt"></asp:TextBox>
      <asp:Label id="Label7" style="Z-INDEX: 109; LEFT: 76px; POSITION: absolute; TOP: 232px" 
                 runat="server" Width="81px" Height="24px">Password</asp:Label>
      <asp:Label id="Label6" style="Z-INDEX: 108; LEFT: 76px; POSITION: absolute; TOP: 200px" 
                 runat="server" Width="81px" Height="24px">Login name</asp:Label>
      <asp:Label id="Label5" style="Z-INDEX: 107; LEFT: 76px; POSITION: absolute; TOP: 168px" 
                 runat="server" Width="81px" Height="24px">Address</asp:Label>
      <asp:Label id="Label4" style="Z-INDEX: 106; LEFT: 76px; POSITION: absolute; TOP: 136px" 
                 runat="server" Width="81px" Height="24px">Cellphone</asp:Label>
      <asp:Label id="Label3" style="Z-INDEX: 105; LEFT: 76px; POSITION: absolute; TOP: 104px" 
                 runat="server" Width="81px" Height="24px">Telephone</asp:Label>
      <asp:Label id="Label2" style="Z-INDEX: 101; LEFT: 76px; POSITION: absolute; TOP: 72px" 
                 Width="81px" runat="server" Height="24px">Teacher name</asp:Label>
      <asp:Button id="btnTeacher_Insert" style="Z-INDEX: 103; LEFT: 400px; POSITION: absolute; TOP: 264px" 
                  Width="72px" runat="server" Height="24px" Text="Add" OnClick ="Teacher_Insert"></asp:Button>
      <asp:Label id="Label1" style="Z-INDEX: 104; LEFT: 208px; POSITION: absolute; TOP: 12px" 
                 Width="136px" runat="server" Font-Bold="True" Height="36px" Font-Size="X-Large">Teacher</asp:Label>
    </form>
  </BODY>
</HTML>
