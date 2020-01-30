<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="WebApplication3.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <link href='http://fonts.googleapis.com/css?family=Open+Sans+Condensed:300' rel='stylesheet' type='text/css'/>
        <style type="text/css">
        .form-style-8{
	        font-family: 'Open Sans Condensed', arial, sans;
	        width: 500px;
	        padding: 30px;
	        background: #FFFFFF;
	        margin: 50px auto;
	        box-shadow: 0px 0px 15px rgba(0, 0, 0, 0.22);
	        -moz-box-shadow: 0px 0px 15px rgba(0, 0, 0, 0.22);
	        -webkit-box-shadow:  0px 0px 15px rgba(0, 0, 0, 0.22);

        }
        .form-style-8 h2{
	        background: #4D4D4D;
	        text-transform: uppercase;
	        font-family: 'Open Sans Condensed', sans-serif;
	        color: #FFF;
	        font-size: 18px;
	        font-weight: 100;
	        padding: 20px;
	        margin: -30px -30px 30px -30px;
        }

	    .form-style-8 h4{
	        background: #4D4D4D;
	        text-transform: uppercase;
	        font-family: 'Open Sans Condensed', sans-serif;
	        color: #FFF;
	        font-size: 18px;
	        font-weight: 100;
	        padding: 20px;
	    }

        .form-style-8 input[type="text"],
        .form-style-8 input[type="date"],
        .form-style-8 input[type="datetime"],
        .form-style-8 input[type="email"],
        .form-style-8 input[type="number"],
        .form-style-8 input[type="search"],
        .form-style-8 input[type="time"],
        .form-style-8 input[type="url"],
        .form-style-8 input[type="password"],
        .form-style-8 textarea,
        .form-style-8 select {
            box-sizing: border-box;
            -webkit-box-sizing: border-box;
            -moz-box-sizing: border-box;
            outline: none;
            display: block;
            width: 100%;
            padding: 7px;
            border: none;
            border-bottom: 1px solid #ddd;
            background: transparent;
            margin-bottom: 10px;
            font: 16px Arial, Helvetica, sans-serif;
            height: 45px;
        }
        .form-style-8 textarea{
	        resize:none;
	        overflow: hidden;
        }
        .form-style-8 input[type="button"], 
        .form-style-8 input[type="submit"]{
	        -moz-box-shadow: inset 0px 1px 0px 0px #45D6D6;
	        -webkit-box-shadow: inset 0px 1px 0px 0px #45D6D6;
	        box-shadow: inset 0px 1px 0px 0px #45D6D6;
	        background-color: #2CBBBB;
	        border: 1px solid #27A0A0;
	        display: inline-block;
	        cursor: pointer;
	        color: #FFFFFF;
	        font-family: 'Open Sans Condensed', sans-serif;
	        font-size: 14px;
	        padding: 8px 18px;
	        text-decoration: none;
	        text-transform: uppercase;
        }
        .form-style-8 input[type="button"]:hover, 
        .form-style-8 input[type="submit"]:hover {
	        background:linear-gradient(to bottom, #34CACA 5%, #30C9C9 100%);
	        background-color:#34CACA;
        }
        </style>
</head>
<body>
    <div class="form-style-8">
      <h2>Convert your file to Excel</h2>
      <form runat="server">
          <%--<asp:TextBox ID="txtFilePath" runat="server" placeholder="i.e. c:\folder\file1.pdf" ></asp:TextBox>--%>
          <asp:FileUpload runat="server" ID="fileUpload" AllowMultiple="true" />
          <asp:Button Text="Convert file" runat="server" ID="btnConvert" OnClick="btnConvert_Click" />
        <%--<input type="button" value="Send Message" />--%>
		  
		  <h4 id="lblErrorFiles"><asp:Label runat="server">Error in Reading Below Files</asp:Label></h4>
          <div>
            <%--<h5><asp:Label runat="server" ID="lbl">No Errors to Show</asp:Label></h5>--%>
            <asp:GridView ID="gvErrorFiles" runat="server" Width ="100%" ViewStateMode="Enabled" AutoGenerateColumns="true" BorderStyle="Solid" Visible="true" ShowHeaderWhenEmpty ="true" EmptyDataText ="No Errors to Show">
                 <Columns>
                     <asp:BoundField DataField="FILENAME" HeaderText="File Name" />
                     <asp:BoundField DataField="ERROR" HeaderText="Error" />
                 </Columns>
                 <FooterStyle BackColor="#990000" Font-Bold="True" ForeColor="White"></FooterStyle>
                 <HeaderStyle BackColor="#990000" Font-Bold="True" ForeColor="White"></HeaderStyle>
            </asp:GridView>
          </div>
      </form>
    </div>
</body>
</html>
