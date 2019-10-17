unit UnitPrincipal;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ComCtrls, StdCtrls, Buttons, IdTCPConnection, IdTCPClient,
  IdMessageClient, IdSMTP, IdComponent, IdIOHandler, IdIOHandlerSocket,
  IdSSLOpenSSL, IdBaseComponent, IdMessage;

type
  TForm1 = class(TForm)
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    Edit1: TEdit;
    Edit2: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Edit3: TEdit;
    Label3: TLabel;
    Label4: TLabel;
    Edit4: TEdit;
    Label5: TLabel;
    Edit5: TEdit;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    Label6: TLabel;
    Edit6: TEdit;
    Label7: TLabel;
    Edit7: TEdit;
    Label8: TLabel;
    Edit8: TEdit;
    ComboBox1: TComboBox;
    OpenDialog1: TOpenDialog;
    StatusBar1: TStatusBar;
    IdMessage1: TIdMessage;
    IdSSLIOHandlerSocket1: TIdSSLIOHandlerSocket;
    IdSMTP1: TIdSMTP;
    Memo1: TMemo;
    ListBox1: TListBox;
    Label9: TLabel;
    Edit9: TEdit;
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.SpeedButton1Click(Sender: TObject);
var
  i: integer;
  IdMessage: TIdMessage;
begin
  if OpenDialog1.Execute then
    begin
      for i:=0 to OpenDialog1.Files.Count -1 do
        if (ListBox1.Items.IndexOf(OpenDialog1.Files[i])= -1) then
          ListBox1.Items.Add(OpenDialog1.Files[i]);
    end;

end;

procedure TForm1.SpeedButton2Click(Sender: TObject);
var
// objetos necessários para o funcionamento
  IdSSLTOHandlerSocket: TIdSSLIOHandlerSocket;
  IdSMTP:TIdSMTP;
  IdMessage: TIdMessage;
  CaminhoAnexo: string;
  i:Integer;
begin
// instanciação dos objetos
  IdSSLTOHandlerSocket := TIdSSLIOHandlerSocket.Create(Self);
  IdSMTP := TIdSMTP.Create(Self);
  IdMessage := TIdMessage.Create(Self);
  try
// configuração do SSL
  IdSSLTOHandlerSocket.SSLOptions.Method := sslvSSLv23;
  IdSSLTOHandlerSocket.SSLOptions.Mode := sslmClient;
// configuração do SMTP
  IdSMTP.IOHandler := IdSSLTOHandlerSocket;
  IdSMTP.AuthenticationType := atLogin;
  IdSMTP.Port := StrToInt(ComboBox1.Text);
  IdSMTP.Host := Edit6.Text;
  IdSMTP.Username := Edit7.Text;
  IdSMTP.Password := Edit8.Text;
// tentativa de conexão e autenticação
  try
    IdSMTP.Connect;
    IdSMTP.Authenticate;
  except
      on E:Exception do
      begin
        MessageDlg('Erro na conexão e/ou autenticação: '
                    + E.Message, mtWarning, [mbOK], 0);
        Exit;
      end;
   end;
// Configuração da mensagem
  IdMessage.From.Address := Edit2.Text;
  IdMessage.From.Name := Edit9.Text;
  IdMessage.ReplyTo.EMailAddresses := IdMessage.From.Address;
  IdMessage.Recipients.EMailAddresses := Edit3.Text;
  IdMessage.CCList.EMailAddresses := Edit4.Text;
  IdMessage.BccList.EMailAddresses := Edit5.Text;
  IdMessage.Subject := Edit1.Text;
  IdMessage.Body.Text := Memo1.Lines.Text;
// anexo da mensagem (opcional)
  if  ListBox1.Items.Count > 0 then
    begin
      for i:= 0 to ListBox1.Items.Count -1 do
        begin
          if FileExists(ListBox1.Items[i]) then
          TIdAttachment.Create(IdMessage.MessageParts, ListBox1.Items[i]);
        end;
    end;
// envio da mensagem
    try
      IdSMTP.Send(IdMessage);
      MessageDlg('Mensagem enviada com sucesso.', mtInformation, [mbOK], 0);
    except
      On E:Exception do
        MessageDlg('Erro ao enviar a mensagem: '
                    + E.Message, mtWarning, [mbOK], 0);
    end;
  finally
// liberação dos objetos da memória
  FreeAndNil(IdMessage);
  FreeAndNil(IdSSLTOHandlerSocket);
  FreeAndNil(IdSMTP);
  end;
end;

end.
