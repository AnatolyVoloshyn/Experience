unit unIni;

interface
uses IniFiles, Forms, SysUtils, System.Classes;

function uReadStringIniFiles(const Section:string; const Ident:string; const default_value:string=''):string;
procedure uSaveStringIniFiles(const Section:string; const Ident:string; const Value:String);

function uReadIntegerIniFiles(const Section:string; const Ident:string; const default_value:integer=0):integer;
procedure uSaveIntegerIniFiles(const Section:string; const Ident:string; const Value:Integer);

function uReadFloatIniFiles(const Section:string; const Ident:string; const default_value:Double=0):Double;
procedure uSaveFloatIniFiles(const Section:string; const Ident:string; const Value:Double);

function uReadBooleanIniFiles(const Section:string; const Ident:string; const default_value:Boolean=False):Boolean;
procedure uSaveBooleanIniFiles(const Section:string; const Ident:string; const Value:Boolean);

function uReadDateTimeIniFiles(const Section:string; const Ident:string; const default_value:TDateTime = 0):TDateTime;
procedure uSaveDateTimeIniFiles(const Section:string; const Ident:string; const Value:TDateTime);

procedure uEraseSection(const Section: string);
function uexistident(const Section:string;const Ident:string):boolean;


function uReadBinaryStreamIniFiles(const Section:string; const Ident:string; var s_value:TMemoryStream):Integer;
procedure uSaveBinaryStreamIniFiles(const Section:string; const Ident:string; Value:TMemoryStream);






//Для коннекта
function uConnectReadStringIniFiles(const Section:string; const Ident:string; const default_value:string=''):string;
procedure uConnectSaveStringIniFiles(const Section:string; const Ident:string; const Value:String);
function uConnectReadBooleanIniFiles(const Section:string; const Ident:string; const default_value:boolean=False):boolean;
procedure uConnectSaveBooleanIniFiles(const Section:string; const Ident:string; const Value:boolean);
//Для коннекта



implementation
var
IniFile:TMemIniFile;
iniFileName:String;

connectIniFile:TIniFile;
connectiniFileName:String;


function uReadStringIniFiles(const Section:string; const Ident:string; const default_value:string):string;
begin
 Result:=IniFile.ReadString(Section,Ident,default_value);
end;

procedure uSaveStringIniFiles(const Section:string; const Ident:string; const Value:String);
begin
 IniFile.WriteString(Section,Ident,Value);
end;

function uReadIntegerIniFiles(const Section:string; const Ident:string; const default_value:integer=0):integer;
begin
 Result:=IniFile.ReadInteger(Section,Ident,default_value);
end;

procedure uSaveIntegerIniFiles(const Section:string; const Ident:string; const Value:Integer);
begin
 IniFile.WriteInteger(Section,Ident,Value);
end;


function uReadFloatIniFiles(const Section:string; const Ident:string; const default_value:Double=0):Double;
begin
  Result:=IniFile.ReadFloat(Section,Ident,default_value);
end;


procedure uSaveFloatIniFiles(const Section:string; const Ident:string; const Value:Double);
begin
 IniFile.WriteFloat(Section,Ident,Value);
end;


function uReadBooleanIniFiles(const Section:string; const Ident:string; const default_value:Boolean=False):Boolean;
begin
 Result:=IniFile.ReadBool(Section,Ident,default_value);
end;


procedure uSaveBooleanIniFiles(const Section:string; const Ident:string; const Value:Boolean);
begin
 IniFile.WriteBool(Section,Ident,Value);
end;



function uReadDateTimeIniFiles(const Section:string; const Ident:string; const default_value:TDateTime = 0):TDateTime;
begin
   Result:=IniFile.ReadDateTime(Section,Ident,default_value);
end;



procedure uSaveDateTimeIniFiles(const Section:string; const Ident:string; const Value:TDateTime);
begin
  IniFile.WriteDateTime(Section,Ident,Value);
end;


function uexistident(const Section:string;const Ident:string):boolean;
begin
  Result:=IniFile.ValueExists(Section,Ident);
end;



procedure uEraseSection(const Section: string);
begin
  IniFile.EraseSection(Section);
end;


function uReadBinaryStreamIniFiles(const Section:string; const Ident:string; var s_value:TMemoryStream):Integer;
var
  Text: string;
  Stream: TMemoryStream;
  Pos: Integer;
  dataLen: Integer;
  DataBytes: TBytes;
begin
  Text := uReadStringIniFiles(Section, Ident, '');
  if Text <> '' then
  begin
    if s_value is TMemoryStream then
      Stream := TMemoryStream(s_value)
    else
      Stream := TMemoryStream.Create;

    try
      dataLen := Length(Text) div 2;
      SetLength(DataBytes, dataLen);
      Pos := Stream.Position;
      HexToBin(BytesOf(Text), 0, DataBytes, 0, dataLen);
      Stream.Write(DataBytes[0], dataLen);
      Stream.Position := Pos;
      if s_value <> Stream then
        s_value.CopyFrom(Stream, dataLen);
      Result := Stream.Size - Pos;
    finally
      if s_value <> Stream then
        Stream.Free;
    end;
  end
  else
    Result := 0;
end;


procedure uSaveBinaryStreamIniFiles(const Section:string; const Ident:string; Value:TMemoryStream);
var
  Text: string;
  Stream: TBytesStream;
  Buffer: TBytes;
begin
  SetLength(Text, (Value.Size - Value.Position) * 2);
  if Text.Length > 0 then
  begin
    if Value is TBytesStream then
      Stream := TBytesStream(Value)
    else
      Stream := TBytesStream.Create;

    try
      if Stream <> Value then
      begin
        Stream.CopyFrom(Value, Value.Size - Value.Position);
        Stream.Position := 0;
      end;
      SetLength(Buffer, Stream.Size * 2);
      BinToHex(Stream.Bytes, Stream.Position, Buffer, 0, Stream.Size - Stream.Position);
      Text := StringOf(Buffer);
    finally
      if Value <> Stream then
        Stream.Free;
    end;
  end;
  uSaveStringIniFiles(Section, Ident, Text);
end;






//Для коннекта
function uConnectReadStringIniFiles(const Section:string; const Ident:string; const default_value:string):string;
begin
 Result:=ConnectIniFile.ReadString(Section,Ident,default_value);
end;
procedure uConnectSaveStringIniFiles(const Section:string; const Ident:string; const Value:String);
begin
 ConnectIniFile.WriteString(Section,Ident,Value);
end;
function uConnectReadBooleanIniFiles(const Section:string; const Ident:string; const default_value:boolean=False):boolean;
begin
  Result:=ConnectIniFile.ReadBool(Section,Ident,default_value);
end;
procedure uConnectSaveBooleanIniFiles(const Section:string; const Ident:string; const Value:boolean);
begin
 ConnectIniFile.WriteBool(Section,Ident,Value);
end;
//Для коннекта

initialization
 iniFileName:=Application.ExeName;
 iniFileName[Length(iniFileName)-2]:='i';
 iniFileName[Length(iniFileName)-1]:='n';
 iniFileName[Length(iniFileName)-0]:='i';
 IniFile:=TMemIniFile.Create(iniFileName);

 connectiniFileName:=ExtractFileDir(Application.ExeName)+'\connect.ini';
 connectiniFile:=TIniFile.Create(connectiniFileName);

finalization
 IniFile.UpdateFile;
 IniFile.Free;
 connectiniFile.Free;
end.
