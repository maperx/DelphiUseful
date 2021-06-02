// --- Copyright © 2020 Matias A. P.
// --- All rights reserved
// --- maperx@gmail.com

unit uConfigJson;

interface

uses JsonDataObjects, Variants, WinApi.Windows, System.SysUtils, System.Classes, uCipher;

type
  TConfigJson = class;

  /// <remarks><remarks>
  TConfigElement = class
  private
    FConfigDB: TConfigJson;
    FElement: TJSONObject;
    FArrayID, FElementID: string;
    function GetCount: integer;
    function GetValueKey(const Index: integer): string;
  public
    property ArrayID: string read FArrayID;
    property ElementID: string read FElementID;
    property Count: integer read GetCount;
    property ValueKeys[const Index: integer]: string read GetValueKey;
    constructor Create(var ConfigDB: TConfigJson; const ArrayID, ElementID: string); overload;
    constructor Create(var ConfigDB: TConfigJson; const ArrayID: string; Element: TJSONObject); overload;
    function WriteValue(const ValueKey: string; const value: Variant): boolean;
    function WriteValueArray(const ValueKey: string; const value: array of Variant): boolean;
    function ReadValue(const ValueKey: string; out value: Variant): boolean;
    function ReadString(const ValueKey: string): string;
    function ReadValueArray(const ValueKey: string; out value: Variant): boolean;
    function ReadBool(const ValueKey: string; Def: boolean = False): boolean;
    function ReadInt64(const ValueKey: string; Default: int64 = 0): int64;
    function GetJSON: string;
  end;

  TConfigArray = class
  private
    FConfigDB: TConfigJson;
    FArray: TJSONArray;
    FArrayID: string;
    function GetCount: integer;
    function GetElement(const Index: integer): TConfigElement;
  public
    property ArrayID: string read FArrayID;
    property Count: integer read GetCount;
    property Elements[const Index: integer]: TConfigElement read GetElement;
    constructor Create(var ConfigDB: TConfigJson; const ArrayID: string);
    function WriteValue(const ElementID, ValueKey: string; const value: Variant): boolean;
  end;

  TConfigJson = class
  private
    JO: TJSONObject;
    FFileName, FPass: string;
    FCompressed, FEncrypted: boolean;
    function _InsertElemValue(var Element: TJSONObject; ValueKey: string; const value: Variant): boolean;
    function _InsertElemValueArray(var Element: TJSONObject; ValueKey: string; const value: array of Variant): boolean;
    function GetArray(const ID: string): TConfigArray;
    function GetSaveStream: TBytesStream;
  public
    property FileName: string read FFileName;
    property Compressed: boolean read FCompressed;
    property Encrypted: boolean read FEncrypted;
    property Arrays[const ID: string]: TConfigArray read GetArray;
    constructor Create; overload;
    constructor Create(var Stream: TBytesStream; const Compressed: boolean = False; const Encrypted: boolean = False; const Pass: string = ''); overload;
    constructor Create(const JSON: string; const Compressed: boolean = False; const Encrypted: boolean = False; const Pass: string = ''); overload;
    function GetJSON(Compact: boolean = True): string;
    function WriteValue(const ArrayID, ElementID, ValueKey: string; const value: Variant; Flush: boolean = False): boolean;
    function WriteValueArray(const ArrayID, ElementID, ValueKey: string; const value: array of Variant): boolean;
    function WriteString(const ArrayID, ElementID, ValueKey: string; const value: string; Cifrar: boolean = False): boolean;
    function ReadValue(const ArrayID, ElementID, ValueKey: string; out value: Variant): boolean;
    function ReadString(const ArrayID, ElementID, ValueKey: string; Cifrada: boolean = False): string;
    function ReadValueArray(const ArrayID, ElementID, ValueKey: string; out value: Variant): boolean;
    function ReadBool(const ArrayID, ElementID, ValueKey: string; Def: boolean = False): boolean;
    function ReadInt64(const ArrayID, ElementID, ValueKey: string; Default: int64 = 0): int64;
    function DeleteArray(const ArrayID: string): boolean;
    function DeleteElement(const ArrayID, ElementID: string): boolean;
    function DeleteValue(const ArrayID, ElementID, ValueKey: string): boolean;
    function Save(const FullPath: string = ''): boolean; // ; Compact: boolean = True
    function SaveToStream(var Stream: TBytesStream): boolean;
    procedure SetEncrypted(const Encrypted: boolean; const Pass: string = '');
    procedure SetCompressed(const Compressed: boolean);
    function EstaVacio: boolean;
    function ExisteArray(const ArrayID: string): boolean;
  end;

implementation

uses System.Zip, REST.JSON, System.NetEncoding;

const
  _id = '_id';
  _pp = 'bWF0aWFzMTk4MQ';

type
  THeader = packed record
    Tag: string[3];
    Compressed: boolean;
    Encrypted: boolean;
  end;

constructor TConfigJson.Create;
begin
  inherited;
  FPass := EmptyStr;
  FEncrypted := False;
  FCompressed := False;
  FFileName := EmptyStr;
  JO := TJSONObject.Create;
end;

constructor TConfigJson.Create(const JSON: string; const Compressed: boolean = False; const Encrypted: boolean = False; const Pass: string = '');
var
  FS: TBytesStream;
  SS: TStringStream;
begin
  FS := TBytesStream.Create;
  try
    if FileExists(JSON) then begin
      FFileName := JSON;
      FS.LoadFromFile(FFileName);
    end
    else begin
      SS := TStringStream.Create;
      if JSON = '' then
        SS.WriteString('{}')
      else
        SS.WriteString(JSON);
      FS.LoadFromStream(SS);
      SS.Free;
    end;
    Create(FS, Compressed, Encrypted, Pass);
  finally
    FS.Free;
  end;
end;

constructor TConfigJson.Create(var Stream: TBytesStream; const Compressed: boolean = False; const Encrypted: boolean = False; const Pass: string = '');
var
  Cipher: TCipher;
  SS: TBytesStream;
  c, s: int64;
  Zip: TZipFile;
  bytes: TBytes;
  header: THeader;
begin
  FPass := Pass;
  FEncrypted := Encrypted;
  FCompressed := Compressed;
  if (Stream = nil) then
    Exit;
  c := Stream.Size;
  s := sizeof(header);
  Stream.Position := 0;
  if c > s then
    Stream.Read(header, s);
  if (header.Tag <> 'CDB') then // si no tiene el header, es texto plano
  begin
    Stream.Position := 0;
    try
      JO := TJSONObject.ParseFromStream(Stream) as TJSONObject;
    except
      JO := TJSONObject.Create;
    end;
    Exit;
  end;
  // sacar el header
  SetLength(bytes, c);
  Stream.Position := s;
  c := Stream.Read(bytes, c);
  Stream.Clear;
  Stream.Write(bytes, c);
  Stream.Position := 0;

  if header.Encrypted then begin
    FEncrypted := True;
    Cipher := TCipher.Create;
    SS := TBytesStream.Create;
    try
      Cipher.DecryptStream(Stream, SS, FPass, haSHA1);
      c := SS.Size;
      SS.Position := 0;
      Stream.Clear;
      Stream.CopyFrom(SS, c);
      Stream.Position := 0;
    finally
      Cipher.Free;
      SS.Free;
    end;
  end;

  if header.Compressed then begin
    FCompressed := True;
    Zip := TZipFile.Create;
    try
      Zip.Open(Stream, zmRead);
      Zip.Read(0, bytes);
      c := Length(bytes);
      Stream.Clear;
      Stream.Write(bytes, c);
      Stream.Position := 0;
    finally
      Zip.Free;
    end;
  end;

  try
    JO := TJSONObject.ParseFromStream(Stream) as TJSONObject;
  except
    JO := TJSONObject.Create;
  end;
end;

function TConfigJson.DeleteArray(const ArrayID: string): boolean;
var
  idx: integer;
begin
  result := False;
  idx := JO.IndexOf(ArrayID);
  if idx >= 0 then begin
    JO.Delete(idx);
    result := True;
  end;
end;

function TConfigJson.DeleteElement(const ArrayID, ElementID: string): boolean;
var
  i: integer;
begin
  result := False;
  if ExisteArray(ArrayID) then
    for i := 0 to JO.A[ArrayID].Count - 1 do
      if JO.A[ArrayID].O[i].s[_id] = ElementID then begin
        JO.A[ArrayID].Delete(i);
        Break;
      end;
end;

function TConfigJson.DeleteValue(const ArrayID, ElementID, ValueKey: string): boolean;
var
  i: integer;
begin
  result := False;
  if ExisteArray(ArrayID) then
    for i := 0 to JO.A[ArrayID].Count - 1 do
      if JO.A[ArrayID].O[i].s[_id] = ElementID then begin
        JO.A[ArrayID].O[i].Delete(JO.A[ArrayID].O[i].IndexOf(ValueKey));
        result := True;
        Break;
      end;
end;

function TConfigJson.GetArray(const ID: string): TConfigArray;
begin
  result := TConfigArray.Create(self, ID);
end;

function TConfigJson.GetJSON(Compact: boolean): string;
begin
  result := JO.ToJSON(Compact);
end;

function TConfigJson.EstaVacio: boolean;
begin
  result := (JO.Count = 0)
end;

function TConfigJson.ExisteArray(const ArrayID: string): boolean;
var
  idx: integer;
begin
  result := False;
  idx := JO.IndexOf(ArrayID);
  result := (idx >= 0) and (JO.Items[idx].Typ = jdtArray);
end;

function TConfigJson._InsertElemValue(var Element: TJSONObject; ValueKey: string; const value: Variant): boolean;
begin
  result := False;
  case VarType(value) of
    varInteger, varInt64, varByte, varWord:
      Element.L[ValueKey] := value;
    varSingle, varDouble, varCurrency:
      Element.F[ValueKey] := value;
    varBoolean:
      Element.B[ValueKey] := value;
  else
    Element.s[ValueKey] := value;
  end;
  result := True;
end;

function TConfigJson._InsertElemValueArray(var Element: TJSONObject; ValueKey: string; const value: array of Variant): boolean;
var
  i: integer;
begin
  result := False;
  if High(value) < 0 then
    Exit;
  Element.A[ValueKey].Clear;
  for i := Low(value) to High(value) do
    Element.A[ValueKey].Add(value[i]);
  result := True;
end;

function TConfigJson.WriteString(const ArrayID, ElementID, ValueKey, value: string; Cifrar: boolean): boolean;
var
  str: string;
  ciph: TCipher;
begin
  if Cifrar then begin
    ciph := TCipher.Create;
    try
      str := ciph.EncryptString(value, _pp, THashAlgorithm.haSHA1);
      str := TBase64Encoding.Base64.Encode(str);
      result := WriteValue(ArrayID, ElementID, ValueKey, str);
    finally
      ciph.Free;
    end;
  end
  else
    result := WriteValue(ArrayID, ElementID, ValueKey, value);
end;

function TConfigJson.WriteValue(const ArrayID, ElementID, ValueKey: string; const value: Variant; Flush: boolean = False): boolean;
var
  elem: TJSONObject;
  i: integer;
begin
  result := False;
  if ElementID = '' then
    Exit;
  elem := nil;
  for i := 0 to JO.A[ArrayID].Count - 1 do // buscar elemento de la coleccion que tenga _id=ElementID
  begin
    if JO.A[ArrayID].O[i].s[_id] = ElementID then begin
      elem := JO.A[ArrayID].O[i];
      Break;
    end;
  end;
  if elem = nil then // si no existe el elemento agregarlo
    elem := JO.A[ArrayID].AddObject;
  elem.s[_id] := ElementID;
  if VarType(value) = varArray then
    result := _InsertElemValueArray(elem, ValueKey, value)
  else
    result := _InsertElemValue(elem, ValueKey, value);
  if Flush and result and (FFileName <> EmptyStr) then
    Save(FFileName);
end;

function TConfigJson.WriteValueArray(const ArrayID, ElementID, ValueKey: string; const value: array of Variant): boolean;
var
  arr: TJSONArray;
  elem: TJSONObject;
  i: integer;
begin
  elem := nil;
  result := False;
  arr := JO.A[ArrayID]; // buscar coleccion sino agregarla
  for i := 0 to arr.Count - 1 do // buscar elemento de la coleccion que tenga _id=ElementID
  begin
    if JO.A[ArrayID].O[i].s[_id] = ElementID then begin
      elem := arr.Items[i].ObjectValue;
      Break;
    end;
  end;
  if elem = nil then // si no existe el elemento agregarlo
    elem := JO.A[ArrayID].AddObject;
  elem.s[_id] := ElementID;
  result := _InsertElemValueArray(elem, ValueKey, value);
end;

function TConfigJson.ReadValue(const ArrayID, ElementID, ValueKey: string; out value: Variant): boolean;
var
  i: integer;
begin
  result := False;
  try
    if ExisteArray(ArrayID) then
      for i := 0 to JO.A[ArrayID].Count - 1 do begin
        if JO.A[ArrayID].O[i].s[_id] = ElementID then begin
          if not JO.A[ArrayID].O[i].Values[ValueKey].IsNull then begin
            value := JO.A[ArrayID].O[i].Values[ValueKey];
            result := True;
            Break;
          end;
        end;
      end;
  except

  end;
end;

function TConfigJson.ReadString(const ArrayID, ElementID, ValueKey: string; Cifrada: boolean = False): string;
var
  ciph: TCipher;
  v: Variant;
begin
  result := EmptyStr;
  if ReadValue(ArrayID, ElementID, ValueKey, v) then
    result := v;

  if Cifrada then begin
    ciph := TCipher.Create;
    try
      result := TBase64Encoding.Base64.Decode(result);
      result := ciph.DecryptString(result, _pp, THashAlgorithm.haSHA1);
    finally
      ciph.Free;
    end;
  end;
end;

function TConfigJson.ReadValueArray(const ArrayID, ElementID, ValueKey: string; out value: Variant): boolean;
var
  i, j: integer;
  arr: TJSONArray;
begin
  result := False;
  try
    if ExisteArray(ArrayID) then
      for i := 0 to JO.A[ArrayID].Count - 1 do begin
        if JO.A[ArrayID].O[i].s[_id] = ElementID then begin
          arr := JO.A[ArrayID].O[i].Values[ValueKey].ArrayValue;
          value := VarArrayCreate([0, arr.Count - 1], varVariant);
          for j := 0 to arr.Count - 1 do
            value[j] := arr.Values[j];
          result := True;
        end;
        Break;
      end;
  except

  end;
end;

function TConfigJson.ReadBool(const ArrayID, ElementID, ValueKey: string; Def: boolean = False): boolean;
var
  value: Variant;
begin
  result := Def;
  if ReadValue(ArrayID, ElementID, ValueKey, value) then
    result := value;
end;

function TConfigJson.ReadInt64(const ArrayID, ElementID, ValueKey: string; Default: int64 = 0): int64;
var
  value: Variant;
begin
  result := Default;
  if ReadValue(ArrayID, ElementID, ValueKey, value) then
    result := value;
end;

function TConfigJson.GetSaveStream: TBytesStream;
var
  Cipher: TCipher;
  Zip: TZipFile;
  FS: TStringStream;
  TS, ZipStream: TBytesStream;
  header: THeader;
  c: int64;
begin
  result := TBytesStream.Create();
  FS := TStringStream.Create;
  FS.WriteString(GetJSON);
  FS.Position := 0;
  try
    if (not FCompressed) and (not FEncrypted) then begin
      result.CopyFrom(FS, 0);
      Exit;
    end;

    header.Tag := 'CDB';
    header.Compressed := FCompressed;
    header.Encrypted := FEncrypted;
    if FCompressed then begin
      ZipStream := TBytesStream.Create;
      Zip := TZipFile.Create;
      try
        Zip.Open(ZipStream, zmWrite);
        Zip.Add(FS, ExtractFileName(FFileName));
        Zip.Close;
      finally
        Zip.Free;
      end;
      ZipStream.Position := 0;
    end;
    TS := TBytesStream.Create;
    TS.Write(header, sizeof(header));
    try
      if FEncrypted then begin
        Cipher := TCipher.Create;
        if FCompressed then
          Cipher.EncryptStream(ZipStream, TS, FPass, haSHA1)
        else
          Cipher.EncryptStream(FS, TS, FPass, haSHA1);
        result.CopyFrom(TS, 0);
        Cipher.Free;
      end
      else if FCompressed then begin
        c := ZipStream.Size;
        ZipStream.Position := 0;
        TS.Write(ZipStream.bytes, c);
        result.CopyFrom(TS, 0);
      end;
    finally
      TS.Free;
      if FCompressed then
        ZipStream.Free;
    end;
  finally
    FS.Free;
  end;
end;

function TConfigJson.Save(const FullPath: string = ''): boolean;
var
  Stream: TBytesStream;
begin
  result := False;
  if FullPath <> '' then
    FFileName := FullPath;
  if (FFileName = '') then
    Exit;
  Stream := GetSaveStream;
  Stream.SaveToFile(FFileName);
  result := True;
{
  FS := TStringStream.Create;
  FS.WriteString(GetJSON(Compact));
  FS.Position := 0;
  try
    if (not FCompressed) and (not FEncrypted) then begin
      if (FFileName <> '') then
        FS.SaveToFile(FFileName);
      result := True;
      Exit;
    end;

    header.Tag := 'CDB';
    header.Compressed := FCompressed;
    header.Encrypted := FEncrypted;
    if FCompressed then begin
      ZipStream := TBytesStream.Create;
      Zip := TZipFile.Create;
      try
        Zip.Open(ZipStream, zmWrite);
        Zip.Add(FS, ExtractFileName(FFileName));
        Zip.Close;
      finally
        Zip.Free;
      end;
      ZipStream.Position := 0;
    end;
    TS := TBytesStream.Create;
    TS.Write(header, sizeof(header));
    try
      if FEncrypted then begin
        Cipher := TCipher.Create;
        if FCompressed then
          Cipher.EncryptStream(ZipStream, TS, FPass, haSHA1)
        else
          Cipher.EncryptStream(FS, TS, FPass, haSHA1);
        if (FFileName <> '') then
          TS.SaveToFile(FFileName);
        Cipher.Free;
      end
      else if FCompressed then begin
        c := ZipStream.Size;
        ZipStream.Position := 0;
        TS.Write(ZipStream.bytes, c);
        if (FFileName <> '') then
          TS.SaveToFile(FFileName);
      end;
    finally
      TS.Free;
      if FCompressed then
        ZipStream.Free;
    end;
  finally
    FS.Free;
  end; }
  result := True;
end;

function TConfigJson.SaveToStream(var Stream: TBytesStream): boolean;
begin
  result := False;
  Stream := GetSaveStream;
  result := (Stream.Size > 0);
end;

procedure TConfigJson.SetCompressed(const Compressed: boolean);
begin
  FCompressed := Compressed;
end;

procedure TConfigJson.SetEncrypted(const Encrypted: boolean; const Pass: string);
begin
  FEncrypted := Encrypted;
  if Pass <> '' then
    FPass := Pass;
end;

{ TConfigElement }

constructor TConfigElement.Create(var ConfigDB: TConfigJson; const ArrayID, ElementID: string);
var
  arr: TJSONArray;
  i: integer;
begin
  FConfigDB := ConfigDB;
  FArrayID := ArrayID;
  FElementID := ElementID;
  FElement := nil;

  arr := FConfigDB.JO.A[ArrayID]; // buscar array sino agregarlo

  for i := 0 to arr.Count - 1 do // buscar elemento de la array que tenga _id=ElementID
  begin
    if FConfigDB.JO.A[ArrayID].O[i].s[_id] = ElementID then begin
      FElement := FConfigDB.JO.A[ArrayID].O[i];
      Break;
    end;
  end;

  if FElement = nil then // si no existe el elemento agregarlo
  begin
    FElement := FConfigDB.JO.A[ArrayID].AddObject;
    FElement.s[_id] := ElementID;
  end;
end;

constructor TConfigElement.Create(var ConfigDB: TConfigJson; const ArrayID: string; Element: TJSONObject);
begin
  FConfigDB := ConfigDB;
  FArrayID := ArrayID;
  FElement := Element;
  FElementID := Element.s[_id];
end;

function TConfigElement.GetCount: integer;
begin
  result := FElement.Count;
end;

function TConfigElement.GetJSON: string;
begin
  result := FElement.ToJSON;
end;

function TConfigElement.GetValueKey(const Index: integer): string;
begin
  result := FElement.Items[Index].value;
end;

function TConfigElement.WriteValue(const ValueKey: string; const value: Variant): boolean;
begin
  result := FConfigDB._InsertElemValue(FElement, ValueKey, value);
end;

function TConfigElement.WriteValueArray(const ValueKey: string; const value: array of Variant): boolean;
begin
  result := FConfigDB._InsertElemValueArray(FElement, ValueKey, value);
end;

function TConfigElement.ReadBool(const ValueKey: string; Def: boolean = False): boolean;
var
  value: Variant;
begin
  result := False;
  if ReadValue(ValueKey, value) then
    result := value;
end;

function TConfigElement.ReadInt64(const ValueKey: string; Default: int64 = 0): int64;
var
  value: Variant;
begin
  result := Default;
  if ReadValue(ValueKey, value) then
    result := value;
end;

function TConfigElement.ReadValue(const ValueKey: string; out value: Variant): boolean;
begin
  result := False;
  if not FElement.Contains(ValueKey) then
    Exit;
  value := FElement.Values[ValueKey].VariantValue;
  result := True;
end;

function TConfigElement.ReadString(const ValueKey: string): string;
var
  v: Variant;
begin
  result := '';
  if ReadValue(ValueKey, v) then
    result := v;
end;

function TConfigElement.ReadValueArray(const ValueKey: string; out value: Variant): boolean;
var
  arr: TJSONArray;
  i: integer;
begin
  result := False;
  if not FElement.Contains(ValueKey) then
    Exit;

  if FElement.Values[ValueKey].Typ = jdtArray then begin
    arr := FElement.Values[ValueKey].ArrayValue;
    value := VarArrayCreate([0, arr.Count - 1], varVariant);
    for i := 0 to arr.Count - 1 do
      value[i] := (arr.Items[i].value);
    result := True;
  end;
end;

{ TConfigArray }

constructor TConfigArray.Create(var ConfigDB: TConfigJson; const ArrayID: string);
begin
  FConfigDB := ConfigDB;
  FArrayID := ArrayID;
  FArray := FConfigDB.JO.A[ArrayID]; // buscar array sino agregarlo
end;

function TConfigArray.GetCount: integer;
begin
  result := FArray.Count;
end;

function TConfigArray.GetElement(const Index: integer): TConfigElement;
begin
  if (Index >= 0) and (Index < FArray.Count) then
    result := TConfigElement.Create(FConfigDB, FArrayID, FArray.O[Index]);
end;

function TConfigArray.WriteValue(const ElementID, ValueKey: string; const value: Variant): boolean;
begin
  result := FConfigDB.WriteValue(FArrayID, ElementID, ValueKey, value);
end;

end.
