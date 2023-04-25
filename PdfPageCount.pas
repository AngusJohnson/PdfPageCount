unit PdfPageCount;

(*******************************************************************************
* Author    :  Angus Johnson                                                   *
* Version   :  2.02                                                            *
* Date      :  26 April 2023                                                     *
* Website   :  http://www.angusj.com                                           *
* Copyright :  Angus Johnson 2010-2022                                         *
* License   :  http://www.boost.org/LICENSE_1_0.txt                            *
*******************************************************************************)

////////////////////////////////////////////////////////////////////////////////
// Summary of steps taken to parse a PDF doc for its page count :-              
////////////////////////////////////////////////////////////////////////////////
//1.  See if there's a 'Linearization dictionary' for easy parsing.
//    Mostly there isn't so ...
//2.  Locate 'startxref' at end of file
//3.  get 'xref' offset and go to xref table
//4.  depending on PDF version, the xref table may or may not be in a compressed
//    object. If it is (ie PDF ver 1.5+), then go to that object and decompress
//    the xref table into the buffer (and move the file pointer to the buffer)
//5.  parse the xref table and fill a list with object numbers and offsets
//6.  handle subsections within xref table.
//7.  read 'trailer' section at end of each xref
//8.  store 'Root' object number if found in 'trailer'
//9.  if 'Prev' xref found in 'trailer' - loop back to step 3
//10. locate Root in the object list
//11. locate 'Pages' object from Root
//12. get Count from Pages.

interface

uses
  Windows, SysUtils, Classes, AnsiStrings, ZLib;

const
  PDF_NO_ERROR              =  0;
  PDF_ERROR_UNDEFINED       = -1;
  PDF_ERROR_FILE_OPEN       = -2;
  PDF_ERROR_FILE_FORMAT     = -3;
  PDF_ERROR_ENCRYPTED_STRM  = -4;

(*******************************************************************************
* GetPageCount                                                                 *
*   Returns a negative value on error                                          *
*   Returned 'error' values (see error const above)                            *
*******************************************************************************)

function GetPageCount(const filename: string): integer;

implementation

type

  PPdfObj = ^TPdfObj;
  TPdfObj = record
    number,
    offset: integer;
    filePtr: PAnsiChar;
    stmObjNum: integer;
  end;

  TSortFunc = function (item1, item2: pointer): boolean;

  TPdfPageCounter = class
    private
      ms        : TMemoryStream;
      p         : PAnsiChar;
      pEnd      : PAnsiChar;
      pSaved    : PAnsiChar;
      PdfObjList: TList;
      bufferSize: integer;
      buffer    : PAnsiChar;
      function  FindStartXRef: boolean;
      procedure SkipBlankSpace;
      procedure DisposeBuffer;
      function  GetInt(out num: integer): boolean;
      function  GetUInt(out num: integer): boolean;
      function  GetFileID: ansistring;
      function  GetPassword(out pwd: ansistring; getOwner: Boolean): boolean;
      function  IsString(const str: ansistring): boolean;
      function  FindStrInDict(const str: ansistring): boolean;
      function  FindStartOfDict: boolean;
      function  FindEndOfDict: boolean;
      function  FindObject(objNum: integer): PPdfObj;
      function  GotoObject(objNum: integer): boolean;
      function  SeekObjectBackward(objNum, revNum: integer;
        maxDist: integer = 0): boolean;
      function  DecompressObjIntoBuffer(objNum, genNum: integer): boolean;
      procedure ReversePngFilter(rowSize: integer);
      function  GetLinearizedPageNum(out pageNum: integer): boolean;
      function  GetPageNumUsingCrossRefStream: integer;
    public
      ErrorFlag: integer;
      constructor Create;
      destructor Destroy; override;
      procedure Clear;
      function  GetPdfPageCount(const filename: string): integer;
  end;

//------------------------------------------------------------------------------
// Miscellaneous functions
//------------------------------------------------------------------------------

procedure QuickSortList(SortList: TPointerList;
  L, R: Integer; sortFunc: TSortFunc);
var
  I, J: Integer;
  P, T: Pointer;
begin
  repeat
    I := L;
    J := R;
    P := SortList[(L + R) shr 1];
    repeat
      while (SortList[I] <> P) and sortFunc(SortList[I], P) do Inc(I);
      while (SortList[J] <> P) and sortFunc(P, SortList[J])  do Dec(J);
      if I <= J then
      begin
        T := SortList[I];
        SortList[I] := SortList[J];
        SortList[J] := T;
        Inc(I);
        Dec(J);
      end;
    until I > J;
    if L < J then QuickSortList(SortList, L, J, sortFunc);
    L := I;
  until I >= R;
end;
//------------------------------------------------------------------------------

function ListSort(item1, item2: pointer): boolean;
var
  obj1: PPdfObj absolute item1;
  obj2: PPdfObj absolute item2;
begin
  Result := obj2.number > obj1.number;
end;
//------------------------------------------------------------------------------

function Paeth(a, b, c: byte): byte; inline;
var
  p,pa,pb,pc: byte;
begin
  //https://www.w3.org/TR/PNG-Filters.html
  //a = left, b = above, c = upper left
  p := (a + b - c) and $FF;
  pa := abs(p - a);
  pb := abs(p - b);
  pc := abs(p - c);
  if (pa <= pb) and (pa <= pc) then result := a
  else if pb <= pc then result := b
  else result := c;
end;
//------------------------------------------------------------------------------

function PaethC({a=0, b=0} c: byte): byte; inline;
begin
  if c > 170 then Result := 0
  else Result := c;
end;
//------------------------------------------------------------------------------

function PaethB({a=0} b: byte {c=0} ): byte; inline;
begin
  Result := b;
end;
//------------------------------------------------------------------------------

function Average(a, b: byte): byte; inline;
begin
  result := (a + b) shr 1;
end;
//------------------------------------------------------------------------------

function PAnsiCharToInt(buffer: PAnsiChar; byteCnt: integer): integer; inline;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to byteCnt -1 do
    Result := Result shl 8 + ord((buffer+i)^);
end;

//------------------------------------------------------------------------------
// TPdfPageCounter methods
//------------------------------------------------------------------------------

constructor TPdfPageCounter.Create;
begin
  PdfObjList:= TList.Create;
  ms        := TMemoryStream.Create;
  bufferSize:= 0;
  p         := nil;
  pSaved    := nil;
  buffer    := nil;
end;
//------------------------------------------------------------------------------

destructor TPdfPageCounter.Destroy;
begin
  Clear;
  PdfObjList.Free;
  ms.Free;
end;
//------------------------------------------------------------------------------

procedure TPdfPageCounter.Clear;
var
  i: integer;
begin
  p       := nil;
  pSaved  := nil;
  DisposeBuffer;
  for i := 0 to PdfObjList.Count -1 do
    Dispose(PPdfObj(PdfObjList[i]));
  PdfObjList.Clear;
  ms.Clear;
end;
//------------------------------------------------------------------------------

procedure TPdfPageCounter.DisposeBuffer;
begin
  if not Assigned(buffer) then Exit;
  FreeMem(buffer);
  buffer := nil;
  bufferSize := 0;
  p := pSaved;
  pEnd := PAnsiChar(ms.Memory) + ms.Size;
end;
//------------------------------------------------------------------------------

procedure SubFilter(p: PAnsiChar; rowSize: integer); inline;
var
  i: integer;
begin
  for i := 1 to rowSize -1 do
  begin
    inc(p);
    p^ := AnsiChar((ord(p^) + ord((p -1)^)) and $FF);
  end;
end;
//------------------------------------------------------------------------------

procedure UpFilter(p: PAnsiChar; rowSize: integer); inline;
var
  i: integer;
begin
  for i := 0 to rowSize -1 do
  begin
    p^ := AnsiChar((ord(p^) + ord((p - rowSize)^)) and $FF);
    inc(p);
  end;
end;
//------------------------------------------------------------------------------

procedure AvgFilter(p: PAnsiChar; rowSize: integer; topRow: Boolean); inline;
var
  i: integer;
begin
  if topRow then
  begin
    for i := 1 to rowSize -1 do
    begin
      inc(p);
      p^ := AnsiChar((ord(p^) + average(ord((p -1)^), 0)) and $FF);
    end;
  end else
  begin
    p^ := AnsiChar(ord(p^) + average(ord((p - rowSize)^), 0) and $FF);
    for i := 1 to rowSize -1 do
    begin
      inc(p);
      p^ := AnsiChar((ord(p^) +
        average(ord((p - 1)^), ord((p - rowSize)^))) and $FF);
    end;
  end;
end;
//------------------------------------------------------------------------------

procedure ReversePaeth(p: PAnsiChar; rowSize: integer; topRow: Boolean); inline;
var
  i: integer;
begin
  if topRow then
  begin
    for i := 1 to rowSize -1 do
    begin
      inc(p);
      p^ := AnsiChar((ord(p^) + PaethC(ord((p -1)^))) and $FF);
    end;
  end else
  begin
    p^ := AnsiChar((ord(p^) + PaethB(ord((p - rowSize)^))) and $FF);
    for i := 1 to rowSize -1 do
    begin
      inc(p);
      p^ := AnsiChar((ord(p^) +
        Paeth(ord((p -1)^), ord((p - rowSize)^), ord((p - rowSize -1)^))) and $FF);
    end;
  end;
end;
//------------------------------------------------------------------------------

procedure TPdfPageCounter.ReversePngFilter(rowSize: integer);
var
  topRow: Boolean;
  pb, pb2, bpEnd: PAnsiChar;
  filterType: AnsiChar;
begin
  topRow := true;
  pb := buffer;
  bpEnd := buffer + bufferSize;
  while pb < bpEnd do
  begin
    filterType := pb^;
    dec(bufferSize);
    dec(bpEnd);
    move((pb +1)^, pb^, bpEnd - pb);
    case filterType of
      #0: ;//no filtering used for this row
      #1: SubFilter(pb, rowSize);
      #2: if not topRow then UpFilter(pb, rowSize);
      #3: AvgFilter(pb, rowSize, topRow);
      #4: ReversePaeth(pb, rowSize, topRow);
    end;
    inc(pb, rowSize);
    topRow := false;
  end;
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.GetInt(out num: integer): boolean;
var
  tmpStr: string;
  isNeg: Boolean;
begin
  tmpStr := '';
  while p^ < #33 do inc(p); //skip leading CR,LF & SPC
  isNeg := p^ = '-';
  if isNeg then inc(p);
  while (p^ in ['0'..'9']) do
  begin
    tmpStr := tmpStr + Char(PAnsiChar(p)^);
    inc(p);
  end;
  result := tmpStr <> '';
  if not result then exit;
  num := strtoint(tmpStr);
  if isNeg then num := -num;
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.GetUInt(out num: integer): boolean;
var
  tmpStr: string;
begin
  tmpStr := '';
  while p^ < #33 do inc(p); //skip leading CR,LF & SPC
  while (p^ in ['0'..'9']) do
  begin
    tmpStr := tmpStr + Char(PAnsiChar(p)^);
    inc(p);
  end;
  result := tmpStr <> '';
  if not result then exit;
  num := strtoint(tmpStr);
end;
//------------------------------------------------------------------------------

function HexToByte(c: ansichar): byte;
begin
  if (c <= '9') then
  begin
    if (c <= '0') then Result := 0
    else Result := Ord(c) - 48;
  end
  else if (c <= 'F') then
  begin
    if (c <= 'A') then Result := 0
    else Result := Ord(c) - 65 + 10;
  end
  else if (c <= 'f') then
  begin
    if (c <= 'a') then Result := 0
    else Result := Ord(c) - 97 + 10;
  end else
    Result := 0;
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.GetFileID: ansistring;
var
  p2: PAnsiChar;
  len: integer;
begin
  Result := '';
  if not FindStrInDict('/ID') or
    (p^ <> '[') or ((p+1)^ <> '<') then Exit;
  Inc(p, 2);
  p2 := p +1;
  while (p2^ <> '>') do inc(p2);
  len := (p2-p);
  SetLength(Result, len);
  Move(p^, Result[1], len);
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.GetPassword(out pwd: ansistring;
  getOwner: Boolean): boolean;
var
  p2: PAnsiChar;
  dl: AnsiChar;
begin
  Result := false;
  if getOwner then
  begin
    if not FindStrInDict('/O') then Exit;
  end else
    if not FindStrInDict('/U') then Exit;

  SkipBlankSpace;
  if (p^ <> '(') and (p^ <> '<') then Exit;
  if (p^ = '(') then dl := ')' else dl := '>';
  Inc(p);
  p2 := p +1;
  while (p2^ <> dl) do inc(p2);
  SetLength(pwd, (p2-p));
  Move(p^, pwd[1], Length(pwd));
  Result := true;
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.IsString(const str: ansistring): boolean;
var
  len: integer;
begin
  len := length(str);
  result := CompareMem(p, PAnsiChar(str), len);
  if result then inc(p, len);
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.FindStrInDict(const str: ansistring): boolean;
var
  nestLvl: integer;
  str1: AnsiChar;
begin
  //nb: PDF 'dictionaries' start with '<<' and terminate with '>>'
  result := false;
  nestLvl := 0;
  str1 := str[1];
  while not result do
  begin
    while not (p^ in ['>','<',str1]) do inc(p);
    if (p^ = '<') then
    begin
      if (p+1)^ = '<' then begin inc(nestLvl); inc(p); end;
    end
    else if (p^ = '>') then
    begin
      if (p+1)^ = '>' then
      begin
        dec(nestLvl);
        inc(p);
        if nestLvl <= 0 then exit;
      end
    end else
    begin
      result := (nestLvl < 2) and IsString(str);
      if result then exit;
    end;
    inc(p);
  end;
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.FindStartOfDict: boolean;
begin
  while (p < pEnd) and not (p^ in ['>','<']) do inc(p);
  result := (p < pEnd) and (p^ = '<') and ((p +1)^ = '<');
  if Result then inc(p, 2);
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.FindEndOfDict: boolean;
var
  nestLvl: integer;
begin
  result := false;
  nestLvl := 1;
  while true do
  begin
    while (p < pEnd) and not (p^ in ['>','<']) do inc(p);
    if (p >= pEnd)  then Exit;
    if (p^ = '<') then
    begin
      if (p+1)^ = '<' then begin inc(nestLvl); inc(p); end;
    end
    else if (p+1)^ = '>' then
    begin
      dec(nestLvl);
      if nestLvl < 0 then
        result := false
      else if nestLvl = 0 then
      begin
        inc(p, 2);
        result := true;
        exit; //found end of Dictionary
      end;
      inc(p); //skips first '>'
    end;
    inc(p);
  end;
end;
//------------------------------------------------------------------------------

procedure TPdfPageCounter.SkipBlankSpace;
begin
  while (p < pEnd) and (p^ < #33) do inc(p);
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.FindObject(objNum: integer): PPdfObj;
var
  l,r,m, mv: integer;
begin
  //precondition: PdfObjList is sorted
  Result := nil;
  //binary search sorted list
  l := 0; m:= 0; r := PdfObjList.Count-1; mv := -1;
  while l <= r do
  begin
    m := (l+r) div 2;
    mv := PPdfObj(PdfObjList[m]).number;
    if mv = objNum then break
    else if mv > objNum then r := m -1
    else l := m +1;
  end;
  if (mv = objNum) then
    Result := PPdfObj(PdfObjList[m]);
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.GotoObject(objNum: integer): boolean;
var
  i,j,k, N, FirstOffset: integer;
  //genNum: integer;
  streamObj: PPdfObj;
begin
  Result := false;
  streamObj := FindObject(objNum);
  if not Assigned(streamObj) then Exit;

  if Assigned(streamObj.filePtr) then
  begin
    p := streamObj.filePtr;
    result := GetUInt(i) and (i = objNum);
    Exit;
  end;

  //the object must be in a compressed stream
  if streamObj.stmObjNum < 0 then Exit;

  DisposeBuffer;
  if not GotoObject(streamObj.stmObjNum) then exit;
  pSaved := p;

  if not FindStrInDict('/Type') then exit;
  SkipBlankSpace;
  if not IsString('/ObjStm') then exit;

  p := pSaved;
  if not FindStrInDict('/N') then exit;
  //N = number of compressed objects in the stream ...
  if not GetUInt(N) then exit;

  p := pSaved;
  if not FindStrInDict('/First') or
    not GetUInt(FirstOffset) or
    not DecompressObjIntoBuffer(objNum, 0) then
      Exit;

  //NB: P IS NOW POINTING TO THE BUFFER BASE

  for i := 0 to N -1 do
  begin
    if not GetUInt(j) then exit; //object number
    if j = objNum then break;
    if not GetUInt(k) then exit;
  end;
  if j <> objNum then Exit;
  if not GetUInt(k) then exit; //byte offset relative to FirstOffset
  p := buffer + k + FirstOffset;
  Result := true;
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.SeekObjectBackward(objNum, revNum: integer;
  maxDist: integer): boolean;
var
  objStr: ansistring;
begin
  if (maxDist <= 0) or (maxDist > p - ms.Memory) then
    maxDist := p - ms.Memory;

  objStr := ansistring(Format('%d %d obj', [objNum, RevNum]));
  while (maxDist <> 0) do
  begin
    if (p^ = objStr[1]) and IsString(objStr) then
    begin
      Result := true;
      Exit;
    end;
    dec(p);
    dec(maxDist);
  end;
  Result := false;
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.DecompressObjIntoBuffer(objNum, genNum: integer): boolean;
var
  i,j, len: integer;
  prot: Int32;
  filterColCnt, predictor: integer;
  fileId, owner, user: ansistring;
  pSaved2: PAnsiChar;
begin
  result := false;
  p := pSaved;
  if not FindStrInDict('/Filter') then exit;
  SkipBlankSpace;
  //check that this a compression type that we can handle ...
 //nb: /FlateDecode WITH SQUARE BRACKETS is used in Tracker's PDF software
  if not IsString('/FlateDecode') and not IsString('[/FlateDecode]') then exit;
  p := pSaved;
  if not FindStrInDict('/DecodeParms') or not
    FindStrInDict('/Columns') or not GetUInt(filterColCnt) then
      filterColCnt := 0; //j = column count (bytes per row)
  if filterColCnt > 0 then
  begin
    SkipBlankSpace;
    if not IsString('/Predictor') or not GetUInt(predictor) then
      predictor := 0;
  end;

  p := pSaved;
  fileId := GetFileID;

  p := pSaved;
  if not FindStrInDict('/Length') then exit;
  if not GetUInt(len) then exit;
  //caution: while len is usually the length of the compressed stream, it may
  //also be an indirect reference to an object containing the length ...
  if GetUInt(i) and IsString(' R') then
  begin
    if not GotoObject(len) or
      not GetUInt(i) or //skip the generation num
      not IsString(' obj') or
      not GetUInt(len) then exit; //OK, this is the stream length
  end;

  p := pSaved;
  if FindStrInDict('/Encrypt') then
  begin
    if fileId = '' then Exit; //required for encryption
    if GetUInt(i) then //indirect object
    begin
      if not GetUInt(j) then Exit;
      p := pSaved;
      if not SeekObjectBackward(i,j, $3FF) then Exit;
      pSaved2 := p;
    end else
      pSaved2 := pSaved;

    if not FindStrInDict('/P') or not GetInt(prot) then Exit;
    p := pSaved2;
    if not GetPassword(owner, true) then Exit;
    p := pSaved2;
    if not GetPassword(user, false) then Exit;
    ErrorFlag := PDF_ERROR_ENCRYPTED_STRM; ////////////////////
  end;

  p := pSaved;
  FindStartOfDict;
  if not FindEndOfDict then exit;
  while (p^ <> 's') do inc(p);
  if not IsString('stream') then exit;
  SkipBlankSpace;

  try
    //decompress the stream ...
    //nb: I'm not sure in which Delphi version these functions were renamed.
    {$IFDEF UNICODE}
    zlib.ZDecompress(p, len, pointer(buffer), bufferSize);
    {$ELSE}
    zlib.DecompressBuf(p, len, len*3, pointer(buffer), bufferSize);
    {$ENDIF}
  except
    ErrorFlag := PDF_ERROR_ENCRYPTED_STRM;
    buffer := nil;
    bufferSize := 0;
    Exit; //fails with any encryption
  end;

  //now de-filter the decompressed output (typically PNG filtering)
  //Filter Columns should match the byte count of /W[X Y Z]
  //The decompressed stream prefiltered size == (X Y Z) * entries
  //(ie allowing extra byte at the start of each column for filter type)
  //see also http://www.w3.org/TR/PNG-Filters.html
  if (filterColCnt > 0) and (predictor > 9) then
    ReversePngFilter(filterColCnt);

  p := buffer;
  pEnd := buffer + bufferSize;
  result := true;
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.FindStartXRef: boolean;
begin
  while p > ms.Memory do
  begin
    case p^ of
      'f': dec(p, 8); 'e': dec(p, 7); 'x': dec(p, 5);
      'r': dec(p, 3); 'a': dec(p, 2); 't': dec(p, 1);
      's':
        if AnsiStrings.StrLComp(p, 'startxref', 9) = 0 then
        begin
          result := true;
          inc(p, 9);
          Exit;
        end
        else dec(p, 9);
      else dec(p, 9);
    end;
  end;
  result := false;
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.GetLinearizedPageNum(out pageNum: integer): boolean;
var
  pStart,pStop: PAnsiChar;
begin
  pageNum := PDF_ERROR_UNDEFINED;
  result := false;
  pStop := p + 32;
  while (p < pStop) and (p^ <> 'o') do inc(p);
  if AnsiStrings.StrLComp( p, 'obj', 3) <> 0 then exit;
  pStart := p;
  if not FindStrInDict('/Linearized') then exit;
  p := pStart;
  if FindStrInDict('/N ') and GetUInt(pageNum) then result := true;
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.GetPageNumUsingCrossRefStream: integer;
var
  i, j, k, pagesNum, rootNum: integer;
  objNum, genNum: integer;
  indexArray: array of integer;
  buffPtr: PAnsiChar;
  w1,w2,w3: integer;
  PdfObj: PPdfObj;
begin
  //presumption: 'trailer' is in cross-reference stream.
  Result := PDF_ERROR_UNDEFINED;
  if not GetUInt(objNum) then exit; //stream obj number
  if not GetUInt(genNum) then exit; //stream obj revision number
  pSaved := p;
  if not FindStrInDict('/Type') then exit;
  SkipBlankSpace;
  if not IsString('/XRef') then exit;

  //todo - check for and manage /Prev too

  p := pSaved;
  if not FindStrInDict('/Root') then exit;
  SkipBlankSpace;
  if not GetUInt(rootNum) then exit;

  //get the stream cross-ref table field sizes ...
  p := pSaved;
  if not FindStrInDict('/W') then exit;
  SkipBlankSpace;
  if p^ <> '[' then exit;
  inc(p);
  if not GetUInt(w1) or (w1 <> 1) or not GetUInt(w2) or
    not GetUInt(w3) then exit;

  //Index [F1 N1, ..., Fn, Nn]. If absent assumes F1 = 0 & N based on size
  //(Fn: first object in table subsection; Nn: number in table subsection)
  indexArray := nil;
  p := pSaved;
  if FindStrInDict('/Index') then
  begin
    SkipBlankSpace;
    if p^ <> '[' then exit;
    inc(p);
    while GetUInt(i) and GetUInt(j) do
    begin
      k := length(indexArray);
      SetLength(indexArray, k +2);
      indexArray[k] := i;
      indexArray[k +1] := j;
    end;
  end;

  //todo - handle uncompressed streams too
  //assume all streams are compressed (though this is really optional)

  p := pSaved;
  if not DecompressObjIntoBuffer(objNum, genNum) or
    (bufferSize mod (w1 + w2 + w3) <> 0) then exit;

  //if the Index array is empty then use the default values ...
  if length(indexArray) = 0 then
  begin
    setLength(indexArray, 2);
    indexArray[0] := 0;
    indexArray[1] := bufferSize div (w1 + w2 + w3);
  end;

  buffPtr := buffer;
  //loop through each subsection in the table and
  //populate our object list ...
  for i := 0 to (length(indexArray) div 2) -1 do
  begin
    k := indexArray[i*2]; //k := base object number
    for j := 0 to indexArray[i*2 +1] -1 do
    begin
      case buffPtr^ of
        #0: //free object (ignore)
          inc(buffPtr, w1 + w2 + w3);
        #1: //uncompressed object
          begin
            inc(buffPtr, w1);
            new(PdfObj);
            PdfObjList.Add(PdfObj);
            PdfObj.number := k;
            PdfObj.stmObjNum := -1;
            PdfObj.offset := PAnsiCharToInt(buffPtr, w2);
            PdfObj.filePtr := PAnsiChar(ms.Memory) + PdfObj.offset;
            inc(buffPtr, w2 + w3);
          end;
        #2: //compressed object
          begin
            inc(buffPtr, w1);
            new(PdfObj);
            PdfObjList.Add(PdfObj);
            PdfObj.number := k;
            PdfObj.stmObjNum := PAnsiCharToInt(buffPtr, w2);
            inc(buffPtr,w2);
            PdfObj.offset := PAnsiCharToInt(buffPtr, w3);
            PdfObj.filePtr := nil;
            inc(buffPtr, w3);
          end;
        else
          Exit; //error
      end;
      inc(k);
    end;
  end;

  DisposeBuffer;
  if rootNum < 0 then exit;

  QuickSortList(PdfObjList.List, 0, PdfObjList.Count -1, ListSort);
  if not GotoObject(rootNum) then exit;
  if not FindStrInDict('/Pages') then exit;
  //get the Pages' object number, go to it and get the page count ...
  if not GetUInt(pagesNum) then exit;

  DisposeBuffer;
  if not GotoObject(pagesNum) or
    not FindStrInDict('/Count') or not GetUInt(k) then exit;
  //if we get this far the page number has been FOUND!!!
  result := k;
  exit;
end;
//------------------------------------------------------------------------------

function TPdfPageCounter.GetPdfPageCount(const filename: string): integer;
var
  k, cnt, pagesNum, rootNum: integer;
  PdfObj: PPdfObj;
begin
  ErrorFlag := PDF_NO_ERROR;
  Result := PDF_ERROR_UNDEFINED;
  try
    try
      ms.LoadFromFile(filename);
    except
      ErrorFlag := PDF_ERROR_FILE_OPEN;
      Exit;
    end;

    p := PAnsiChar(ms.Memory);
    pEnd := PAnsiChar(ms.Memory) + ms.Size;

    //for an easy life let's hope the file has a 'linearization dictionary'
    //at the beginning of the document ...
    if GetLinearizedPageNum(result) then exit;

    //find 'startxref' and the beginning of 'trailer dictionary'
    //ignoring '%%EOF' at end of file
    p := pEnd -5 - 9;
    if not FindStartXRef then
    begin
      ErrorFlag := PDF_ERROR_FILE_FORMAT;
      exit;
    end;

    rootNum := -1; //ie flag as not yet found

    if not GetUInt(k) or       //xref offset ==> k
      (k >= ms.size) then exit;

    p :=  PAnsiChar(ms.Memory) + k;
    if AnsiStrings.StrLComp(p, 'xref', 4) <> 0 then
    begin
      //'xref' table not found so assume the doc contains
      //a cross-reference stream instead (ie PDF doc ver 1.5+)
      Result := GetPageNumUsingCrossRefStream;
      Exit;
    end;

    //to get here this is most likely an older PDF doc
    //ie ver 1.4 or earlier

    inc(p,4);
    while true do //top of loop
    begin
      //get base object number ==> k
      if not GetUInt(k) then exit;
      //get object count ==> cnt
      if not GetUInt(cnt) then exit;
      //it is possible to have 0 objects in a section
      SkipBlankSpace;

      //add all objects in section to list ...
      for cnt := 0 to cnt-1 do
      begin
        new(PdfObj);
        PdfObjList.Add(PdfObj);
        PdfObj.number := k + cnt;
        if not GetUInt(PdfObj.offset) then exit;
        PdfObj.filePtr := PAnsiChar(ms.Memory) + PdfObj.offset;
        //while each entry SHOULD be exactly 20 bytes, not all PDF document
        //creators seems to adhere to this :( ...
        while not (p^ in [#10,#13]) do inc(p);
        while (p^ in [#10,#13]) do inc(p);
      end;
      //check for and process further subsections ...
      if p^ in ['0'..'9'] then continue;

      // parse 'trailer dictionary' ...
      if not IsString('trailer') then exit;
      pSaved := p;
      // get Root (aka Catalog) ...
      if (rootNum = -1) and FindStrInDict('/Root') then
        if not GetUInt(rootNum) then exit;
      p := pSaved;
      if not FindStrInDict('/Prev') then break; //no more xrefs

      //next xref offset ==> k
      if not GetUInt(k) then exit;
      p :=  PAnsiChar(ms.Memory) + k +4;

    end; //bottom of loop

    //Make sure we've got Root's object number ...
    if rootNum < 0 then exit;

    QuickSortList(PdfObjList.List, 0, PdfObjList.Count -1, ListSort);
    if not GotoObject(rootNum) then exit;

    if not FindStrInDict('/Pages') then exit;
    //get Pages object number then goto pagesNum ...
    if not GetUInt(pagesNum) or not GotoObject(pagesNum) then exit;
    if not FindStrInDict('/Count') then exit;
    if not GetUInt(cnt) then exit;
    //occasionally the 'count' value is an indirect object
    if GetUInt(k) and IsString(' R') then
    begin
      if not GotoObject(cnt) or
        not GetUInt(k) or //skip the generation num
        not IsString(' obj') or
        not GetUInt(cnt) then exit;
    end;
    result := cnt; //FOUND!!!!!!

  finally
    Clear;
  end;
end;

//------------------------------------------------------------------------------
//------------------------------------------------------------------------------

function GetPageCount(const filename: string): integer;
begin
  with TPdfPageCounter.Create do
  try
    Result := GetPdfPageCount(fileName);
    if ErrorFlag <> PDF_NO_ERROR then
      Result := ErrorFlag;
  finally
    free;
  end;
end;
//------------------------------------------------------------------------------

end.
