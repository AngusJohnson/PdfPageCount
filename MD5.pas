unit MD5;

(*******************************************************************************
*                                                                              *
* Author    :  Angus Johnson                                                   *
* Version   :  1.2                                                             *
* Date      :  12 February 2024                                                *
* Website   :  http://www.angusj.com                                           *
* Copyright :  Angus Johnson 2010-2024                                         *
*                                                                              *
* License:                                                                     *
* Use, modification & distribution is subject to Boost Software License Ver 1. *
* http://www.boost.org/LICENSE_1_0.txt                                         *
*                                                                              *
*******************************************************************************)

interface

{$Q-}

uses
  Windows, SysUtils, Classes;

const
  BuffLen64 = 64;

type
{$IFNDEF UNICODE}
  PByte = PChar;
{$ENDIF}
  PByteArray    = ^TByteArray;
  TByteArray    = array[0..MaxInt-1] of Byte;
  TByteArray16  = array [0..15] of Byte;

  Md5Record = record
  private
    state     : array [0..3] of cardinal;
    count     : array [0..1] of cardinal;
    digest    : TByteArray16;
    buffer64  : array [0..BuffLen64 -1] of Byte;
    procedure Transform;

    procedure F1(var a: cardinal; b, c, d, x, s: cardinal);
    procedure F2(var a: cardinal; b, c, d, x, s: cardinal);
    procedure F3(var a: cardinal; b, c, d, x, s: cardinal);
    procedure F4(var a: cardinal; b, c, d, x, s: cardinal);
  public
    procedure Init;
    procedure Update(start: PByteArray; length: cardinal);
    procedure Finalize;
    property  Hash: TByteArray16 read digest;
  end;

implementation

//------------------------------------------------------------------------------
//------------------------------------------------------------------------------

procedure Md5Record.Init;
begin
  count[0] := 0;
  count[1] := 0;

  //Load magic constants ...
  state[0] := $67452301;
  state[1] := $efcdab89;
  state[2] := $98badcfe;
  state[3] := $10325476;
end;
//------------------------------------------------------------------------------

procedure Md5Record.Finalize;
var
  cnt: cardinal;
  pb  : system.PByte;
begin
  cnt := (count[0] shr 3) and $3f;
  pb := @buffer64[cnt];
  pb^ := $80;
  inc(pb);
  cnt := 63 - cnt;
  if cnt < 8 then
  begin
    // Two lots of padding:  Pad the first block to 64 bytes */
    FillChar(pb^, cnt, 0);
    Transform;
    // Now fill the next block with 56 bytes */
    FillChar(buffer64[0], 56, 0);
  end
  else if cnt > 8 then
    FillChar(pb^, cnt -8, 0);

  // Append length in bits and transform ...
  move(count[0], buffer64[56], 4);
  move(count[1], buffer64[60], 4);
  Transform;
  move(state[0], digest[0], 16);
end;
//------------------------------------------------------------------------------

procedure Md5Record.Transform;
var
  a,b,c,d: cardinal;
  x: array[0..15] of cardinal;
begin
  Move(buffer64[0], x[0], BuffLen64);
  a := state[0]; b := state[1]; c := state[2]; d := state[3];

  // Round 1 ...
  F1 (a, b, c, d, x[ 0] + $d76aa478, 7);
  F1 (d, a, b, c, x[ 1] + $e8c7b756, 12);
  F1 (c, d, a, b, x[ 2] + $242070db, 17);
  F1 (b, c, d, a, x[ 3] + $c1bdceee, 22);
  F1 (a, b, c, d, x[ 4] + $f57c0faf, 7);
  F1 (d, a, b, c, x[ 5] + $4787c62a, 12);
  F1 (c, d, a, b, x[ 6] + $a8304613, 17);
  F1 (b, c, d, a, x[ 7] + $fd469501, 22);
  F1 (a, b, c, d, x[ 8] + $698098d8, 7);
  F1 (d, a, b, c, x[ 9] + $8b44f7af, 12);
  F1 (c, d, a, b, x[10] + $ffff5bb1, 17);
  F1 (b, c, d, a, x[11] + $895cd7be, 22);
  F1 (a, b, c, d, x[12] + $6b901122, 7);
  F1 (d, a, b, c, x[13] + $fd987193, 12);
  F1 (c, d, a, b, x[14] + $a679438e, 17);
  F1 (b, c, d, a, x[15] + $49b40821, 22);

  // Round 2 ...
  F2 (a, b, c, d, x[ 1] + $f61e2562, 5);
  F2 (d, a, b, c, x[ 6] + $c040b340, 9);
  F2 (c, d, a, b, x[11] + $265e5a51, 14);
  F2 (b, c, d, a, x[ 0] + $e9b6c7aa, 20);
  F2 (a, b, c, d, x[ 5] + $d62f105d, 5);
  F2 (d, a, b, c, x[10] + $02441453, 9);
  F2 (c, d, a, b, x[15] + $d8a1e681, 14);
  F2 (b, c, d, a, x[ 4] + $e7d3fbc8, 20);
  F2 (a, b, c, d, x[ 9] + $21e1cde6, 5);
  F2 (d, a, b, c, x[14] + $c33707d6, 9);
  F2 (c, d, a, b, x[ 3] + $f4d50d87, 14);
  F2 (b, c, d, a, x[ 8] + $455a14ed, 20);
  F2 (a, b, c, d, x[13] + $a9e3e905, 5);
  F2 (d, a, b, c, x[ 2] + $fcefa3f8, 9);
  F2 (c, d, a, b, x[ 7] + $676f02d9, 14);
  F2 (b, c, d, a, x[12] + $8d2a4c8a, 20);

  // Round 3 ..
  F3 (a, b, c, d, x[ 5] + $fffa3942, 4);
  F3 (d, a, b, c, x[ 8] + $8771f681, 11);
  F3 (c, d, a, b, x[11] + $6d9d6122, 16);
  F3 (b, c, d, a, x[14] + $fde5380c, 23);
  F3 (a, b, c, d, x[ 1] + $a4beea44, 4);
  F3 (d, a, b, c, x[ 4] + $4bdecfa9, 11);
  F3 (c, d, a, b, x[ 7] + $f6bb4b60, 16);
  F3 (b, c, d, a, x[10] + $bebfbc70, 23);
  F3 (a, b, c, d, x[13] + $289b7ec6, 4);
  F3 (d, a, b, c, x[ 0] + $eaa127fa, 11);
  F3 (c, d, a, b, x[ 3] + $d4ef3085, 16);
  F3 (b, c, d, a, x[ 6] + $4881d05, 23);
  F3 (a, b, c, d, x[ 9] + $d9d4d039, 4);
  F3 (d, a, b, c, x[12] + $e6db99e5, 11);
  F3 (c, d, a, b, x[15] + $1fa27cf8, 16);
  F3 (b, c, d, a, x[ 2] + $c4ac5665, 23);

  // Round 4 ..
  F4 (a, b, c, d, x[ 0] + $f4292244, 6);
  F4 (d, a, b, c, x[ 7] + $432aff97, 10);
  F4 (c, d, a, b, x[14] + $ab9423a7, 15);
  F4 (b, c, d, a, x[ 5] + $fc93a039, 21);
  F4 (a, b, c, d, x[12] + $655b59c3, 6);
  F4 (d, a, b, c, x[ 3] + $8f0ccc92, 10);
  F4 (c, d, a, b, x[10] + $ffeff47d, 15);
  F4 (b, c, d, a, x[ 1] + $85845dd1, 21);
  F4 (a, b, c, d, x[ 8] + $6fa87e4f, 6);
  F4 (d, a, b, c, x[15] + $fe2ce6e0, 10);
  F4 (c, d, a, b, x[ 6] + $a3014314, 15);
  F4 (b, c, d, a, x[13] + $4e0811a1, 21);
  F4 (a, b, c, d, x[ 4] + $f7537e82, 6);
  F4 (d, a, b, c, x[11] + $bd3af235, 10);
  F4 (c, d, a, b, x[ 2] + $2ad7d2bb, 15);
  F4 (b, c, d, a, x[ 9] + $eb86d391, 21);

  inc(state[0], a);
  inc(state[1], b);
  inc(state[2], c);
  inc(state[3], d);
end;
//------------------------------------------------------------------------------

procedure Md5Record.Update(start: PByteArray; length: cardinal);
var
  t, idx, spc: cardinal;
  //sl: TStringList;
begin

  // Update number of bits
  t := count[0];
  count[0] := t + (length shl 3);
  if count[0] < t then inc(count[1]);
  inc(count[1], length shr 29);

  // Compute number of bytes mod 64
  idx := (t shr 3) and $3f;
  spc := BuffLen64 - idx;

  if idx > 0 then
  begin
    if (length < spc) then
    begin
      move(start^, buffer64[idx], length);
      exit;
    end;
    move(start^, buffer64[idx], spc);
    Transform;
    dec(length, spc);
    inc(PByte(start), spc);
  end;

  //Process data in 64-byte chunks ...
  while (length >= BuffLen64) do
  begin
    move(start^, buffer64[0], BuffLen64);
    Transform;
    dec(length, BuffLen64);
    inc(PByte(start), BuffLen64);
  end;

  //Finally, handle any remaining bytes of data ...
    move(start^, buffer64[0], length);
end;
//------------------------------------------------------------------------------

procedure Md5Record.F1(var a: cardinal; b, c, d, x, s: cardinal);
begin
 inc(a, (d xor (b and (c xor d))) + x);
 a := ((a shl s) or (a shr (32-s))) + b;
end;
//------------------------------------------------------------------------------

procedure Md5Record.F2(var a: cardinal; b, c, d, x, s: cardinal);
begin
 inc(a, ((b and d) or (c and not d)) + x);
 a := ((a shl s) or (a shr (32-s))) + b;
end;
//------------------------------------------------------------------------------

procedure Md5Record.F3(var a: cardinal; b, c, d, x, s: cardinal);
begin
 inc(a, (b xor c xor d) + x);
 a := ((a shl s) or (a shr (32-s))) + b;
end;
//------------------------------------------------------------------------------

procedure Md5Record.F4(var a: cardinal; b, c, d, x, s: cardinal);
begin
 inc(a, (c xor (b or not d)) + x);
 a := ((a shl s) or (a shr (32-s))) + b;
end;
//------------------------------------------------------------------------------

end.
