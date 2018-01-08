{
  Copyright (C) 2017-2018 Alexander Kernozhitsky <sh200105@mail.ru>

  This source is free software; you can redistribute it and/or modify it under
  the terms of the GNU General Public License as published by the Free
  Software Foundation; either version 2 of the License, or (at your option)
  any later version.

  This code is distributed in the hope that it will be useful, but WITHOUT ANY
  WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
  FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more
  details.

  A copy of the GNU General Public License is available on the World Wide Web
  at <http://www.gnu.org/copyleft/gpl.html>. You can also obtain it by writing
  to the Free Software Foundation, Inc., 51 Franklin Street - Fifth Floor,
  Boston, MA 02110-1335, USA.
}
unit mainunit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, ExtCtrls, StdCtrls, EditBtn, Dialogs, jsonparser,
  fpjson, jsonscanner, LazFileUtils, fgl, fpspreadsheet, gmap, gutil, gvector,
  fpsTypes, fpspreadsheetgrid, fpsRPN, Math, fpsutils;

type

  { TSolution }

  TSolution = class
  private
    FOwner: string;
    FScore: double;
    FTaskLetter: char;
    procedure SetOwner(AValue: string);
    procedure SetScore(AValue: double);
    procedure SetTaskLetter(AValue: char);
  public
    property Owner: string read FOwner write SetOwner;
    property TaskLetter: char read FTaskLetter write SetTaskLetter;
    property Score: double read FScore write SetScore;
    constructor Create;
    constructor Create(AOwner: string; ATaskLetter: char; AScore: double);
  end;

  TSolutionList = specialize TFPGObjectList<TSolution>;

  { TMainForm }

  TMainForm = class(TForm)
    FileNameEdit: TFileNameEdit;
    FilePanel: TPanel;
    FileLabel: TLabel;
    Panel: TPanel;
    SaveBtn: TButton;
    AboutBtn: TButton;
    SaveDialog: TSaveDialog;
    sWorksheetGrid: TsWorksheetGrid;
    procedure AboutBtnClick(Sender: TObject);
    procedure FileNameEditChange(Sender: TObject);
    procedure SaveBtnClick(Sender: TObject);
    procedure SaveDialogTypeChange(Sender: TObject);
  private
    FList: TSolutionList;
    procedure DoCollect;
    procedure DoSave;
  public
    constructor Create(TheOwner: TComponent); override;
    destructor Destroy; override;
  end;

var
  MainForm: TMainForm;

implementation

{$R *.lfm}

{ TSolution }

procedure TSolution.SetOwner(AValue: string);
begin
  if FOwner = AValue then
    Exit;
  FOwner := AValue;
end;

procedure TSolution.SetScore(AValue: double);
begin
  if FScore = AValue then
    Exit;
  FScore := AValue;
end;

procedure TSolution.SetTaskLetter(AValue: char);
begin
  if FTaskLetter = AValue then
    Exit;
  FTaskLetter := AValue;
end;

constructor TSolution.Create;
begin
  FOwner := '';
  FScore := 0.0;
  FTaskLetter := #0;
end;

constructor TSolution.Create(AOwner: string; ATaskLetter: char; AScore: double);
begin
  FOwner := AOwner;
  FTaskLetter := ATaskLetter;
  FScore := AScore;
end;

{ TMainForm }

procedure TMainForm.DoCollect;
var
  JSONFile: string;
  Parser: TJSONParser;
  Stream: TFileStream;
  Contents: TJSONData;
  ContentsItems: TJSONData;
  Tests: TJSONData;
  I, J: integer;
  Solution: TJSONObject;
  SolutionOwner: string;
  TaskLetter: char;
  SolutionFileName: string;
  Score: double;
  AddScore: double;
begin
  // pre-requisites
  FList.Clear;
  if FileNameEdit.DialogFiles.Count = 0 then
    exit;

  // iterate over the selected JSON files
  for JSONFile in FileNameEdit.DialogFiles do
  begin
    Stream := TFileStream.Create(JSONFile, fmOpenRead);
    try
      Parser := TJSONParser.Create(Stream, DefaultOptions);
      try
        Contents := Parser.Parse;
        try
          ContentsItems := (Contents as TJSONObject).Elements['Items'];

          // extract all the solutions
          for I := 0 to ContentsItems.Count - 1 do
          begin
            Solution := ((ContentsItems.Items[I] as TJSONObject)
              .Elements['Tester'] as TJSONObject);

            // getting the source owner and task ID from filename
            SolutionFileName := Solution.Elements['SourceFile'].AsString;
            SolutionOwner := ExtractFileNameOnly(SolutionFileName);

            // Solutions named as <nickname><problem>.<ext>
            // e. g. GepardovA.pas, alex1.cpp
            // TaskLetter := SolutionOwner[Length(SolutionOwner)];
            // Delete(SolutionOwner, Length(SolutionOwner), 1);

            // Solutions named as in Mogilev city olympiad
            // T<participant code>Z<problem code><region>.<ext>
            // e. g. T111Z233.pas, T444Z566.cpp
            TaskLetter := SolutionOwner[6];
            SolutionOwner := Copy(SolutionOwner, 2, 3);

            // calculating the total score
            Score := 0;
            Tests := (((ContentsItems.Items[I] as TJSONObject)
              .Elements['Tester'] as TJSONObject).Elements['Results'] as
              TJSONObject).Elements['Items'];
            for J := 0 to Tests.Count - 1 do
            begin
              AddScore := (Tests.Items[J] as TJSONObject).Elements['Score'].AsFloat;
              Score := Score + AddScore;
            end;

            // adding results to the list
            FList.Add(TSolution.Create(SolutionOwner, TaskLetter, Score));
          end;
        finally
          FreeAndNil(Contents);
        end;
      finally
        FreeAndNil(Parser);
      end;
    finally
      FreeAndNil(Stream);
    end;
  end;
end;

procedure TMainForm.DoSave;
type
  // generics specializations
  TStringLess = specialize TLess<string>;
  TCharLess = specialize TLess<char>;
  TStringMap = specialize TMap<string, integer, TStringLess>;
  TCharMap = specialize TMap<char, integer, TCharLess>;
  TStringVector = specialize TVector<string>;
  TCharVector = specialize TVector<char>;
var
  Owners: TStringMap;
  OwnerNames: TStringVector;
  OwnerCount, ProblemCount: integer;
  Problems: TCharMap;
  ProblemNames: TCharVector;
  OwnerIter: TStringMap.TIterator;
  ProblemIter: TCharMap.TIterator;
  Solution: TSolution;

  procedure CompressData;
  begin
    // add all the owners and problems to the map
    for Solution in FList do
    begin
      Owners[Solution.Owner] := -1;
      Problems[Solution.TaskLetter] := -1;
    end;

    // assign an index for each owner in the map
    OwnerCount := 0;
    OwnerIter := Owners.Min;
    try
      if OwnerIter <> nil then
      begin
        repeat
          OwnerIter.Value := OwnerCount;
          OwnerNames.PushBack(OwnerIter.Key);
          Inc(OwnerCount);
        until not OwnerIter.Next;
      end;
    finally
      FreeAndNil(OwnerIter);
    end;

    // assign an index for each problem in the map
    ProblemCount := 0;
    ProblemIter := Problems.Min;
    try
      if ProblemIter <> nil then
      begin
        repeat
          ProblemIter.Value := ProblemCount;
          ProblemNames.PushBack(ProblemIter.Key);
          Inc(ProblemCount);
        until not ProblemIter.Next;
      end;
    finally
      FreeAndNil(ProblemIter);
    end;
  end;

  procedure MakeSpreadSheet;
  const
    // color theme for the table
    EvenTopBk: TsColor = $00F5E7CF;
    OddTopBk: TsColor = $00FAC081;
    EvenCellBk: TsColor = $00FAF3E7;
    OddCellBk: TsColor = $00FCDFC0;
    HeaderBk: TsColor = $00FAC081;
    DoubleBk: TsColor = $00FF9933;
  var
    Workbook: TsWorkbook;
    Worksheet: TsWorksheet;
    I, J: integer;
    Solution: TSolution;
    TotalCol: integer;
    OwnerId, ProblemId: integer;
    MaxH, MaxW: integer;
    MaxTextLength: integer;
    SortParams: TsSortParams;

    procedure PutHeaderStyle(CurCell: PCell; const S: string;
      AColor: TsColor; ABkColor: TsColor);
    begin
      Worksheet.WriteText(CurCell, S);
      Worksheet.WriteFont(CurCell, 'sans-serif', 10.0, [fssBold], AColor);
      Worksheet.WriteHorAlignment(CurCell, haRight);
      Worksheet.WriteBackgroundColor(CurCell, ABkColor);
    end;

    procedure PutCellValue(CurCell: PCell; Val: double);
    begin
      Worksheet.WriteNumber(CurCell, Val);
      Worksheet.WriteFont(CurCell, 'sans-serif', 10.0, [], scBlack);
      Worksheet.WriteHorAlignment(CurCell, haRight);
    end;

    function GetTopBk(Col: integer): TsColor;
    begin
      if Odd(Col) then
        Result := OddTopBk
      else
        Result := EvenTopBk;
    end;

    function GetCellBk(Col: integer): TsColor;
    begin
      if Odd(Col) then
        Result := OddCellBk
      else
        Result := EvenCellBk;
    end;

  begin
    // ask user for the file name
    if not SaveDialog.Execute then
      exit;

    // create the workbook
    Workbook := TsWorkbook.Create;
    Workbook.SetDefaultFont('sans-serif', 10);
    Workbook.Options := Workbook.Options + [boAutoCalc, boCalcBeforeSaving];
    try
      Worksheet := Workbook.AddWorksheet('Results');
      MaxH := OwnerNames.Size;
      MaxW := ProblemNames.Size + 1;
      TotalCol := ProblemNames.Size + 1;

      // add the headers
      for I := 0 to integer(ProblemNames.Size) - 1 do
      begin
        PutHeaderStyle(Worksheet.GetCell(0, I + 1), ProblemNames[I],
          scBlack, GetTopBk(I + 1));
      end;

      PutHeaderStyle(Worksheet.GetCell(0, 0), 'Name', scBlack, DoubleBk);
      Worksheet.WriteFontStyle(0, 0, [fssItalic, fssBold]);

      PutHeaderStyle(Worksheet.GetCell(0, TotalCol), 'Total', scBlack, DoubleBk);
      Worksheet.WriteFontStyle(0, TotalCol, [fssItalic, fssBold]);

      for I := 0 to integer(OwnerNames.Size) - 1 do
      begin
        PutHeaderStyle(Worksheet.GetCell(I + 1, TotalCol), '', scBlack, HeaderBk);
        PutHeaderStyle(Worksheet.GetCell(I + 1, 0), OwnerNames[I], scBlack, HeaderBk);
      end;

      // calculate first column width
      MaxTextLength := 0;
      for I := 0 to integer(OwnerNames.Size) - 1 do
      begin
        MaxTextLength := Max(MaxTextLength, Length(OwnerNames[I]));
      end;
      Worksheet.WriteColWidth(0, MaxTextLength * 1.2, suChars);

      // add the scores
      for Solution in FList do
      begin
        OwnerId := Owners[Solution.Owner];
        ProblemId := Problems[Solution.TaskLetter];
        PutCellValue(Worksheet.GetCell(OwnerId + 1, ProblemId + 1), Solution.Score);
      end;

      for I := 0 to integer(OwnerNames.Size) - 1 do
      begin
        for J := 0 to integer(ProblemNames.Size) - 1 do
        begin
          Worksheet.WriteBackgroundColor(I + 1, J + 1, GetCellBk(J + 1));
        end;
      end;

      // calculate total values
      for I := 0 to integer(OwnerNames.Size) - 1 do
      begin
        Worksheet.WriteRPNFormula(I + 1, TotalCol, CreateRPNFormula(
          RPNCellRange(I + 1, 1, I + 1, TotalCol - 1, [rfRelRow, rfRelRow2],
          RPNFunc('SUM', nil))));
      end;

      // add borders
      for I := 0 to MaxH do
      begin
        for J := 0 to MaxW do
        begin
          Worksheet.WriteBorders(I, J, [cbNorth, cbSouth, cbEast, cbWest]);
        end;
      end;

      // sort the table
      SortParams := InitSortParams(True, 2);
      with SortParams.Keys[0] do
      begin
        ColRowIndex := TotalCol;
        Options := [ssoDescending];
      end;
      with SortParams.Keys[1] do
      begin
        ColRowIndex := 0;
        Options := [];
      end;
      Worksheet.Sort(SortParams, 1, 0, MaxH, TotalCol);

      // recalculate total values - it's needed because TsWorksheet.Sort() works
      // wrong with columns that contain formulas
      // see this bug report: https://bugs.freepascal.org/view.php?id=31887
      for I := 0 to integer(OwnerNames.Size) - 1 do
      begin
        Worksheet.WriteRPNFormula(I + 1, TotalCol, CreateRPNFormula(
          RPNCellRange(I + 1, 1, I + 1, TotalCol - 1, [rfRelRow, rfRelRow2],
          RPNFunc('SUM', nil))));
      end;

      // finally, save it
      Workbook.WriteToFile(SaveDialog.FileName, True);
      sWorksheetGrid.LoadFromSpreadsheetFile(SaveDialog.FileName);
    finally
      FreeAndNil(Workbook);
    end;
  end;

begin
  Owners := nil;
  OwnerNames := nil;
  Problems := nil;
  ProblemNames := nil;

  Owners := TStringMap.Create;
  OwnerNames := TStringVector.Create;
  Problems := TCharMap.Create;
  ProblemNames := TCharVector.Create;
  try
    CompressData;
    MakeSpreadSheet;
  finally
    FreeAndNil(Owners);
    FreeAndNil(OwnerNames);
    FreeAndNil(Problems);
    FreeAndNil(ProblemNames);
  end;
end;

procedure TMainForm.SaveDialogTypeChange(Sender: TObject);
const
  DefaultExts: array [1 .. 3] of string = ('.xls', '.xlsx', '.ods');
begin
  with SaveDialog do
  begin
    DefaultExt := DefaultExts[FilterIndex];
  end;
end;

procedure TMainForm.SaveBtnClick(Sender: TObject);
begin
  DoCollect;
  DoSave;
end;

procedure TMainForm.FileNameEditChange(Sender: TObject);
begin
  SaveBtn.Enabled := FileNameEdit.DialogFiles.Count > 0;
end;

procedure TMainForm.AboutBtnClick(Sender: TObject);
begin
  MessageDlg('About XLS Collector',
    'XLS Collector' + LineEnding +
    '' + LineEnding +
    'Copyright (C) 2017-2018 Alexander Kernozhitsky' + LineEnding +
    '' + LineEnding +
    'Collects Tester testing results into a spreadsheet.' + LineEnding +
    '' + LineEnding +
    'This program is licensed under GNU GPL v2 (or later)',
  mtInformation, [mbOK], 0);
end;

constructor TMainForm.Create(TheOwner: TComponent);
begin
  inherited Create(TheOwner);
  FList := TSolutionList.Create;
end;

destructor TMainForm.Destroy;
begin
  FreeAndNil(FList);
  inherited Destroy;
end;

end.
