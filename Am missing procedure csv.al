page 50313 AttendaceUploader
{
    PageType = List;
    ApplicationArea = All;
    UsageCategory = Administration;
    SourceTable = Attendance;
    Caption = 'Attendance Uploader';
    // Editable = false;

    layout
    {
        area(Content)
        {
            repeater(GroupName)
            {
                field("Employee Id"; Rec."Employee Id")
                {
                    ApplicationArea = All;
                    Editable = false;

                }
                field("Attendance Date"; Rec."Attendance Date")
                {
                    ApplicationArea = All;
                    Editable = false;

                }
                field("Check In"; Rec."Check In")
                {
                    ApplicationArea = All;
                    Editable = false;

                }
                field("Check Out"; Rec."Check Out")
                {
                    ApplicationArea = All;
                    Editable = false;

                }
            }
        }
    }

    actions
    {
        area(Processing)
        {
            action("Upload Attendance")
            {
                Caption = 'Upload Attendance';
                Image = ImportExcel;
                ApplicationArea = All;

                trigger OnAction()
                begin

                    if ReadExcelSheet() = True then begin
                        ImportExcelData();
                    end;
                end;
            }
        }
    }
    local procedure ReadExcelSheet(): Boolean
    var
        FileMgt: Codeunit "File Management";
        IStream: InStream;
        FromFile: Text[100];
        isValidate: Boolean;
        FileName: Text[100];
        SheetName: Text[100];
        IsColor: Boolean;
        IsEditable: Boolean;
        AllExceptions: Text[2000];
    begin

        RowNo := 0;
        ColNo := 0;
        LineNo := 0;
        MaxRowNo := 0;
        isValidate := true;

        UploadIntoStream(UploadExcelMsg, '', '', FromFile, IStream);
        if FromFile <> '' then begin
            FileName := FileMgt.GetFileName(FromFile);
            SheetName := TempExcelBuffer.SelectSheetsNameStream(IStream);


            TempExcelBuffer.DeleteAll();
            TempExcelBuffer.OpenBookStream(IStream, SheetName);
            TempExcelBuffer.ReadSheet();

            if TempExcelBuffer.FindLast() then begin
                MaxRowNo := TempExcelBuffer."Row No.";
            end;

            exit(true);

        end;
    end;

    local procedure ImportExcelData()
    var
        Tbl_Attendance: Record Attendance;
        Tbl_AttendanceForLineNo: Record Attendance;
        BillingAddress: Text[100];
        CheckInTimeText: Text[50];
        CheckInTime: Time;
        AMCheckInTimeText: Text[50];
    begin
        for RowNo := 2 to MaxRowNo do begin
            Tbl_AttendanceForLineNo.Reset();
            IF Tbl_AttendanceForLineNo.FindLast() then begin
                LineNo := Tbl_AttendanceForLineNo."Line No." + 1;
            end Else
                LineNo := 1;

            Tbl_Attendance."Line No." := LineNo;
            Evaluate(Tbl_Attendance."Employee Id", GetValueAtCell(RowNo, 1));
            Evaluate(Tbl_Attendance."Attendance Date", GetValueAtCell(RowNo, 2));
            CheckInTimeText := GetValueAtCell(RowNo, 3);
            if CheckInTimeText <> '' then begin
                if STRPOS(CheckInTimeText, 'PM') = 0 then begin
                    AMCheckInTimeText := CheckInTimeText + ' AM';
                end
                else begin
                    AMCheckInTimeText := CheckInTimeText;
                end;
            end;
            Evaluate(Tbl_Attendance."Check In", AMCheckInTimeText);

            Evaluate(Tbl_Attendance."Check Out", GetValueAtCell(RowNo, 4));
            Tbl_Attendance.Insert();
        end;
    end;



    local procedure GetValueAtCell(RowNo: Integer; ColNo: Integer): Text
    begin

        TempExcelBuffer.Reset();
        If TempExcelBuffer.Get(RowNo, ColNo) then
            exit(TempExcelBuffer."Cell Value as Text")
        else
            exit('');
    end;

    var


        myInt: Integer;
        RowNo: Integer;
        ColNo: Integer;
        LineNo: Integer;
        MaxRowNo: Integer;
        TempExcelBuffer: Record "Excel Buffer" temporary;
        UploadExcelMsg: Label 'Please Choose the Excel file.';
        ExcelImportSucess: Label 'Excel is successfully imported.';
}
