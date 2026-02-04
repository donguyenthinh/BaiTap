import re
from pathlib import Path

import win32com.client as win32


WB_PATH = r"e:\Python\HocTap2\TaiHoaDonDienTu_v6.1_luuUserPass_GPE,.xlsm"


def get_module_text(vbcomponent) -> str:
    cm = vbcomponent.CodeModule
    return cm.Lines(1, cm.CountOfLines)


def set_module_text(vbcomponent, new_text: str) -> None:
    cm = vbcomponent.CodeModule
    # Delete all lines then add back
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)
    cm.AddFromString(new_text)


def replace_vba_block(text: str, pattern: str, replacement: str, what: str) -> str:
    new_text, n = re.subn(pattern, replacement, text, flags=re.IGNORECASE | re.DOTALL)
    if n != 1:
        raise RuntimeError(f"Expected to patch exactly 1 {what} block, patched {n}.")
    return new_text


HTTPGET_REPLACEMENT = r'''Function httpGet(ByVal url As String, bearer As String, res As String) As Boolean
    'GET with retry/backoff when server returns 429 (Too Many Requests)
    On Error GoTo EH

    Dim attempt As Long
    Dim backoffMs As Long
    Dim retryAfter As Long
    Dim st As String

    backoffMs = 1500 'start backoff (ms)

    For attempt = 1 To 6
        With xmlHttp 'CreateObject("MSXML2.serverXMLHTTP.6.0")
            .Open "GET", url, False
            .setRequestHeader "User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
            .setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
            .setRequestHeader "Accept-Encoding", "gzip;q=1.0"
            .setRequestHeader "Content-Encoding", "gzip"
            .setRequestHeader "Content-type", "application/gzip;application/json; application/x-www-form-urlencoded; charset=UTF-8"
            If Len(bearer) > 0 Then .setRequestHeader "Authorization", "Bearer " & bearer & ""
            .send

            st = CStr(.Status)
            getStatus = st
            res = .responseText

            If st = "200" Or st = "300" Then
                httpGet = True
                Exit Function
            End If

            If st = "429" Then
                'Rate-limited by server: wait and retry with exponential backoff.
                retryAfter = 0
                On Error Resume Next
                retryAfter = CLng(Val(.getResponseHeader("Retry-After"))) * 1000
                On Error GoTo 0

                If retryAfter > 0 Then
                    backoffMs = retryAfter
                End If

                'Cap backoff to 30s
                If backoffMs > 30000 Then backoffMs = 30000
                Sleep backoffMs
                backoffMs = backoffMs * 2
                If backoffMs < 1500 Then backoffMs = 1500
                GoTo next_attempt
            ElseIf st = "500" Then
                'Temporary server error - small wait and retry
                Sleep backoffMs
                backoffMs = backoffMs * 2
                If backoffMs > 30000 Then backoffMs = 30000
                GoTo next_attempt
            Else
                httpGet = False
                Exit Function
            End If
        End With

next_attempt:
        DoEvents
    Next attempt

    httpGet = False
    Exit Function

EH_Exit:
    Exit Function

EH:
    Select Case err.Number
        Case -2147012889
            MsgBoxUni "Loi ket noi Server hoac sai dia chi Web", vbExclamation, "Thông báo - Http get error"
        Case -2147012894
            MsgBoxUni "Qu" & ChrW(225) & " th" & ChrW(7901) & "i gian k" & ChrW(7871) & "t n" & ChrW(7889) & "i v" & ChrW(7899) & "i m" & ChrW(225) & "y ch" & ChrW(7911) & " Web." & vbCrLf _
            & "Ki" & ChrW(7875) & "m tra " & ChrW(273) & ChrW(432) & ChrW(7901) & "ng truy" & ChrW(7873) & "n Internet v" & ChrW(224) & " th" & ChrW(7917) & " k" & ChrW(7871) & "t n" & ChrW(7889) & "i l" & ChrW(7841) & "i.", vbExclamation, "Thông báo - Http get error"
        Case Else
            MsgBox "ErrNum: " & err.Number & vbCrLf & "Err description: " & err.Description, vbExclamation, "Thông báo - Http Get error"
    End Select
    httpGet = False
    Resume EH_Exit
End Function'''


TAIXMLZIP_REPLACEMENT = r'''Sub taiXML_zip(soHD As Long)
    On Error GoTo EH
    Dim nbmst As String, khhdon As String, shdon As String, khmshdon As String
    Dim urlXml As String, payload As String, filePath As Variant, m As Long

    '/ Tao folder luu file giai nen
    Dim strDate As String
    strDate = Format(Now, "_yyyymmdd_hhmmss")
    oSaveUnzipFolder = Me.txtXMLFolderPath & "FileGiaiNen" & strDate
    MkDir oSaveUnzipFolder
    MsgBox oSaveUnzipFolder
    '-----------------------------------/

    Application.Cursor = xlWait

    Dim attempt As Long
    Dim backoffMs As Long
    Dim retryAfter As Long

    For m = 0 To soHD - 1
        If arrHDChiTiet(m, 4) = 1 Then  'query
            urlXml = "https://hoadondientu.gdt.gov.vn:30000/query/invoices/export-xml?"
        Else    'sco-query
            urlXml = "https://hoadondientu.gdt.gov.vn:30000/sco-query/invoices/export-xml?"
        End If
        nbmst = arrHDChiTiet(m, 0)
        khhdon = arrHDChiTiet(m, 1)
        shdon = arrHDChiTiet(m, 2)
        khmshdon = arrHDChiTiet(m, 3)
        payload = "nbmst=" & nbmst & "&khhdon=" & khhdon & "&shdon=" & shdon & "&khmshdon=" & khmshdon
        urlXml = urlXml & payload

        backoffMs = 1500

        For attempt = 1 To 6
            Dim http As Object, stream As Object
            Set stream = CreateObject("ADODB.Stream")
            Set http = CreateObject("MSXML2.XMLHTTP")

            With http
                .Open "GET", urlXml, False
                .setRequestHeader "User-Agent", "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"
                .setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9"
                .setRequestHeader "Accept-Encoding", "gzip"
                .setRequestHeader "Content-type", "application/zip"
                .setRequestHeader "Authorization", "Bearer " & bearer & ""
                .send
            End With

            Select Case http.Status
                Case 200
                    stream.Type = 1 ' Binary
                    stream.Open
                    stream.Write http.responseBody
                    filePath = Me.txtXMLFolderPath & khhdon & "_" & shdon & ".zip"
                    stream.SaveToFile filePath, 2
                    stream.Close
                    DoEvents

                    '//Unzip file vua tai, chi lay file xml/html
                    Call Unzip(filePath, oSaveUnzipFolder)
                    GoTo nextInvoice_ok

                Case 500
                    Debug.Print "Loi download file hoa don (zip): " & http.Status & "__" & khhdon & "_" & shdon & " res: " & http.responseText
                    errMsg = errMsg & "- Khong ton tai ho so goc cua hoa don: " & khhdon & "_" & shdon & vbCrLf
                    GoTo nextInvoice_ok

                Case 429
                    'Rate-limited: wait then retry with backoff
                    retryAfter = 0
                    On Error Resume Next
                    retryAfter = CLng(Val(http.getResponseHeader("Retry-After"))) * 1000
                    On Error GoTo 0

                    If retryAfter > 0 Then backoffMs = retryAfter
                    If backoffMs > 30000 Then backoffMs = 30000
                    Sleep backoffMs
                    backoffMs = backoffMs * 2
                    If backoffMs < 1500 Then backoffMs = 1500
                    DoEvents
                    'Try again
                Case Else
                    errMsg = errMsg & "- Loi ket noi: " & khhdon & "_" & shdon & " (Status " & http.Status & ")" & vbCrLf
                    GoTo nextInvoice_ok
            End Select

            Set stream = Nothing
            Set http = Nothing
        Next attempt

        'After retries still failing with 429 -> log
        errMsg = errMsg & "- Loi 429 (Too Many Requests): " & khhdon & "_" & shdon & vbCrLf

nextInvoice_ok:
        Set stream = Nothing
        Set http = Nothing
        DoEvents
    Next m

EH_Exit:
    Application.Cursor = xlDefault
    If errMsg <> "" Then
        ghiLog errMsg
    End If
    Exit Sub
EH:
    Application.Cursor = xlDefault
    MsgBoxUni "C" & ChrW(243) & " l" & ChrW(7895) & "i ph" & ChrW(225) & "t sinh trong qu" & ChrW(225) & " tr" & ChrW(236) & "nh t" & ChrW(7843) & "i h" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n.", vbCritical, "L" & ChrW(7895) & "i t" & ChrW(7843) & "i h" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n"
    MsgBox "Ma loi: " & err.Number & vbCrLf & "Noi dung: " & err.Description, vbCritical, "L" & ChrW(7895) & "i t" & ChrW(7843) & "i h" & ChrW(243) & "a " & ChrW(273) & ChrW(417) & "n"
    Resume EH_Exit
End Sub'''


def patch_mod_http_request(text: str) -> str:
    # Replace entire Function httpGet(...) ... End Function
    pat = r"Function\s+httpGet\s*\([\s\S]*?\nEnd Function"
    return replace_vba_block(text, pat, HTTPGET_REPLACEMENT, "httpGet")


def patch_frm_tai_hoa_don(text: str) -> str:
    # Replace entire Sub taiXML_zip(...) ... End Sub
    pat = r"Sub\s+taiXML_zip\s*\([\s\S]*?\nEnd Sub"
    new_text = replace_vba_block(text, pat, TAIXMLZIP_REPLACEMENT, "taiXML_zip")

    # Set default txtSleep to 1500 in UserForm_Initialize
    # If it already assigns txtSleep, replace value; otherwise insert assignment near end.
    init_pat = r"(Private\s+Sub\s+UserForm_Initialize\s*\(\)\s*[\s\S]*?\nEnd Sub)"
    m = re.search(init_pat, new_text, flags=re.IGNORECASE)
    if not m:
        raise RuntimeError("Could not find UserForm_Initialize in frmTaiHoaDon.")
    init_block = m.group(1)

    if re.search(r"Me\.txtSleep\s*=", init_block, flags=re.IGNORECASE):
        init_block2 = re.sub(
            r"(Me\.txtSleep\s*=\s*)(\d+|\"[^\"]*\"|[A-Za-z_]\w*)",
            r"\g<1>1500",
            init_block,
            count=1,
            flags=re.IGNORECASE,
        )
    else:
        init_block2 = init_block.replace(
            "End Sub",
            '    Me.txtSleep = 1500\nEnd Sub',
        )

    new_text = new_text[: m.start(1)] + init_block2 + new_text[m.end(1) :]
    return new_text


def main() -> None:
    path = Path(WB_PATH)
    if not path.exists():
        raise FileNotFoundError(WB_PATH)

    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False

    wb = None
    try:
        wb = xl.Workbooks.Open(str(path), ReadOnly=False)
        vp = wb.VBProject

        mod_http = vp.VBComponents("modHTTPRequest")
        frm_main = vp.VBComponents("frmTaiHoaDon")

        http_text = get_module_text(mod_http)
        frm_text = get_module_text(frm_main)

        http_text_new = patch_mod_http_request(http_text)
        frm_text_new = patch_frm_tai_hoa_don(frm_text)

        set_module_text(mod_http, http_text_new)
        set_module_text(frm_main, frm_text_new)

        wb.Save()
        print("Patched VBA successfully.")
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        xl.Quit()


if __name__ == "__main__":
    main()

