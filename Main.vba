Option Explicit

Public column_headers As Collection
Public valid_fa_copany_types As Collection
Public valid_country_inputs As Collection
Public valid_paye_ni_period_inputs As Collection
Public valid_sales_tax_registration_status_inputs As Collection
Public valid_initial_vat_basis_inputs As Collection
Public valid_short_date_format_inputs As Collection
Public valid_status_inputs As Collection
Public valid_bank_account_type_inputs As Collection
Public invalid_sales_tax_registration_status_inputs_for_universal As Collection
Public valid_user_role_inputs As Collection
Public valid_user_permission_level_inputs As Collection
Public mySheet As Worksheet

Public niChecker As New RegExp
Public company_number_checker_eng_and_wales As New RegExp
Public company_number_checker_scotland As New RegExp
Public email_address_checker As New RegExp
Public paye_reference_checker As New RegExp
Public vat_registration_number_checker As New RegExp
Public postcode_checker As New RegExp
Public sort_code_checker As New RegExp
Public account_number_checker As New RegExp
Sub perform_checks_button()

    ' Prepare the Constants
    ValidationConstants.prepare

    ' Get the Sheet
    Set mySheet = Sheet1

    ' Clear all current formatting
    Call clear_all_cells

    ' Get the last active row
    Dim lastRow As Long
    lastRow = mySheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row

    Dim companyList As New Collection

    ' Check Each Row until last active row and return the object
    Dim row As Integer

    Dim checkedCompanies As New Collection

    For row = 2 To lastRow
        checkedCompanies.Add CheckRow(row)
    Next row

    Call check_subdomain_uniqueness(checkedCompanies)

End Sub
Function CheckRow(row As Integer) As Company

    Dim sRow As String
    sRow = CSng(row)

    ' New up Company Object
    Dim companyToTest As Company
    Set companyToTest = New Company

    ' Set properties of the Company Object
    With companyToTest
        .row = row
        .subdomain = Range(column_headers("subdomain") + sRow).Value
        .name = Range(column_headers("name") + sRow).Value
        .country = Range(column_headers("country") + sRow).Value
        .faCompanyType = Range(column_headers("type") + sRow).Value
        .paye_ni_period = Range(column_headers("paye_ni_period") + sRow).Value
        .sales_tax_registration_status = Range(column_headers("sales_tax_registration_status") + sRow).Value
        .initial_vat_basis = Range(column_headers("initial_vat_basis") + sRow).Value
        .short_date_format = Range(column_headers("short_date_format") + sRow).Value
        .account_manager_email = Range(column_headers("account_manager_email") + sRow).Value
        .status = Range(column_headers("status") + sRow).Value
        .initial_vat_frs_type_index = Range(column_headers("initial_vat_frs_type_index") + sRow).Value

        .registration_number = Range(column_headers("registration_number") + sRow).Value
        .paye_reference = Range(column_headers("paye_reference") + sRow).Value
        .vat_registration_number = Range(column_headers("vat_registration_number") + sRow).Value
        .postcode = Range(column_headers("postcode") + sRow).Value


        .bank_account_1.Add Range(column_headers("bank_account_1_type") + sRow).Value, "type"
        .bank_account_2.Add Range(column_headers("bank_account_2_type") + sRow).Value, "type"
        .bank_account_3.Add Range(column_headers("bank_account_3_type") + sRow).Value, "type"

        .bank_account_1.Add Range(column_headers("bank_account_1_name") + sRow).Value, "name"
        .bank_account_2.Add Range(column_headers("bank_account_2_name") + sRow).Value, "name"
        .bank_account_3.Add Range(column_headers("bank_account_3_name") + sRow).Value, "name"

        .bank_account_1.Add Range(column_headers("bank_account_1_email") + sRow).Value, "email"
        .bank_account_2.Add Range(column_headers("bank_account_2_email") + sRow).Value, "email"
        .bank_account_3.Add Range(column_headers("bank_account_3_email") + sRow).Value, "email"

        .bank_account_1.Add Range(column_headers("bank_account_1_account_number") + sRow).Value, "account_number"
        .bank_account_2.Add Range(column_headers("bank_account_2_account_number") + sRow).Value, "account_number"
        .bank_account_3.Add Range(column_headers("bank_account_3_account_number") + sRow).Value, "account_number"

        .bank_account_1.Add Range(column_headers("bank_account_1_sort_code") + sRow).Value, "sort_code"
        .bank_account_2.Add Range(column_headers("bank_account_2_sort_code") + sRow).Value, "sort_code"
        .bank_account_3.Add Range(column_headers("bank_account_3_sort_code") + sRow).Value, "sort_code"

        .bank_account_1.Add Range(column_headers("bank_account_1_opening_balance") + sRow).Value, "opening_balance"
        .bank_account_2.Add Range(column_headers("bank_account_2_opening_balance") + sRow).Value, "opening_balance"
        .bank_account_3.Add Range(column_headers("bank_account_3_opening_balance") + sRow).Value, "opening_balance"

        .user_1.Add Range(column_headers("user_1_first_name") + sRow).Value, "first_name"
        .user_2.Add Range(column_headers("user_2_first_name") + sRow).Value, "first_name"

        .user_1.Add Range(column_headers("user_1_last_name") + sRow).Value, "last_name"
        .user_2.Add Range(column_headers("user_2_last_name") + sRow).Value, "last_name"

        .user_1.Add Range(column_headers("user_1_role") + sRow).Value, "role"
        .user_2.Add Range(column_headers("user_2_role") + sRow).Value, "role"

        .user_1.Add Range(column_headers("user_1_permission_level") + sRow).Value, "permission_level"
        .user_2.Add Range(column_headers("user_2_permission_level") + sRow).Value, "permission_level"

        .user_1.Add Range(column_headers("user_1_ni_number") + sRow).Value, "ni_number"
        .user_2.Add Range(column_headers("user_2_ni_number") + sRow).Value, "ni_number"

        .user_1.Add Range(column_headers("user_1_email") + sRow).Value, "email"
        .user_2.Add Range(column_headers("user_2_email") + sRow).Value, "email"

        .user_1.Add Range(column_headers("user_1_capital_opening_balance") + sRow).Value, "capital_opening_balance"
        .user_2.Add Range(column_headers("user_2_capital_opening_balance") + sRow).Value, "capital_opening_balance"

        .user_1.Add Range(column_headers("user_1_directors_loan_opening_balance") + sRow).Value, "directors_loan_opening_balance"
        .user_2.Add Range(column_headers("user_2_directors_loan_opening_balance") + sRow).Value, "directors_loan_opening_balance"

        .user_1.Add Range(column_headers("user_1_expense_opening_balance") + sRow).Value, "expense_opening_balance"
        .user_2.Add Range(column_headers("user_2_expense_opening_balance") + sRow).Value, "expense_opening_balance"

        .user_1.Add Range(column_headers("user_1_salary_opening_balance") + sRow).Value, "salary_opening_balance"
        .user_2.Add Range(column_headers("user_2_salary_opening_balance") + sRow).Value, "salary_opening_balance"
    End With

    'Run the validations
    companyToTest.runTests

    'Return the Company
    Set CheckRow = companyToTest
End Function
Sub clear_all_cells()
    Set mySheet = Sheet1
    mySheet.Cells.Interior.ColorIndex = 0
End Sub
Sub ColourCell(sColumn As String, sRow As String)
    ActiveSheet.Range(sColumn + sRow).Interior.ColorIndex = 3
End Sub
Function is_in(c As Collection, to_check As String) As Boolean
    Dim Item As Integer
    Dim check_passed As Boolean
    check_passed = False
    For Item = 1 To c.Count
        If c.Item(Item) = to_check Then
            check_passed = True
        End If
    Next Item
    is_in = check_passed
End Function
Sub check_subdomain_uniqueness(companyCollection As Collection)
    Dim c As Integer
    Dim currentSubdomains As New Collection
    For c = 1 To companyCollection.Count
        Dim currentSubdomain As String
        currentSubdomain = companyCollection.Item(c).subdomain
        If currentSubdomain = "" Then
        Else
            If is_in(currentSubdomains, currentSubdomain) = True Then
                Call ColourCell("A", CStr(companyCollection.Item(c).row))
                  Dim i As Integer
                   For i = 1 To companyCollection.Count
                     If companyCollection.Item(i).subdomain = currentSubdomain Then
                            Call ColourCell("A", CStr(i + 1))
                        End If
                    Next i
            End If
        End If
        currentSubdomains.Add companyCollection.Item(c).subdomain
    Next c
End Sub
