Option Explicit

Public row As Integer
Public subdomain As String
Public name As String
Public country As String
Public registration_number As String
Public errors As Collection
Public faCompanyType As String
Public sales_tax_registration_status As String
Public paye_ni_period As String
Public account_manager_email As String
Public initial_vat_basis As String
Public short_date_format As String
Public status As String
Public paye_reference As String
Public vat_registration_number As String
Public postcode As String
Public initial_vat_frs_type_index As String

Public bank_accounts As Collection
Public bank_account_1 As Collection
Public bank_account_2 As Collection
Public bank_account_3 As Collection

Public users As Collection
Public user_1 As Collection
Public user_2 As Collection
Public user_3 As Collection
Private Sub Class_Initialize()
    Set bank_accounts = New Collection
    Set bank_account_1 = New Collection
    Set bank_account_2 = New Collection
    Set bank_account_3 = New Collection
    bank_accounts.Add bank_account_1
    bank_accounts.Add bank_account_2
    bank_accounts.Add bank_account_3

    Set users = New Collection
    Set user_1 = New Collection
    Set user_2 = New Collection
    users.Add user_1
    users.Add user_2
End Sub
Sub runTests()
    If check_valid_fa_company_type = False Then
        ColourCell (column_headers("Type"))
    End If

    If check_valid_paye_ni_period = False Then
        ColourCell (column_headers("paye_ni_period"))
    End If

    If check_valid_country = False Then
        ColourCell (column_headers("country"))
    End If

    If check_valid_sales_tax_registration_status = False Then
        ColourCell (column_headers("sales_tax_registration_status"))
    End If

    If check_initial_vat_basis = False Then
        ColourCell (column_headers("initial_vat_basis"))
    End If

    If check_valid_short_date_format = False Then
        ColourCell (column_headers("short_date_format"))
    End If

    If check_status = False Then
        ColourCell (column_headers("status"))
    End If

    ' If check_registration_number = False Then
    '     ColourCell (column_headers("registration_number"))
    ' End If

    If check_initial_vat_frs_type_index = False Then
        ColourCell (column_headers("initial_vat_frs_type_index"))
    End If

    ' Check account manager email address
    ' If regex_checker(account_manager_email, email_address_checker) = False Then
    '     ColourCell (column_headers("account_manager_email"))
    ' End If

    ' ' Check postcode
    ' If regex_checker(postcode, postcode_checker) = False Then
    '     ColourCell (column_headers("postcode"))
    ' End If

    ' ' Check vat_registration_number
    ' If regex_checker(vat_registration_number, vat_registration_number_checker) = False Then
    '     ColourCell (column_headers("vat_registration_number"))
    ' End If

    ' ' Check paye_reference
    ' If regex_checker(paye_reference, paye_reference_checker) = False Then
    '     ColourCell (column_headers("paye_reference"))
    ' End If

    'check_bank_accounts
    ' check_users
    'check_ni_number
    check_unique_bank_account_names
    check_required_fields_bas_and_users

End Sub
Function check_valid_fa_company_type() As Boolean
    check_valid_fa_company_type = is_in(valid_fa_copany_types, faCompanyType)
End Function
Function check_valid_paye_ni_period()
    check_valid_paye_ni_period = is_in(valid_paye_ni_period_inputs, paye_ni_period)
End Function
Function check_valid_country()
    check_valid_country = is_in(valid_country_inputs, country)
End Function
Function check_valid_sales_tax_registration_status() As Boolean
    check_valid_sales_tax_registration_status = is_in(valid_sales_tax_registration_status_inputs, sales_tax_registration_status)

    If faCompanyType = "Universal" Or faCompanyType = "" Then
        If is_in(invalid_sales_tax_registration_status_inputs_for_universal, sales_tax_registration_status) Then
            check_valid_sales_tax_registration_status = False
        End If
    End If
End Function
Function check_valid_short_date_format()
    check_valid_short_date_format = is_in(valid_short_date_format_inputs, short_date_format)
End Function
Function check_initial_vat_basis()
    check_initial_vat_basis = is_in(valid_initial_vat_basis_inputs, initial_vat_basis)
End Function
Function check_status()
    check_status = is_in(valid_status_inputs, status)
End Function
Function check_user_role(u As Integer)
    check_user_role = is_in(valid_user_role_inputs, users(u).Item("role"))
End Function

Function check_user_permission_level(u As Integer)
    check_user_permission_level = is_in(valid_user_permission_level_inputs, users(u).Item("permission_level"))
End Function
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
Sub ColourCell(sColumn As String)
    Dim sRow As String
    sRow = CSng(row)
    ActiveSheet.Range(sColumn + sRow).Interior.ColorIndex = 3
End Sub
Function get_subdomain() As String
    get_subdomain = subdomain
End Function
Function check_bank_accounts()
    Dim ba As Integer
    For ba = 1 To bank_accounts.Count
        If is_in(valid_bank_account_type_inputs, bank_accounts(ba).Item("type")) Then
            If bank_accounts(ba).Item("type") = "Credit Card" Then
                If Len(bank_accounts(ba).Item("account_number")) <> 8 Then
                    ColourCell (column_headers("bank_account_" & ba & "_account_number"))
                End If
            End If
            If bank_accounts(ba).Item("type") = "PayPal" Then
                If Len(bank_accounts(ba).Item("email")) < 1 Then
                    ColourCell (column_headers("bank_account_" & ba & "_email"))
                End If
            End If
        Else
            ColourCell (column_headers("bank_account_" & ba & "_type"))
        End If

        If check_email_address(bank_accounts(ba).Item("email")) = False Then
            ColourCell (column_headers("bank_account_" & ba & "_email"))
        End If

        If check_email_address(bank_accounts(ba).Item("email")) = False Then
            ColourCell (column_headers("bank_account_" & ba & "_email"))
        End If

        ' Check sort_code
        If regex_checker(bank_accounts(ba).Item("sort_code"), sort_code_checker) = False Then
            ColourCell (column_headers("bank_account_" & ba & "_sort_code"))
        End If

        ' Check account_number
        If regex_checker(bank_accounts(ba).Item("account_number"), account_number_checker) = False Then
            ColourCell (column_headers("bank_account_" & ba & "_account_number"))
        End If
    Next ba
End Function
Function check_users()
    Dim u As Integer
    For u = 1 To users.Count
        If check_user_role(u) = False Then
            ColourCell (column_headers("user_" & u & "_role"))
        End If

        If check_user_permission_level(u) = False Then
            ColourCell (column_headers("user_" & u & "_permission_level"))
        End If

        If check_email_address(users(u).Item("email")) = False Then
            ColourCell (column_headers("user_" & u & "_email"))
        End If
    Next u
End Function
Function check_ni_number()
    Dim i As Integer
    For i = 1 To users.Count
        If users(i).Item("ni_number") = "" Then
        Else
            If niChecker.Test((users(i).Item("ni_number"))) = False Then
                ColourCell (column_headers("user_" & i & "_ni_number"))
            End If
        End If
    Next i
End Function
Function check_registration_number() As Boolean
    If registration_number = "" Then
        check_registration_number = True
    Else
        If company_number_checker_scotland.Test(registration_number) = True Or company_number_checker_eng_and_wales.Test(registration_number) = True Then
            check_registration_number = True
        Else
            check_registration_number = False
        End If
    End If
End Function
Function check_email_address(email_address_string As String) As Boolean
    If email_address_string = "" Then
        check_email_address = True
    Else
        check_email_address = email_address_checker.Test(email_address_string)
    End If
End Function

' Function regex_checker(string_to_check As String, regex As RegExp) As Boolean
'     If string_to_check = "" Then
'         regex_checker = True
'     Else
'         regex_checker = regex.Test(string_to_check)
'     End If
' End Function
Function check_unique_bank_account_names()
    ' TODO
End Function
Function check_initial_vat_frs_type_index() As Boolean
    check_initial_vat_frs_type_index = False
    If IsNumeric(initial_vat_frs_type_index) Then
        If initial_vat_frs_type_index > 0 And initial_vat_frs_type_index < 55 And Int(initial_vat_frs_type_index) = initial_vat_frs_type_index Then
            check_initial_vat_frs_type_index = True
        End If
    End If
    If initial_vat_frs_type_index = "" Then
        check_initial_vat_frs_type_index = True
    End If
End Function
Sub check_required_fields_bas_and_users()
    If name = "" Then
        Call ColourCell(column_headers("name"))
    End If

    Dim ba As Integer
    For ba = 1 To bank_accounts.Count
        If (bank_accounts(ba).Item("Type") <> "" Or _
                bank_accounts(ba).Item("sort_code") <> "" Or _
                bank_accounts(ba).Item("account_number") <> "" Or _
                bank_accounts(ba).Item("email") <> "" Or _
                bank_accounts(ba).Item("opening_balance") <> "") And _
                (bank_accounts(ba).Item("name") = "") Then
            Call ColourCell(column_headers("bank_account_" & ba & "_name"))
        End If
    Next ba

    Dim u As Integer
    For u = 1 To users.Count
        If (users(u).Item("role") <> "" Or _
                users(u).Item("permission_level") <> "" Or _
                users(u).Item("ni_number") <> "" Or _
                users(u).Item("capital_opening_balance") <> "" Or _
                users(u).Item("directors_loan_opening_balance") <> "" Or _
                users(u).Item("expense_opening_balance") <> "" Or _
                users(u).Item("salary_opening_balance") <> "") And _
                (users(u).Item("first_name") = "" Or _
                users(u).Item("last_name") = "" Or _
                users(u).Item("email") = "") Then
            Call ColourCell(column_headers("user_" & u & "_first_name"))
            Call ColourCell(column_headers("user_" & u & "_last_name"))
            Call ColourCell(column_headers("user_" & u & "_email"))
        End If
    Next u

End Sub

