#include <iostream>
#include <searchapi.h>
#include <atlbase.h>

// lib imports
#pragma comment(lib, "searchsdk.lib")

// NOTES: Check predefined debug main arguments: Properties -> Debugging -> Command Arguments

std::wstring ReplaceAll(std::wstring str, const std::wstring& from, const std::wstring& to) {
    size_t start_pos = 0;
    while ((start_pos = str.find(from, start_pos)) != std::wstring::npos) {
        str.replace(start_pos, from.length(), to);
        start_pos += to.length();
    }
    return str;
}

void ExecuteOLEDBQuery(const LPWSTR* pSqlQuery);

void PerformFileSearch(const std::wstring& pattern, const std::wstring& userQuery) {
    CComPtr<ISearchManager> pSearchManager;
    HRESULT hr = CoInitialize(nullptr);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to initialize COM." << std::endl;
        return;
    }

    hr = CoCreateInstance(CLSID_CSearchManager, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pSearchManager));
    if (FAILED(hr)) {
        std::wcerr << L"Failed to create ISearchManager instance." << std::endl;
        return;
    }

    CComPtr<ISearchCatalogManager> pSearchCatalogManager;
    hr = pSearchManager->GetCatalog(L"SystemIndex", &pSearchCatalogManager);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get SystemIndex catalog." << std::endl;
        return;
    }

    CComPtr<ISearchQueryHelper> pQueryHelper;
    hr = pSearchCatalogManager->GetQueryHelper(&pQueryHelper);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to get ISearchQueryHelper." << std::endl;
        return;
    }


    pQueryHelper->put_QueryMaxResults(10);
    pQueryHelper->put_QuerySelectColumns(L"System.ItemPathDisplay, System.Size");
    pQueryHelper->put_QuerySorting(L"System.DateModified DESC");


    std::wstring queryWhereRestrictions = L"AND scope='file:'";

    if (pattern != L"*") {
        std::wstring modifiedPattern = ReplaceAll(pattern, L"*", L"%");
        modifiedPattern = ReplaceAll(modifiedPattern, L"?", L"_");

        if (modifiedPattern.find(L"%") != std::wstring::npos || modifiedPattern.find(L"_") != std::wstring::npos) {
            queryWhereRestrictions += L" AND System.ItemPathDisplay LIKE '" + modifiedPattern + L"' ";
        }
        else {
            queryWhereRestrictions += L" AND Contains(System.ItemPathDisplay, '" + modifiedPattern + L"') ";
        }
    }

    hr = pQueryHelper->put_QueryWhereRestrictions(queryWhereRestrictions.c_str());
    if (FAILED(hr)) {
        std::wcerr << L"Failed to set query restrictions." << std::endl;
        return;
    }

    size_t bufferSize = 1024;
    LPWSTR pSqlQuery = (LPWSTR)CoTaskMemAlloc(bufferSize * sizeof(WCHAR));
    if (!pSqlQuery) {
        std::wcout << L"Memory allocation failed!" << std::endl;
        return;
    }
    hr = pQueryHelper->GenerateSQLFromUserQuery(userQuery.c_str(), &pSqlQuery);
    if (FAILED(hr)) {
        std::wcerr << L"Failed to generate SQL query." << std::endl;
        return;
    }

    std::wcout << L"Generated SQL query: " << std::endl << pSqlQuery << std::endl;
    ExecuteOLEDBQuery(&pSqlQuery);
    CoUninitialize();
}


void ExecuteOLEDBQuery(const LPWSTR* pSqlQuery) {

    CComPtr<IDBInitialize> pDBInitialize;
    CComPtr<IDBCreateSession> pDBCreateSession;
    CComPtr<IDBCreateCommand> pDBCreateCommand;
    CComPtr<ICommandText> pCommandText;
    CComPtr<IRowset> pRowset;

    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"Search.CollatorDSO", &clsid);
    if (FAILED(hr)) {
        std::wcout << L"Failed to get CLSID for Search.CollatorDSO. HRESULT: " << hr << std::endl;
        return;
    }
    hr = CoCreateInstance(clsid, nullptr, CLSCTX_INPROC_SERVER, IID_IDBInitialize, (void**)&pDBInitialize);
    if (FAILED(hr))
    {
		std::wcout << L"Failed to create IDBInitialize instance. HRESULT: " << hr << std::endl;
		return;
    }

    hr = pDBInitialize->Initialize();
    if (FAILED(hr)) {
        std::wcout << L"Failed to initialize OLEDB. HRESULT: " << hr << std::endl;
        return;
    }
    hr = pDBInitialize->QueryInterface(IID_IDBCreateSession, (void**)&pDBCreateSession);
    if (FAILED(hr))
    {
		std::wcout << L"Failed to get IDBCreateSession interface. HRESULT: " << hr << std::endl;
		return;
    }

    hr = pDBCreateSession->CreateSession(nullptr, IID_IDBCreateCommand, (IUnknown**)&pDBCreateCommand);
    if (FAILED(hr))
    {
		std::wcout << L"Failed to get IDBCreateCommand interface. HRESULT: " << hr << std::endl;
		return;
    }

    hr = pDBCreateCommand->CreateCommand(nullptr, IID_ICommandText, (IUnknown**)&pCommandText);
    if (FAILED(hr))
    {
		std::wcout << L"Failed to get ICommandText interface. HRESULT: " << hr << std::endl;
		return;
    }

    hr = pCommandText->SetCommandText(DBGUID_DEFAULT, *pSqlQuery);
    if (FAILED(hr))
    {
		std::wcout << L"Failed to set command text. HRESULT: " << hr << std::endl;
		return;
    }

    hr = pCommandText->Execute(nullptr, IID_IRowset, nullptr, nullptr, (IUnknown**)&pRowset);
    if (FAILED(hr)) {
        std::wcout << L"Execute failed with HRESULT: 0x" << std::hex << hr << std::endl;
        return;
    }

    // Process rows
    struct RowData {
        wchar_t fileName[260];
        ULONGLONG fileSize;
    };

    HROW hRow = 0;
    HROW* pRows = &hRow;

    CComPtr<IAccessor> pAccessor;
    DBBINDING binding[2] = {};
    RowData rowData = {};
    HACCESSOR hAccessor = 0;

    binding[0].iOrdinal = 1; // Column 1 (File Name)
    binding[0].dwPart = DBPART_VALUE;
    binding[0].wType = DBTYPE_WSTR;
    binding[0].cbMaxLen = sizeof(rowData.fileName);
    binding[0].obValue = offsetof(RowData, fileName);

    binding[1].iOrdinal = 2; // Column 2 (File Size)
    binding[1].dwPart = DBPART_VALUE;
    binding[1].wType = DBTYPE_UI8;
    binding[1].cbMaxLen = sizeof(rowData.fileSize);
    binding[1].obValue = offsetof(RowData, fileSize);

    hr = pRowset->QueryInterface(IID_IAccessor, (void**)&pAccessor);
    if (FAILED(hr)) {
        std::wcout << L"Error: QueryInterface failed with HRESULT " << hr << std::endl;
        return;
    }

    hr = pAccessor->CreateAccessor(DBACCESSOR_ROWDATA, 2, binding, 0, &hAccessor, nullptr);
    if (FAILED(hr)) {
        std::wcout << L"Error: CreateAccessor failed with HRESULT " << hr << std::endl;
        return;
    }

    DBCOUNTITEM cRows = 0;
    while (SUCCEEDED(hr = pRowset->GetNextRows(0, 0, 1, &cRows, &pRows)) && cRows > 0) {
        if (FAILED(hr)) {
            std::wcout << L"Error: GetNextRows failed with HRESULT " << hr << std::endl;
            break;
        }
        hr = pRowset->GetData(hRow, hAccessor, &rowData);
        if (FAILED(hr)) {
            std::wcout << L"Error: GetData failed with HRESULT " << hr << std::endl;
            break;
        }

        std::wcout << L"File Name: " << rowData.fileName << L", Size: " << rowData.fileSize << L" bytes" << std::endl;
        pRowset->ReleaseRows(1, pRows, nullptr, nullptr, nullptr);
    }
}

int wmain(int argc, wchar_t* argv[]) {
    if (argc < 2) {
        std::wcout << L"Usage: ds [file search path pattern] [userQuery]" << std::endl;
        return 1; // Error code
    }

    std::wstring pattern = argv[1];
    std::wstring userQuery;
    for (int i = 2; i < argc; i++) {
        userQuery += argv[i];
        userQuery += L" ";
    }

    std::wcout << L"Starting search with path pattern: " << pattern << L" and query: " << userQuery << std::endl;

    PerformFileSearch(pattern, userQuery);

    std::wcout << L"Search completed successfully." << std::endl;
    return 0;
}