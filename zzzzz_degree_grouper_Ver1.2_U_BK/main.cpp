#include <windows.h>
#include <commdlg.h>
#include <shlobj.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <iostream>
#include <xlnt/xlnt.hpp>
#include <fstream>
#include <sstream>
#include <algorithm>
#include <filesystem>
#include <map>
#include <set>
#include <thread>

using Row = std::vector<std::string>;
using DataFrame = std::vector<Row>;

// Column mapping from Python code - generate columns from U to BK in steps of 2
// Python: col_letters = [get_column_letter(i) for i in range(start_col, end_col + 1, 2)]
// where start_col = column_index_from_string('U') = 21, end_col = column_index_from_string('BK') = 63
std::vector<std::string> col_order = {
    "U", "W", "Y", "AA", "AC", "AE", "AG", "AI", "AK", "AM", "AO", "AQ",
    "AS", "AU", "AW", "AY", "BA", "BC", "BE", "BG", "BI", "BK"
};

std::map<std::string, int> col_num = {
    {"U", 21}, {"W", 23}, {"Y", 25}, {"AA", 27}, {"AC", 29}, {"AE", 31},
    {"AG", 33}, {"AI", 35}, {"AK", 37}, {"AM", 39}, {"AO", 41}, {"AQ", 43},
    {"AS", 45}, {"AU", 47}, {"AW", 49}, {"AY", 51}, {"BA", 53}, {"BC", 55},
    {"BE", 57}, {"BG", 59}, {"BI", 61}, {"BK", 63}
};

class CSVManager {
public:
    static DataFrame read(const std::wstring& filename) {
        std::wstring ext = getExtension(filename);
        if (ext == L".csv") {
            return readCSVFile(filename);
        }
        else if (ext == L".xlsx") {
            return readXLSXFile(filename);
        }
        else {
            throw std::runtime_error("Unsupported file type");
        }
    }

    static void write(const DataFrame& data, const std::wstring& filename) {
        std::wstring ext = getExtension(filename);
        if (ext == L".csv") {
            writeCSVFile(data, filename);
        }
        else if (ext == L".xlsx") {
            writeXLSXFile(data, filename);
        }
        else {
            throw std::runtime_error("Unsupported file type");
        }
    }

    static std::string ws2s(const std::wstring& wstr) {
        int size_needed = WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.size(), NULL, 0, NULL, NULL);
        std::string strTo(size_needed, 0);
        WideCharToMultiByte(CP_UTF8, 0, wstr.c_str(), (int)wstr.size(), &strTo[0], size_needed, NULL, NULL);
        return strTo;
    }
    
    static std::wstring s2ws(const std::string& str) {
        int size_needed = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), (int)str.size(), NULL, 0);
        std::wstring wstrTo(size_needed, 0);
        MultiByteToWideChar(CP_UTF8, 0, str.c_str(), (int)str.size(), &wstrTo[0], size_needed);
        return wstrTo;
    }

private:
    static std::wstring getExtension(const std::wstring& filename) {
        size_t pos = filename.find_last_of(L'.');
        if (pos == std::wstring::npos) return L"";
        std::wstring ext = filename.substr(pos);
        std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
        return ext;
    }

    static DataFrame readCSVFile(const std::wstring& filename) {
        DataFrame data;
        std::wifstream file(filename.c_str());
        file.imbue(std::locale::classic());
        if (!file.is_open()) {
            throw std::runtime_error("Cannot open file");
        }
        std::wstring line;
        while (std::getline(file, line)) {
            Row row;
            std::wstringstream ss(line);
            std::wstring cell;
            while (std::getline(ss, cell, L',')) {
                // Remove quotes and trim whitespace
                cell.erase(std::remove(cell.begin(), cell.end(), L'"'), cell.end());
                cell.erase(0, cell.find_first_not_of(L" \t"));
                cell.erase(cell.find_last_not_of(L" \t") + 1);
                row.push_back(ws2s(cell));
            }
            if (!row.empty()) {
                data.push_back(row);
            }
        }
        return data;
    }

    static DataFrame readXLSXFile(const std::wstring& filename) {
        DataFrame data;
        xlnt::workbook wb;
        wb.load(ws2s(filename));
        auto ws = wb.active_sheet();
        
        // Read all rows (equivalent to pandas read_excel with header=None)
        // Get the used range to ensure we read all data
        auto range = ws.calculate_dimension();
        for (auto row : ws.rows(false)) {
            Row row_data;
            for (auto cell : row) {
                // Convert cell to string, handling empty cells properly
                if (cell.has_value()) {
                    row_data.push_back(cell.to_string());
                } else {
                    row_data.push_back(""); // Empty cell
                }
            }
            data.push_back(row_data);
        }
        return data;
    }

    static void writeCSVFile(const DataFrame& data, const std::wstring& filename) {
        std::wofstream file(filename.c_str());
        file.imbue(std::locale::classic());
        if (!file.is_open()) {
            throw std::runtime_error("Cannot create file");
        }
        for (const auto& row : data) {
            for (size_t i = 0; i < row.size(); ++i) {
                if (i > 0) file << L",";
                file << s2ws(row[i]);
            }
            file << L"\n";
        }
    }

    static void writeXLSXFile(const DataFrame& data, const std::wstring& filename) {
        xlnt::workbook wb;
        auto ws = wb.active_sheet();
        
        // Write data without headers (equivalent to pandas to_excel with index=False, header=None)
        for (size_t i = 0; i < data.size(); ++i) {
            for (size_t j = 0; j < data[i].size(); ++j) {
                ws.cell(static_cast<uint32_t>(j + 1), static_cast<uint32_t>(i + 1)).value(data[i][j]);
            }
        }
        wb.save(ws2s(filename));
    }
};

// Global handles for GUI controls
HWND hMainWindow;
HWND hExcelEntry;
HWND hTxtEntry;
HWND hProcessButton;
HWND hStatusText;
std::map<std::string, HWND> hCheckboxes;

// Function declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void OnBrowseExcel();
void OnBrowseTxt();
void OnProcess();
std::wstring OpenFileDialog(const wchar_t* filter);
std::wstring OpenFolderDialog();

// Helper function to map values to ranges (equivalent to Python's map_to_range)
std::string mapToRange(const std::string& val, const std::vector<std::string>& groupList) {
    try {
        int intVal = std::stoi(val);
        for (const auto& group : groupList) {
            size_t dashPos = group.find('-');
            if (dashPos != std::string::npos) {
                int start = std::stoi(group.substr(0, dashPos));
                int end = std::stoi(group.substr(dashPos + 1));
                if (start <= intVal && intVal <= end) {
                    return group;
                }
            }
        }
    }
    catch (...) {
        // If conversion fails, return original value (same as Python)
        return val;
    }
    return val; // Return original value if no range matches
}

// Main processing logic
void ProcessFile() {
    try {
        wchar_t excel_path[260];
        wchar_t txt_path[260];
        GetWindowTextW(hExcelEntry, excel_path, 260);
        GetWindowTextW(hTxtEntry, txt_path, 260);
        
        if (wcslen(excel_path) == 0 || wcslen(txt_path) == 0) {
            MessageBoxW(hMainWindow, L"Please select both Excel file and TXT file.", L"Error", MB_OK | MB_ICONERROR);
            return;
        }

        SetWindowTextW(hStatusText, L"Reading Excel file...");
        DataFrame df = CSVManager::read(excel_path);

        SetWindowTextW(hStatusText, L"Reading TXT file...");
        std::vector<std::string> groupList;
        std::wifstream txtFile(txt_path);
        txtFile.imbue(std::locale::classic());
        if (!txtFile.is_open()) {
            MessageBoxW(hMainWindow, L"Cannot open TXT file!", L"Error", MB_OK | MB_ICONERROR);
            return;
        }
        
        std::wstring line;
        while (std::getline(txtFile, line)) {
            std::string trimmed = CSVManager::ws2s(line);
            // Trim whitespace
            trimmed.erase(0, trimmed.find_first_not_of(" \t"));
            trimmed.erase(trimmed.find_last_not_of(" \t") + 1);
            if (!trimmed.empty()) {
                groupList.push_back(trimmed);
            }
        }

        // Get selected columns
        std::vector<std::string> selectedCols;
        for (const auto& col : col_order) {
            if (hCheckboxes.find(col) != hCheckboxes.end()) {
                if (SendMessageW(hCheckboxes[col], BM_GETCHECK, 0, 0) == BST_CHECKED) {
                    selectedCols.push_back(col);
                }
            }
        }

        if (selectedCols.empty()) {
            MessageBoxW(hMainWindow, L"Please select at least one column.", L"Error", MB_OK | MB_ICONERROR);
            return;
        }

        SetWindowTextW(hStatusText, L"Processing data...");
        
        // Process selected columns
        for (const auto& col : selectedCols) {
            int colIndex = col_num[col] - 1; // Convert to 0-based index (Excel columns are 1-based)
            if (colIndex >= 0) {
                for (auto& row : df) {
                    if (colIndex < static_cast<int>(row.size())) {
                        // Only process if the cell contains a numeric value
                        if (!row[colIndex].empty()) {
                            row[colIndex] = mapToRange(row[colIndex], groupList);
                        }
                    }
                }
            }
        }

        // Generate output filename
        std::wstring outputPath = excel_path;
        size_t dotPos = outputPath.find_last_of(L'.');
        if (dotPos != std::wstring::npos) {
            outputPath = outputPath.substr(0, dotPos);
        }
        
        std::string selectedColsStr;
        for (const auto& col : selectedCols) {
            selectedColsStr += col;
        }
        
        outputPath += L"_" + CSVManager::s2ws(selectedColsStr) + L"_Grouped.xlsx";

        SetWindowTextW(hStatusText, L"Saving file...");
        CSVManager::write(df, outputPath);
        
        std::wstring successMsg = L"File saved to:\n" + outputPath;
        SetWindowTextW(hStatusText, successMsg.c_str());
        MessageBoxW(hMainWindow, successMsg.c_str(), L"Success", MB_OK | MB_ICONINFORMATION);
    }
    catch (const std::exception& e) {
        std::wstring err = L"Error: ";
        err += CSVManager::s2ws(e.what());
        SetWindowTextW(hStatusText, err.c_str());
        MessageBoxW(hMainWindow, err.c_str(), L"Error", MB_OK | MB_ICONERROR);
    }
}

void OnProcess() {
    EnableWindow(hProcessButton, FALSE);
    std::thread([=]() {
        ProcessFile();
        EnableWindow(hProcessButton, TRUE);
    }).detach();
}

// WinMain: Entry point
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
    INITCOMMONCONTROLSEX icex = { sizeof(INITCOMMONCONTROLSEX), ICC_WIN95_CLASSES };
    InitCommonControlsEx(&icex);

    WNDCLASSW wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"DegreeGrouperMainWindow";
    wc.hCursor = LoadCursorW(nullptr, (LPCWSTR)IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    RegisterClassW(&wc);

    hMainWindow = CreateWindowExW(0, wc.lpszClassName, L"Degree Grouper", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_THICKFRAME | WS_MAXIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 600, nullptr, nullptr, hInstance, nullptr);

    ShowWindow(hMainWindow, nCmdShow);
    UpdateWindow(hMainWindow);

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    return 0;
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE: {
        // Excel File Label
        CreateWindowW(L"STATIC", L"Input Excel File:", WS_VISIBLE | WS_CHILD,
            10, 20, 150, 20, hwnd, nullptr, nullptr, nullptr);
        // Excel File Entry
        hExcelEntry = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
            170, 20, 500, 20, hwnd, nullptr, nullptr, nullptr);
        // Browse Excel Button
        CreateWindowW(L"BUTTON", L"Browse", WS_VISIBLE | WS_CHILD,
            690, 20, 80, 20, hwnd, (HMENU)1, nullptr, nullptr);

        // TXT File Label
        CreateWindowW(L"STATIC", L"Groups TXT File:", WS_VISIBLE | WS_CHILD,
            10, 50, 150, 20, hwnd, nullptr, nullptr, nullptr);
        // TXT File Entry
        hTxtEntry = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
            170, 50, 500, 20, hwnd, nullptr, nullptr, nullptr);
        // Browse TXT Button
        CreateWindowW(L"BUTTON", L"Browse", WS_VISIBLE | WS_CHILD,
            690, 50, 80, 20, hwnd, (HMENU)2, nullptr, nullptr);

        // Column Selection Label
        CreateWindowW(L"STATIC", L"Select Columns:", WS_VISIBLE | WS_CHILD,
            10, 90, 150, 20, hwnd, nullptr, nullptr, nullptr);

        // Create checkboxes for columns
        int checkboxId = 100;
        int row = 0, col = 0;
        for (const auto& colName : col_order) {
            std::wstring checkboxText = L"Column " + CSVManager::s2ws(colName);
            
            HWND hCheckbox = CreateWindowW(L"BUTTON", checkboxText.c_str(), 
                WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
                170 + col * 120, 90 + row * 25, 110, 20, 
                hwnd, (HMENU)checkboxId, nullptr, nullptr);
            
            hCheckboxes[colName] = hCheckbox;
            
            col++;
            if (col >= 5) { // 5 columns per row
                col = 0;
                row++;
            }
            checkboxId++;
        }

        // Process Button
        hProcessButton = CreateWindowW(L"BUTTON", L"Process", WS_VISIBLE | WS_CHILD,
            350, 90 + (row + 1) * 25, 150, 30, hwnd, (HMENU)3, nullptr, nullptr);

        // Status Text
        hStatusText = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD | SS_LEFT,
            10, 90 + (row + 2) * 25, 760, 50, hwnd, nullptr, nullptr, nullptr);

        break;
    }
    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        switch (wmId) {
        case 1: OnBrowseExcel(); break;
        case 2: OnBrowseTxt(); break;
        case 3: OnProcess(); break;
        }
        break;
    }
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProcW(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

void OnBrowseExcel() {
    std::wstring file = OpenFileDialog(L"Excel files\0*.xlsx;*.xls\0All Files\0*.*\0");
    if (!file.empty()) {
        SetWindowTextW(hExcelEntry, file.c_str());
    }
}

void OnBrowseTxt() {
    std::wstring file = OpenFileDialog(L"TXT files\0*.txt\0All Files\0*.*\0");
    if (!file.empty()) {
        SetWindowTextW(hTxtEntry, file.c_str());
    }
}

std::wstring OpenFileDialog(const wchar_t* filter) {
    OPENFILENAMEW ofn = { 0 };
    wchar_t szFile[260] = { 0 };
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hMainWindow;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = filter;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
    if (GetOpenFileNameW(&ofn)) {
        return szFile;
    }
    return L"";
}

std::wstring OpenFolderDialog() {
    BROWSEINFOW bi = { 0 };
    wchar_t szFolder[260] = { 0 };
    bi.hwndOwner = hMainWindow;
    bi.pszDisplayName = szFolder;
    bi.lpszTitle = L"Select Folder";
    LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
    if (pidl != nullptr) {
        SHGetPathFromIDListW(pidl, szFolder);
        return szFolder;
    }
    return L"";
}
