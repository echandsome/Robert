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
#include <thread>
#include <cmath>

using Row = std::vector<std::string>;
using DataFrame = std::vector<Row>;

class CSVManager {
public:
    static DataFrame read(const std::wstring& filename) {
        std::wstring ext = getExtension(filename);
        if (ext == L".csv") {
            return readCSVFile(filename);
        }
        else if (ext == L".xlsx" || ext == L".xls") {
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
        else if (ext == L".xlsx" || ext == L".xls") {
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
        for (auto row : ws.rows(false)) {
            Row row_data;
            for (auto cell : row) {
                row_data.push_back(cell.to_string());
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
        for (size_t i = 0; i < data.size(); ++i) {
            for (size_t j = 0; j < data[i].size(); ++j) {
                ws.cell(static_cast<uint32_t>(j + 1), static_cast<uint32_t>(i + 1)).value(data[i][j]);
            }
        }
        wb.save(ws2s(filename));
    }
};

// Function to convert column letter to index (A=0, B=1, etc.)
int col_letter_to_index(const std::string& letter) {
    std::string upper_letter = letter;
    std::transform(upper_letter.begin(), upper_letter.end(), upper_letter.begin(), ::toupper);
    
    int total = 0;
    for (int i = 0; i < static_cast<int>(upper_letter.length()); ++i) {
        char c = upper_letter[upper_letter.length() - 1 - i];
        if (c >= 'A' && c <= 'Z') {
            total += (c - 'A' + 1) * static_cast<int>(std::pow(26, i));
        }
    }
    return total - 1;
}

// Function to split file by column
void split_file_by_column(const std::wstring& file_path, const std::string& column_letter) {
    // Read the file
    DataFrame df = CSVManager::read(file_path);
    
    // Convert column letter to index
    int col_index = col_letter_to_index(column_letter);
    if (col_index >= static_cast<int>(df[0].size())) {
        throw std::runtime_error("Column letter exceeds available columns in the file.");
    }
    
    // Create outputs folder in same directory
    std::wstring base_dir = file_path.substr(0, file_path.find_last_of(L'\\'));
    std::wstring output_dir = base_dir + L"\\outputs";
    std::filesystem::create_directories(output_dir);
    
    // Group data by the specified column
    std::map<std::string, DataFrame> grouped_data;
    for (const auto& row : df) {
        if (col_index < static_cast<int>(row.size())) {
            std::string group_key = row[col_index];
            grouped_data[group_key].push_back(row);
        }
    }
    
    // Get file extension
    std::wstring ext = file_path.substr(file_path.find_last_of(L'.'));
    std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
    
    // Write separate files for each group
    for (const auto& group : grouped_data) {
        std::string safe_name = group.first;
        // Replace invalid characters for filename
        std::replace(safe_name.begin(), safe_name.end(), '/', '_');
        std::replace(safe_name.begin(), safe_name.end(), '\\', '_');
        
        std::wstring output_filename = L"split_" + CSVManager::s2ws(column_letter) + L"_" + 
                                     CSVManager::s2ws(safe_name) + ext;
        std::wstring output_path = output_dir + L"\\" + output_filename;
        
        CSVManager::write(group.second, output_path);
    }
}

// Global handles for GUI controls
HWND hMainWindow;
HWND hFileEntry;
HWND hColumnEntry;
HWND hSplitButton;
HWND hStatusText;

// Function declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void OnBrowseFile();
void OnSplit();
std::wstring OpenFileDialog();

// Main processing logic
void ProcessSplit(const std::wstring& file_path, const std::string& column_letter) {
    try {
        SetWindowTextW(hStatusText, L"Processing...");
        
        split_file_by_column(file_path, column_letter);
        
        SetWindowTextW(hStatusText, L"Split complete! Files saved in 'outputs' folder.");
        MessageBoxW(hMainWindow, L"Split complete! Files saved in 'outputs' folder.", L"Success", MB_OK | MB_ICONINFORMATION);
    }
    catch (const std::exception& e) {
        std::wstring err = L"Error: ";
        err += CSVManager::s2ws(e.what());
        SetWindowTextW(hStatusText, err.c_str());
        MessageBoxW(hMainWindow, err.c_str(), L"Error", MB_OK | MB_ICONERROR);
    }
}

void OnSplit() {
    wchar_t file_path[260];
    wchar_t column_letter[10];
    GetWindowTextW(hFileEntry, file_path, 260);
    GetWindowTextW(hColumnEntry, column_letter, 10);
    
    if (wcslen(file_path) == 0 || wcslen(column_letter) == 0) {
        MessageBoxW(hMainWindow, L"Please select a file and enter a column letter.", L"Error", MB_OK | MB_ICONERROR);
        return;
    }
    
    std::string column_str = CSVManager::ws2s(column_letter);
    // Trim whitespace
    column_str.erase(0, column_str.find_first_not_of(" \t"));
    column_str.erase(column_str.find_last_not_of(" \t") + 1);
    
    if (column_str.empty()) {
        MessageBoxW(hMainWindow, L"Please enter a valid column letter.", L"Error", MB_OK | MB_ICONERROR);
        return;
    }
    
    EnableWindow(hSplitButton, FALSE);
    std::thread([=]() {
        ProcessSplit(file_path, column_str);
        EnableWindow(hSplitButton, TRUE);
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
    wc.lpszClassName = L"FileSplitterMainWindow";
    wc.hCursor = LoadCursorW(nullptr, (LPCWSTR)IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    RegisterClassW(&wc);

    hMainWindow = CreateWindowExW(0, wc.lpszClassName, L"Universal File Splitter", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_THICKFRAME | WS_MAXIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 500, 250, nullptr, nullptr, hInstance, nullptr);

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
        // File Selection Button
        CreateWindowW(L"BUTTON", L"Select CSV/XLSX File", WS_VISIBLE | WS_CHILD,
            10, 10, 150, 30, hwnd, (HMENU)1, nullptr, nullptr);
        
        // File Path Entry
        hFileEntry = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_READONLY,
            170, 10, 300, 30, hwnd, nullptr, nullptr, nullptr);
        
        // Column Letter Label
        CreateWindowW(L"STATIC", L"Enter Column Letter to Split By (e.g., A or P):", WS_VISIBLE | WS_CHILD,
            10, 60, 300, 20, hwnd, nullptr, nullptr, nullptr);
        
        // Column Letter Entry
        hColumnEntry = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
            10, 85, 100, 25, hwnd, nullptr, nullptr, nullptr);
        
        // Split Button
        hSplitButton = CreateWindowW(L"BUTTON", L"Split File", WS_VISIBLE | WS_CHILD,
            10, 120, 150, 35, hwnd, (HMENU)2, nullptr, nullptr);
        
        // Status Text
        hStatusText = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD | SS_LEFT,
            10, 170, 460, 50, hwnd, nullptr, nullptr, nullptr);
        break;
    }
    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        switch (wmId) {
        case 1: OnBrowseFile(); break;
        case 2: OnSplit(); break;
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

void OnBrowseFile() {
    std::wstring file = OpenFileDialog();
    if (!file.empty()) {
        SetWindowTextW(hFileEntry, file.c_str());
    }
}

std::wstring OpenFileDialog() {
    OPENFILENAMEW ofn = { 0 };
    wchar_t szFile[260] = { 0 };
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hMainWindow;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = L"Excel and CSV files\0*.xlsx;*.xls;*.csv\0All Files\0*.*\0";
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;
    if (GetOpenFileNameW(&ofn)) {
        return szFile;
    }
    return L"";
}
