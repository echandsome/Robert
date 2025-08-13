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
#include <thread>
#include <map>
#include <cmath>

using Row = std::vector<std::string>;
using DataFrame = std::vector<Row>;

// Constants matching Python script
const std::vector<std::string> all_columns = { "Player", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK" };
const std::vector<std::string> daily_cols = { "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK" };
const std::vector<std::string> degree_cols = { "AQ","AS","AU", "AW", "AY", "BA", "BC", "BE", "BG", "BI", "BK" };

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
            writeCSVFile(filename, data);
        }
        else if (ext == L".xlsx") {
            writeXLSXFile(filename, data);
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

    static void writeCSVFile(const std::wstring& filename, const DataFrame& data) {
        std::wofstream file(filename.c_str());
        file.imbue(std::locale::classic());
        if (!file.is_open()) {
            throw std::runtime_error("Cannot create file");
        }
        for (const auto& row : data) {
            for (size_t i = 0; i < row.size(); ++i) {
                if (i > 0) file << L",";
                
                std::wstring cell_value = s2ws(row[i]);
                
                // Check if cell contains comma, quote, or newline - if so, wrap in quotes
                bool needs_quotes = (cell_value.find(L',') != std::wstring::npos || 
                                   cell_value.find(L'"') != std::wstring::npos || 
                                   cell_value.find(L'\n') != std::wstring::npos ||
                                   cell_value.find(L'\r') != std::wstring::npos);
                
                if (needs_quotes) {
                    // Escape existing quotes by doubling them
                    size_t pos = 0;
                    while ((pos = cell_value.find(L'"', pos)) != std::wstring::npos) {
                        cell_value.insert(pos, L"\"");
                        pos += 2;
                    }
                    file << L"\"" << cell_value << L"\"";
                } else {
                    file << cell_value;
                }
            }
            file << L"\n";
        }
    }

    static void writeXLSXFile(const std::wstring& filename, const DataFrame& data) {
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

// Global handles for GUI controls
HWND hMainWindow;
HWND hDailyEntry;
HWND hHistEntry;
HWND hProcessButton;
HWND hStatusText;
HWND hProgressBar;

// Function declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void OnBrowseDaily();
void OnBrowseHist();
void OnProcess();
std::wstring OpenFileDialog();
std::wstring OpenFolderDialog();


// Helper: filter daily data to columns 0 and 41-62 (matching Python script)
DataFrame FilterDailyData(const DataFrame& raw_daily_df) {
    DataFrame filtered_data;
    for (const Row& row : raw_daily_df) {
        Row filtered_row;
        if (!row.empty()) filtered_row.push_back(row[0]); // Player column
        for (int i = 41; i <= 62 && i < static_cast<int>(row.size()); ++i) {
            filtered_row.push_back(row[i]);
        }
        if (!filtered_row.empty()) filtered_data.push_back(filtered_row);
    }
    return filtered_data;
}

// Helper: get column index from column name in daily_df
size_t GetColumnIndex(const std::string& col_name) {
    auto it = std::find(all_columns.begin(), all_columns.end(), col_name);
    if (it != all_columns.end()) {
        return std::distance(all_columns.begin(), it);
    }
    return SIZE_MAX; // Column not found
}

// Helper: parse row to dictionary (matching Python parse_row_to_dict function)
std::map<std::string, std::string> ParseRowToDict(const Row& row) {
    std::map<std::string, std::string> data;
    if (row.size() < 4) return data;
    
    data["Player"] = row[0];
    
    // Process key-value pairs (skip last 4: Count, Total, WinTotal, WinPercent)
    for (size_t i = 1; i + 1 < row.size() - 4; i += 2) {
        if (i + 1 < row.size()) {
            data[row[i]] = row[i + 1];
        }
    }
    
    // Add the last 4 fields
    if (row.size() >= 4) {
        data["Count"] = row[row.size() - 4];
        data["Total"] = row[row.size() - 3];
        data["WinTotal"] = row[row.size() - 2];
        data["WinPercent"] = row[row.size() - 1];
    }
    
    return data;
}

// Helper: check if two values match (matching Python logic)
bool ValuesMatch(const std::string& daily_val, const std::string& hist_val) {
    if (daily_val.empty() || hist_val.empty()) return false;
    
    try {
        return std::stoi(daily_val) == std::stoi(hist_val);
    }
    catch (...) {
        return false;
    }
}



// Main processing logic (matching Python process_files function)
void ProcessMatching(const std::wstring& daily_file, const std::wstring& hist_folder) {
    try {
        SetWindowTextW(hStatusText, L"Reading daily file...");
        DataFrame raw_daily_df = CSVManager::read(daily_file);
        DataFrame daily_df = FilterDailyData(raw_daily_df);
        
        // Debug: Log the data structure
        std::wstring debug_info = L"Daily data: " + std::to_wstring(raw_daily_df.size()) + L" rows, " + 
                                 std::to_wstring(raw_daily_df[0].size()) + L" columns. Filtered: " + 
                                 std::to_wstring(daily_df.size()) + L" rows, " + std::to_wstring(daily_df[0].size()) + L" columns";
        SetWindowTextW(hStatusText, debug_info.c_str());
        
        // Debug: Show column mapping
        if (!daily_df.empty() && !daily_df[0].empty()) {
            std::wstring col_debug = L"Column mapping: 0->Player, 1->" + CSVManager::s2ws(daily_cols[0]) + 
                                    L", 2->" + CSVManager::s2ws(daily_cols[1]) + L", 3->" + CSVManager::s2ws(daily_cols[2]);
            SetWindowTextW(hStatusText, col_debug.c_str());
        }
        
        // Set column labels for daily_df (matching Python script)
        // daily_df columns are: Player, AP, AQ, AR, AS, AT, AU, AV, AW, AX, AY, AZ, BA, BB, BC, BD, BE, BF, BG, BH, BI, BJ, BK
        // This corresponds to original columns: 0, 41, 42, 43, ..., 62

        std::vector<Row> all_matches;
        std::vector<std::pair<size_t, Row>> all_rows;
        
        // Collect all historical rows (matching Python script)
        for (const auto& entry : std::filesystem::directory_iterator(hist_folder)) {
            if (!entry.is_regular_file()) continue;
            std::wstring ext = entry.path().extension().wstring();
            std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
            if (ext != L".csv" && ext != L".xlsx") continue;
            
            std::wstring file_name = entry.path().filename().wstring();
            std::wstring status = L"Reading: " + file_name;
            SetWindowTextW(hStatusText, status.c_str());
            
            DataFrame raw_hist_df = CSVManager::read(entry.path().wstring());
            for (size_t idx = 0; idx < raw_hist_df.size(); ++idx) {
                all_rows.push_back({idx, raw_hist_df[idx]});
            }
        }
        
        // Progress bar setup
        SendMessageW(hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, all_rows.size()));
        SendMessageW(hProgressBar, PBM_SETPOS, 0, 0);
        
        int processed = 0;
        
        // Process each historical row (matching Python multiprocess_rows logic)
        for (const auto& [idx, hist_row] : all_rows) {
            try {
                std::wstring status = L"Processing row " + std::to_wstring(idx) + L"...";
                SetWindowTextW(hStatusText, status.c_str());
                
                std::map<std::string, std::string> row_dict = ParseRowToDict(hist_row);
                
                // For each daily row, check match (matching Python logic)
                for (size_t i = 0; i < daily_df.size(); ++i) {
                    const Row& daily_row = daily_df[i];
                    if (daily_row.empty()) continue;
                    
                    bool is_match = true;
                    

                    
                    // Check each field in historical row
                    for (const auto& [col, hist_val] : row_dict) {
                        if (col == "Count" || col == "Total" || col == "WinTotal" || col == "WinPercent") {
                            continue; // Skip these fields as per Python script
                        }

                        // Skip if historical value is empty or 0 (same logic as Python)
                        if (hist_val.empty() || hist_val == "0") {
                            continue;
                        }
                        
                        std::string daily_val;
                        if (col == "Player") {
                            daily_val = daily_row[0];
                            if (daily_val != hist_val) {
                                is_match = false;
                                break;
                            }
                        } else {
                            // Get the column index using the helper function
                            size_t col_idx = GetColumnIndex(col);
                            if (col_idx == SIZE_MAX || col_idx >= daily_row.size()) {
                                continue; // Column not found or out of bounds
                            }
                            
                            daily_val = daily_row[col_idx];
                            
                            if (daily_val != hist_val) {
                                is_match = false;
                                break;
                            }
                        }
                    }    
                    
                    if (is_match) {
                        // Found match - combine daily and historical data
                        Row matched_row = raw_daily_df[i];
                        matched_row.insert(matched_row.end(), hist_row.begin(), hist_row.end());
                        all_matches.push_back(matched_row);
                    }
                }
                
                processed++;
                SendMessageW(hProgressBar, PBM_SETPOS, processed, 0);
                
            } catch (const std::exception& e) {
                // Continue processing other rows if one fails
                continue;
            }
        }
        
        // Write output (matching Python script output format)
        if (!all_matches.empty()) {
            std::wstring out_path = daily_file.substr(0, daily_file.find_last_of(L'.')) + L"_Matches.csv";
            CSVManager::write(all_matches, out_path);
            std::wstring success_msg = L"Processing finished. Found " + std::to_wstring(all_matches.size()) + L" matches. Output: " + out_path;
            SetWindowTextW(hStatusText, success_msg.c_str());
            MessageBoxW(hMainWindow, success_msg.c_str(), L"Success", MB_OK | MB_ICONINFORMATION);
        }
        else {
            std::wstring no_match_msg = L"NO Matches found... Processed " + std::to_wstring(all_rows.size()) + L" historical rows against " + std::to_wstring(daily_df.size()) + L" daily rows";
            SetWindowTextW(hStatusText, no_match_msg.c_str());
            MessageBoxW(hMainWindow, no_match_msg.c_str(), L"No Results", MB_OK | MB_ICONWARNING);
        }
        SendMessageW(hProgressBar, PBM_SETPOS, 0, 0);
    }
    catch (const std::exception& e) {
        std::wstring err = L"Error: ";
        err += CSVManager::s2ws(e.what());
        SetWindowTextW(hStatusText, err.c_str());
        MessageBoxW(hMainWindow, err.c_str(), L"Error", MB_OK | MB_ICONERROR);
        SendMessageW(hProgressBar, PBM_SETPOS, 0, 0);
    }
    EnableWindow(hProcessButton, TRUE);
}

void OnProcess() {
    wchar_t daily_path[260];
    wchar_t hist_path[260];
    GetWindowTextW(hDailyEntry, daily_path, 260);
    GetWindowTextW(hHistEntry, hist_path, 260);
    if (wcslen(daily_path) == 0 || wcslen(hist_path) == 0) {
        MessageBoxW(hMainWindow, L"Please select both daily file and historical folder.", L"Error", MB_OK | MB_ICONERROR);
        return;
    }
    
    EnableWindow(hProcessButton, FALSE);
    std::thread([=]() {
        ProcessMatching(daily_path, hist_path);
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
    wc.lpszClassName = L"ZmatcherMainWindow";
    wc.hCursor = LoadCursorW(nullptr, (LPCWSTR)IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    RegisterClassW(&wc);

    hMainWindow = CreateWindowExW(0, wc.lpszClassName, L"Zmatcher Non-Coloring Processor", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_THICKFRAME | WS_MAXIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 900, 400, nullptr, nullptr, hInstance, nullptr);

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
        // Daily File Label
        CreateWindowW(L"STATIC", L"Daily File:", WS_VISIBLE | WS_CHILD,
            10, 20, 162, 20, hwnd, nullptr, nullptr, nullptr);
        // Daily File Entry
        hDailyEntry = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
            165, 20, 543, 20, hwnd, nullptr, nullptr, nullptr);
        // Browse Daily Button
        CreateWindowW(L"BUTTON", L"Browse", WS_VISIBLE | WS_CHILD,
            735, 20, 80, 20, hwnd, (HMENU)1, nullptr, nullptr);
        // Historical Folder Label
        CreateWindowW(L"STATIC", L"Historical % Input Folder:", WS_VISIBLE | WS_CHILD,
            10, 50, 162, 20, hwnd, nullptr, nullptr, nullptr);
        // Historical Folder Entry
        hHistEntry = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
            165, 50, 543, 20, hwnd, nullptr, nullptr, nullptr);
        // Browse Hist Button
        CreateWindowW(L"BUTTON", L"Browse", WS_VISIBLE | WS_CHILD,
            735, 50, 80, 20, hwnd, (HMENU)2, nullptr, nullptr);
        // Process Button
        hProcessButton = CreateWindowW(L"BUTTON", L"Process", WS_VISIBLE | WS_CHILD,
            375, 90, 150, 30, hwnd, (HMENU)3, nullptr, nullptr);
        // Status Text
        hStatusText = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD | SS_LEFT,
            10, 140, 870, 50, hwnd, nullptr, nullptr, nullptr);
        // Progress Bar
        hProgressBar = CreateWindowW(PROGRESS_CLASSW, L"", WS_VISIBLE | WS_CHILD,
            10, 200, 870, 20, hwnd, nullptr, nullptr, nullptr);
        break;
    }
    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        switch (wmId) {
        case 1: OnBrowseDaily(); break;
        case 2: OnBrowseHist(); break;
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

void OnBrowseDaily() {
    std::wstring file = OpenFileDialog();
    if (!file.empty()) {
        SetWindowTextW(hDailyEntry, file.c_str());
    }
}

void OnBrowseHist() {
    std::wstring folder = OpenFolderDialog();
    if (!folder.empty()) {
        SetWindowTextW(hHistEntry, folder.c_str());
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

std::wstring OpenFolderDialog() {
    BROWSEINFOW bi = { 0 };
    wchar_t szFolder[260] = { 0 };
    bi.hwndOwner = hMainWindow;
    bi.pszDisplayName = szFolder;
    bi.lpszTitle = L"Select Historical Folder";
    LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
    if (pidl != nullptr) {
        SHGetPathFromIDListW(pidl, szFolder);
        return szFolder;
    }
    return L"";
}
