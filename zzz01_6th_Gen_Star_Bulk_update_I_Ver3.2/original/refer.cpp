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

using Row = std::vector<std::string>;
using DataFrame = std::vector<Row>;

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

// Global handles for GUI controls
HWND hMainWindow;
HWND hDailyEntry;
HWND hHistEntry;
HWND hRadioCSV;
HWND hRadioExcel;
HWND hProcessButton;
HWND hStatusText;
HWND hProgressBar;
HWND hThreadNumEntry;
int THREAD_NUM = 8;

// Function declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void OnBrowseDaily();
void OnBrowseHist();
void OnProcess();
std::wstring OpenFileDialog();
std::wstring OpenFolderDialog();

// Helper: filter daily data to columns 0 and 41-62
DataFrame FilterDailyData(const DataFrame& raw_daily_df) {
    DataFrame filtered_data;
    for (const Row& row : raw_daily_df) {
        Row filtered_row;
        if (!row.empty()) filtered_row.push_back(row[0]);
        for (int i = 41; i <= 62 && i < static_cast<int>(row.size()); ++i) {
            filtered_row.push_back(row[i]);
        }
        if (!filtered_row.empty()) filtered_data.push_back(filtered_row);
    }
    return filtered_data;
}

// Helper: get output format from radio buttons
std::wstring GetOutputFormat() {
    if (SendMessageW(hRadioCSV, BM_GETCHECK, 0, 0) == BST_CHECKED) return L"csv";
    return L"xlsx";
}

// Main processing logic
void ProcessMatching(const std::wstring& daily_file, const std::wstring& hist_folder, const std::wstring& output_format) {
    try {
        SetWindowTextW(hStatusText, L"Reading daily file...");
        DataFrame raw_daily_df = CSVManager::read(daily_file);
        DataFrame daily_df = FilterDailyData(raw_daily_df);

        std::vector<Row> all_matches;
        int file_count = 0;
        for (const auto& entry : std::filesystem::directory_iterator(hist_folder)) {
            if (!entry.is_regular_file()) continue;
            std::wstring ext = entry.path().extension().wstring();
            std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
            if (ext != L".csv" && ext != L".xlsx") continue;
            file_count++;
        }
        // Progress bar setup
        SendMessageW(hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, file_count));
        SendMessageW(hProgressBar, PBM_SETPOS, 0, 0);
        int processed = 0;
        for (const auto& entry : std::filesystem::directory_iterator(hist_folder)) {
            if (!entry.is_regular_file()) continue;
            std::wstring ext = entry.path().extension().wstring();
            std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
            if (ext != L".csv" && ext != L".xlsx") continue;
            std::wstring file_name = entry.path().filename().wstring();
            std::wstring status = L"Processing: " + file_name;
            SetWindowTextW(hStatusText, status.c_str());
            DataFrame raw_hist_df = CSVManager::read(entry.path().wstring());
            for (size_t idx = 0; idx < raw_hist_df.size(); ++idx) {
                const Row& hist_row = raw_hist_df[idx];
                if (hist_row.size() < 5) continue;
                std::string player = hist_row[0];
                std::string degrees_str = hist_row[1];
                std::string degrees_count_str = hist_row[2];
                std::string win_percent = hist_row[4];
                // Parse degree columns
                std::vector<std::string> degree_cols;
                for (size_t i = 0; i + 1 < degrees_str.size(); i += 2) {
                    degree_cols.push_back(degrees_str.substr(i, 2));
                }
                int hist_degrees_count = 0;
                try { hist_degrees_count = std::stoi(degrees_count_str); }
                catch (...) { continue; }
                // For each daily row, check match
                for (size_t i = 0; i < daily_df.size(); ++i) {
                    const Row& daily_row = daily_df[i];
                    if (daily_row.empty()) continue;
                    if (daily_row[0] != player) continue;
                    int daily_degree_count = 0;
                    for (const auto& col : degree_cols) {
                        // Find column index in daily_cols
                        static const std::vector<std::string> daily_cols = { "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK" };
                        auto it = std::find(daily_cols.begin(), daily_cols.end(), col);
                        if (it == daily_cols.end()) continue;
                        size_t col_idx = std::distance(daily_cols.begin(), it) + 1; // +1 for Player
                        if (col_idx < daily_row.size()) {
                            try { daily_degree_count += std::stoi(daily_row[col_idx]); }
                            catch (...) {}
                        }
                    }
                    if (daily_degree_count != hist_degrees_count) continue;
                    // Matched
                    Row matched_row = raw_daily_df[i];
                    matched_row.insert(matched_row.end(), hist_row.begin(), hist_row.end());
                    all_matches.push_back(matched_row);
                }
            }
            processed++;
            std::wstring done = L"Processed: " + file_name + L" (" + std::to_wstring(processed) + L"/" + std::to_wstring(file_count) + L")";
            SetWindowTextW(hStatusText, done.c_str());
            SendMessageW(hProgressBar, PBM_SETPOS, processed, 0);
        }
        // Write output
        if (!all_matches.empty()) {
            std::wstring out_path = daily_file.substr(0, daily_file.find_last_of(L'.')) + L"_Matches." + output_format;
            CSVManager::write(all_matches, out_path);
            SetWindowTextW(hStatusText, (L"Processing finished. Output: " + out_path).c_str());
            MessageBoxW(hMainWindow, (L"Processing finished. Output: " + out_path).c_str(), L"Success", MB_OK | MB_ICONINFORMATION);
        }
        else {
            SetWindowTextW(hStatusText, L"NO Matches found...");
            MessageBoxW(hMainWindow, L"NO Matches found...", L"No Results", MB_OK | MB_ICONWARNING);
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
    // Get thread number
    wchar_t thread_num_str[16];
    GetWindowTextW(hThreadNumEntry, thread_num_str, 16);
    int thread_num = _wtoi(thread_num_str);
    if (thread_num < 1) thread_num = 1;
    THREAD_NUM = thread_num;
    EnableWindow(hProcessButton, FALSE);
    std::wstring output_format = GetOutputFormat();
    std::thread([=]() {
        ProcessMatching(daily_path, hist_path, output_format);
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
    wc.lpszClassName = L"WInPercentMainWindow";
    wc.hCursor = LoadCursorW(nullptr, (LPCWSTR)IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    RegisterClassW(&wc);

    hMainWindow = CreateWindowExW(0, wc.lpszClassName, L"WInPercent Matcher Processor", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_THICKFRAME | WS_MAXIMIZEBOX,
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
        CreateWindowW(L"STATIC", L"Historical % Input File:", WS_VISIBLE | WS_CHILD,
            10, 50, 162, 20, hwnd, nullptr, nullptr, nullptr);
        // Historical Folder Entry
        hHistEntry = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
            165, 50, 543, 20, hwnd, nullptr, nullptr, nullptr);
        // Browse Hist Button
        CreateWindowW(L"BUTTON", L"Browse", WS_VISIBLE | WS_CHILD,
            735, 50, 80, 20, hwnd, (HMENU)2, nullptr, nullptr);
        // Output Format Label
        CreateWindowW(L"STATIC", L"Output Format:", WS_VISIBLE | WS_CHILD,
            10, 80, 162, 20, hwnd, nullptr, nullptr, nullptr);
        // Output Format Radio Buttons
        hRadioCSV = CreateWindowW(L"BUTTON", L"CSV", WS_VISIBLE | WS_CHILD | BS_RADIOBUTTON,
            165, 80, 60, 20, hwnd, (HMENU)3, nullptr, nullptr);
        hRadioExcel = CreateWindowW(L"BUTTON", L"Excel", WS_VISIBLE | WS_CHILD | BS_RADIOBUTTON,
            235, 80, 80, 20, hwnd, (HMENU)4, nullptr, nullptr);
        SendMessageW(hRadioCSV, BM_SETCHECK, BST_CHECKED, 0); // Default to CSV
        // Thread Number Label
        CreateWindowW(L"STATIC", L"Threads:", WS_VISIBLE | WS_CHILD,
            10, 110, 162, 20, hwnd, nullptr, nullptr, nullptr);
        // Thread Number Entry
        hThreadNumEntry = CreateWindowW(L"EDIT", L"8", WS_VISIBLE | WS_CHILD | WS_BORDER,
            165, 110, 60, 20, hwnd, nullptr, nullptr, nullptr);
        // Process Button
        hProcessButton = CreateWindowW(L"BUTTON", L"Process", WS_VISIBLE | WS_CHILD,
            375, 140, 150, 30, hwnd, (HMENU)5, nullptr, nullptr);
        // Status Text
        hStatusText = CreateWindowW(L"STATIC", L"", WS_VISIBLE | WS_CHILD | SS_LEFT,
            10, 190, 870, 50, hwnd, nullptr, nullptr, nullptr);
        // Progress Bar
        hProgressBar = CreateWindowW(PROGRESS_CLASSW, L"", WS_VISIBLE | WS_CHILD,
            10, 250, 870, 20, hwnd, nullptr, nullptr, nullptr);
        // Progress Bar (optional, placeholder for now)
        // CreateWindowW(L"msctls_progress32", L"", WS_VISIBLE | WS_CHILD,
        //     10, 210, 870, 20, hwnd, nullptr, nullptr, nullptr);
        break;
    }
    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        switch (wmId) {
        case 1: OnBrowseDaily(); break;
        case 2: OnBrowseHist(); break;
        case 3: SendMessageW(hRadioCSV, BM_SETCHECK, BST_CHECKED, 0); SendMessageW(hRadioExcel, BM_SETCHECK, BST_UNCHECKED, 0); break;
        case 4: SendMessageW(hRadioCSV, BM_SETCHECK, BST_UNCHECKED, 0); SendMessageW(hRadioExcel, BM_SETCHECK, BST_CHECKED, 0); break;
        case 5: OnProcess(); break;
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
