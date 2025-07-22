#include <iostream>
#include <fstream>
#include <sstream>
#include <string>
#include <vector>
#include <map>
#include <filesystem>
#include <thread>
#include <mutex>
#include <algorithm>
#include <cmath>
#include <future>
#include <regex>
#include <unordered_map> // Added for faster lookup
#include <xlnt/xlnt.hpp> // Add this include for xlnt

#define THREAD_NUM 8

#ifdef _WIN32
#include <windows.h>
#include <commdlg.h>
#include <shlobj.h>
#include <commctrl.h>
#endif

// CSV/Excel-like data structure
using Row = std::vector<std::string>;
using DataFrame = std::vector<Row>;

// Column definitions
const std::vector<std::string> daily_cols = {
    "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ",
    "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK"
};

const std::vector<std::string> degree_cols = {
    "AQ", "AS", "AU", "AW", "AY", "BA", "BC", "BE", "BG", "BI", "BK"
};

// Global variables for GUI
HWND hMainWindow;
HWND hDailyEntry;
HWND hHistEntry;
HWND hProcessButton;
HWND hStatusText;
HWND hProgressBar;

// Forward declaration for DataProcessor
class DataProcessor;

// Window procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

class CSVReader {
public:
    static DataFrame readCSV(const std::string& filename) {
        if (endsWith(filename, ".csv")) {
            return readCSVFile(filename);
        }
        else if (endsWith(filename, ".xlsx")) {
            return readXLSXFile(filename);
        }
        else {
            throw std::runtime_error("Unsupported file type: " + filename);
        }
    }

    static void writeCSV(const DataFrame& data, const std::string& filename) {
        std::ofstream file(filename);
        if (!file.is_open()) {
            throw std::runtime_error("Cannot create file: " + filename);
        }

        for (const auto& row : data) {
            for (size_t i = 0; i < row.size(); ++i) {
                if (i > 0) file << ",";
                file << "\"" << row[i] << "\"";
            }
            file << "\n";
        }
    }

private:
    static bool endsWith(const std::string& str, const std::string& suffix) {
        return str.size() >= suffix.size() &&
            str.compare(str.size() - suffix.size(), suffix.size(), suffix) == 0;
    }

    static DataFrame readCSVFile(const std::string& filename) {
        DataFrame data;
        std::ifstream file(filename);
        if (!file.is_open()) {
            throw std::runtime_error("Cannot open file: " + filename);
        }

        std::string line;
        while (std::getline(file, line)) {
            Row row;
            std::stringstream ss(line);
            std::string cell;

            while (std::getline(ss, cell, ',')) {
                // Remove quotes and trim whitespace
                cell.erase(std::remove(cell.begin(), cell.end(), '"'), cell.end());
                cell.erase(0, cell.find_first_not_of(" \t"));
                cell.erase(cell.find_last_not_of(" \t") + 1);
                row.push_back(cell);
            }
            if (!row.empty()) {
                data.push_back(row);
            }
        }
        return data;
    }

    static DataFrame readXLSXFile(const std::string& filename) {
        DataFrame data;
        xlnt::workbook wb;
        wb.load(filename);
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
};

class DataProcessor {
private:
    std::mutex matches_mutex;

public:
    struct RowData {
        std::string player;
        std::map<std::string, std::string> data;
        std::string total;
        std::string winPercent;
    };

    RowData parseRowToDict(const Row& row) {
        RowData data;
        if (row.empty()) return data;

        data.player = row[0];

        // Parse pairs of key-value starting from index 1
        for (size_t i = 1; i < row.size() - 2; i += 2) {
            if (i + 1 < row.size()) {
                data.data[row[i]] = row[i + 1];
            }
        }

        if (row.size() >= 2) {
            data.total = row[row.size() - 2];
            data.winPercent = row[row.size() - 1];
        }

        return data;
    }

    bool degreeMatch(const std::string& daily_val, const std::string& hist_range) {
        try {
            std::regex range_regex(R"((\d+)-(\d+))");
            std::smatch matches;

            if (std::regex_match(hist_range, matches, range_regex)) {
                int low = std::stoi(matches[1]);
                int high = std::stoi(matches[2]);
                int daily_int = std::stoi(daily_val);
                return daily_int >= low && daily_int <= high;
            }
            return false;
        }
        catch (...) {
            return false;
        }
    }

    std::vector<Row> processChunk(const std::vector<std::pair<size_t, const Row*>>& chunk,
        const DataFrame& daily_df,
        const DataFrame& raw_daily_df) {
        std::vector<Row> matches;

        int match_count = 0; // Counter for matches in this chunk
        for (const auto& [idx, row] : chunk) {
            try {
                // Update status text
                if (idx % 10000 == 0) {
                    std::string status_msg = "Processing row " + std::to_string(idx) + "...";
                    SetWindowTextA(hStatusText, status_msg.c_str());
                }

                RowData hist_row = parseRowToDict(*row);

                for (size_t i = 0; i < daily_df.size(); ++i) {
                    const Row& daily_row = daily_df[i];
                    if (daily_row[0] != hist_row.player) continue;
                    bool is_match = true;

                    // Check each column for matching
                    for (const auto& [col, hist_val] : hist_row.data) {
                        if (col == "WinPercent" || col == "Total") continue;

                        // Find column index in daily data
                        auto col_it = std::find(daily_cols.begin(), daily_cols.end(), col);
                        if (col_it == daily_cols.end()) continue;

                        size_t col_idx = std::distance(daily_cols.begin(), col_it) + 1; // +1 for Player column
                        if (col_idx >= daily_row.size()) continue;

                        std::string daily_val = daily_row[col_idx];

                        if (daily_val.empty() || hist_val.empty()) continue;

                        // Check if this is a degree column
                        bool is_degree_col = std::find(degree_cols.begin(), degree_cols.end(), col) != degree_cols.end();

                        if (is_degree_col) {
                            if (!degreeMatch(daily_val, hist_val)) {
                                is_match = false;
                                break;
                            }
                        }
                        else {
                            if (daily_val != hist_val) {
                                is_match = false;
                                break;
                            }
                        }
                    }

                    if (is_match) {
                        match_count++;
                        if (match_count % 100 == 0) {
                            std::string match_msg = "***** Found matching result for row " + std::to_string(idx) + " (" + std::to_string(match_count) + " matches) *****";
                            SetWindowTextA(hStatusText, match_msg.c_str());
                        }
                        Row matched_row = raw_daily_df[i];
                        matched_row.insert(matched_row.end(), row->begin(), row->end());
                        matches.push_back(matched_row);
                    }
                }
            }
            catch (const std::exception& e) {
                std::string error_msg = "Error occurred in loop: " + std::string(e.what());
                SetWindowTextA(hStatusText, error_msg.c_str());
                continue;
            }
        }

        return matches;
    }

    DataFrame filterDailyData(const DataFrame& raw_daily_df) {
        DataFrame filtered_data;

        for (const Row& row : raw_daily_df) {
            Row filtered_row;

            // Extract column 0 (Player)
            if (!row.empty()) {
                filtered_row.push_back(row[0]);
            }

            // Extract columns 41-62 (indices 41-62)
            for (int i = 41; i <= 62 && i < static_cast<int>(row.size()); ++i) {
                filtered_row.push_back(row[i]);
            }

            if (!filtered_row.empty()) {
                filtered_data.push_back(filtered_row);
            }
        }

        return filtered_data;
    }

    void processFiles(const std::string& daily_file, const std::string& historical_folder) {
        try {
            SetWindowTextA(hStatusText, "Starting processing...");
            EnableWindow(hProcessButton, FALSE);

            // Check if files exist
            if (!std::filesystem::exists(daily_file) || !std::filesystem::exists(historical_folder)) {
                MessageBoxA(hMainWindow, "One or both files not found.", "Error", MB_OK | MB_ICONERROR);
                EnableWindow(hProcessButton, TRUE);
                return;
            }

            // Read daily file (now supports .csv and .xlsx)
            SetWindowTextA(hStatusText, "Reading daily file...");
            DataFrame raw_daily_df = CSVReader::readCSV(daily_file);
            DataFrame daily_df = filterDailyData(raw_daily_df);

            std::vector<Row> all_matches;

            // Count total files for progress
            int total_files = 0;
            for (const auto& entry : std::filesystem::directory_iterator(historical_folder)) {
                if (entry.is_regular_file()) {
                    std::string ext = entry.path().extension().string();
                    if (ext == ".csv" || ext == ".xlsx") {
                        total_files++;
                    }
                }
            }

            SendMessage(hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, total_files));
            SendMessage(hProgressBar, PBM_SETPOS, 0, 0);

            int processed_files = 0;

            // Process each file in historical folder
            for (const auto& entry : std::filesystem::directory_iterator(historical_folder)) {
                if (entry.is_regular_file()) {
                    std::string file_path = entry.path().string();
                    std::string file_name = entry.path().filename().string();

                    // Skip non-CSV/XLSX files
                    std::string ext = entry.path().extension().string();
                    if (ext != ".csv" && ext != ".xlsx") {
                        continue;
                    }

                    std::string status_msg = "Processing file: " + file_name;
                    SetWindowTextA(hStatusText, status_msg.c_str());

                    DataFrame raw_hist_df = CSVReader::readCSV(file_path);

                    // Create chunks for parallel processing
                    std::vector<std::pair<size_t, const Row*>> all_rows;
                    for (size_t i = 0; i < raw_hist_df.size(); ++i) {
                        all_rows.emplace_back(i, &raw_hist_df[i]);
                    }

                    size_t num_threads = (static_cast<size_t>(THREAD_NUM) < all_rows.size()) ? static_cast<size_t>(THREAD_NUM) : all_rows.size();
                    size_t chunk_size = static_cast<size_t>(std::ceil(static_cast<double>(all_rows.size()) / num_threads));
                    std::vector<std::future<std::vector<Row>>> futures;

                    // Process chunks in parallel
                    for (size_t i = 0; i < all_rows.size(); i += chunk_size) {
                        size_t end = (i + chunk_size < all_rows.size()) ? i + chunk_size : all_rows.size();
                        std::vector<std::pair<size_t, const Row*>> chunk(all_rows.begin() + i, all_rows.begin() + end);

                        futures.push_back(std::async(std::launch::async, [this, chunk, &daily_df, &raw_daily_df]() {
                            return processChunk(chunk, daily_df, raw_daily_df);
                            }));
                    }

                    // Collect results
                    for (auto& future : futures) {
                        auto chunk_matches = future.get();
                        all_matches.insert(all_matches.end(), chunk_matches.begin(), chunk_matches.end());
                    }

                    processed_files++;
                    SendMessage(hProgressBar, PBM_SETPOS, processed_files, 0);

                    std::string success_msg = "--------------- " + file_name + " Processed successfully ---------------";
                    SetWindowTextA(hStatusText, success_msg.c_str());
                }
            }

            // Save results
            if (!all_matches.empty()) {
                SetWindowTextA(hStatusText, "Saving matches...");
                // Remove empty rows
                all_matches.erase(
                    std::remove_if(all_matches.begin(), all_matches.end(),
                        [](const Row& row) { return row.empty(); }),
                    all_matches.end()
                );

                std::string output_path = daily_file.substr(0, daily_file.find_last_of('.')) + "_Matches.csv";
                CSVReader::writeCSV(all_matches, output_path);

                std::string success_msg = "Processing finished. Results saved to: " + output_path;
                SetWindowTextA(hStatusText, success_msg.c_str());
                MessageBoxA(hMainWindow, success_msg.c_str(), "Success", MB_OK | MB_ICONINFORMATION);
            }
            else {
                SetWindowTextA(hStatusText, "NO Matches found...");
                MessageBoxA(hMainWindow, "NO Matches found...", "No Results", MB_OK | MB_ICONWARNING);
            }

        }
        catch (const std::exception& e) {
            std::string error_msg = "Error occurred: " + std::string(e.what());
            SetWindowTextA(hStatusText, error_msg.c_str());
            MessageBoxA(hMainWindow, error_msg.c_str(), "Error", MB_OK | MB_ICONERROR);
        }

        EnableWindow(hProcessButton, TRUE);
        SendMessage(hProgressBar, PBM_SETPOS, 0, 0);
    }
};

DataProcessor* g_processor = nullptr;

// File dialog functions
std::string openFileDialog() {
    OPENFILENAMEA ofn;
    char szFile[260] = { 0 };

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = "Excel Files\0*.xlsx\0CSV Files\0*.csv\0All Files\0*.*\0";
    ofn.nFilterIndex = 1;
    ofn.lpstrFileTitle = NULL;
    ofn.nMaxFileTitle = 0;
    ofn.lpstrInitialDir = NULL;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileNameA(&ofn)) {
        return std::string(szFile);
    }
    return "";
}

std::string openFolderDialog() {
    BROWSEINFOA bi = { 0 };
    bi.lpszTitle = "Select Historical Folder";
    bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;

    LPITEMIDLIST pidl = SHBrowseForFolderA(&bi);
    if (pidl != nullptr) {
        char path[MAX_PATH];
        if (SHGetPathFromIDListA(pidl, path)) {
            CoTaskMemFree(pidl);
            return std::string(path);
        }
        CoTaskMemFree(pidl);
    }
    return "";
}

// Button click handlers
void OnBrowseDaily() {
    std::string file_path = openFileDialog();
    if (!file_path.empty()) {
        SetWindowTextA(hDailyEntry, file_path.c_str());
    }
}

void OnBrowseHist() {
    std::string folder_path = openFolderDialog();
    if (!folder_path.empty()) {
        SetWindowTextA(hHistEntry, folder_path.c_str());
    }
}

void OnProcess() {
    char daily_path[260];
    char hist_path[260];

    GetWindowTextA(hDailyEntry, daily_path, sizeof(daily_path));
    GetWindowTextA(hHistEntry, hist_path, sizeof(hist_path));

    if (strlen(daily_path) == 0 || strlen(hist_path) == 0) {
        MessageBoxA(hMainWindow, "Please select both daily file and historical folder.", "Error", MB_OK | MB_ICONERROR);
        return;
    }

    // Start processing in a separate thread
    std::thread([daily_path, hist_path]() {
        g_processor->processFiles(daily_path, hist_path);
        }).detach();
}

// Window procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE:
        return 0;

    case WM_SIZE: {
        int width = LOWORD(lParam);
        int height = HIWORD(lParam);

        // Resize controls based on window size
        if (hDailyEntry && hHistEntry && hProcessButton && hStatusText && hProgressBar) {
            // Calculate new positions and sizes
            int labelWidth = 162; // 90% of original 180px
            int entryWidth = width - labelWidth - 100; // Leave space for browse button
            int buttonX = width - 90;

            // Resize and reposition controls
            SetWindowPos(hDailyEntry, NULL, 165, 20, entryWidth, 20, SWP_NOZORDER);
            SetWindowPos(hHistEntry, NULL, 165, 50, entryWidth, 20, SWP_NOZORDER);

            // Move browse buttons
            SetWindowPos(GetDlgItem(hwnd, 1001), NULL, buttonX, 20, 80, 20, SWP_NOZORDER);
            SetWindowPos(GetDlgItem(hwnd, 1002), NULL, buttonX, 50, 80, 20, SWP_NOZORDER);

            // Center process button
            SetWindowPos(hProcessButton, NULL, (width - 150) / 2, 90, 150, 30, SWP_NOZORDER);

            // Resize status text and progress bar
            SetWindowPos(hStatusText, NULL, 10, 140, width - 20, 30, SWP_NOZORDER);
            SetWindowPos(hProgressBar, NULL, 10, 180, width - 20, 20, SWP_NOZORDER);
        }
        return 0;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case 1001: // Browse Daily button
            OnBrowseDaily();
            break;
        case 1002: // Browse Hist button
            OnBrowseHist();
            break;
        case 1003: // Process button
            OnProcess();
            break;
        }
        return 0;

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // Initialize COM
    CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

    // Initialize common controls
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_PROGRESS_CLASS;
    InitCommonControlsEx(&icex);

    // Register window class
    const char* CLASS_NAME = "MatcherWindow";
    WNDCLASSA wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);

    RegisterClassA(&wc);

    // Create main window
    hMainWindow = CreateWindowExA(
        0,
        CLASS_NAME,
        "Matcher Processor with non-coloring",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_THICKFRAME | WS_MAXIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 900, 400,
        NULL, NULL, hInstance, NULL
    );

    if (hMainWindow == NULL) {
        return 0;
    }

    // Create controls
    // Daily File Label
    CreateWindowA("STATIC", "Daily File:", WS_VISIBLE | WS_CHILD,
        10, 20, 162, 20, hMainWindow, NULL, hInstance, NULL);

    // Daily File Entry
    hDailyEntry = CreateWindowA("EDIT", "", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
        165, 20, 543, 20, hMainWindow, NULL, hInstance, NULL);

    // Browse Daily Button
    CreateWindowA("BUTTON", "Browse", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        735, 20, 80, 20, hMainWindow, (HMENU)1001, hInstance, NULL);

    // Historical Folder Label
    CreateWindowA("STATIC", "Historical % Input File:", WS_VISIBLE | WS_CHILD,
        10, 50, 162, 20, hMainWindow, NULL, hInstance, NULL);

    // Historical Folder Entry
    hHistEntry = CreateWindowA("EDIT", "", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
        165, 50, 543, 20, hMainWindow, NULL, hInstance, NULL);

    // Browse Hist Button
    CreateWindowA("BUTTON", "Browse", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        735, 50, 80, 20, hMainWindow, (HMENU)1002, hInstance, NULL);

    // Process Button
    hProcessButton = CreateWindowA("BUTTON", "Process", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        375, 90, 150, 30, hMainWindow, (HMENU)1003, hInstance, NULL);

    // Status Text
    hStatusText = CreateWindowA("STATIC", "Ready to process...", WS_VISIBLE | WS_CHILD | SS_LEFT,
        10, 140, 870, 50, hMainWindow, NULL, hInstance, NULL);

    // Progress Bar
    hProgressBar = CreateWindowA("msctls_progress32", NULL, WS_VISIBLE | WS_CHILD,
        10, 170, 870, 20, hMainWindow, NULL, hInstance, NULL);

    // Initialize processor
    g_processor = new DataProcessor();

    // Show window
    ShowWindow(hMainWindow, nCmdShow);
    UpdateWindow(hMainWindow);

    // Message loop
    MSG msg = {};
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // Cleanup
    delete g_processor;
    CoUninitialize();

    return 0;
}