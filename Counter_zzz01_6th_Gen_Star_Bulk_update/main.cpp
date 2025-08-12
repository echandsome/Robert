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
#include <set>
#include <iomanip>
#include <cmath>

using Row = std::vector<std::string>;
using DataFrame = std::vector<Row>;

// Available columns for selection (equivalent to Python COLUMN_MAPPING)
const std::map<std::string, int> COLUMN_MAPPING = {
    {"AQ", 42}, {"AS", 44}, {"AU", 46}, {"AW", 48}, {"AY", 50}, {"BA", 52},
    {"BC", 54}, {"BE", 56}, {"BG", 58}, {"BI", 60}, {"BK", 62}
};

// Structure to hold output row data
struct OutputRow {
    std::string player;
    std::map<std::string, std::string> col_data;
    std::map<std::string, std::string> val_data;
    int count;
    int match_total;
    int win_total;
    double win_percent_over;
};

// Structure to hold combination data
struct Combination {
    std::string col_name;
    int col_index;
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
            writeXLSXFile(filename);
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

    static void writeXLSXFile(const std::wstring& filename) {
        // Placeholder for XLSX writing - not implemented in this migration
        throw std::runtime_error("XLSX writing not implemented");
    }
};

// Generate combinations (equivalent to Python itertools.combinations)
std::vector<std::vector<Combination>> generateCombinations(int set_size) {
    std::vector<std::vector<Combination>> result;
    std::vector<Combination> items;
    
    for (const auto& pair : COLUMN_MAPPING) {
        items.push_back({pair.first, pair.second});
    }
    
    if (set_size > items.size()) return result;
    
    std::vector<bool> mask(items.size(), false);
    std::fill(mask.begin(), mask.begin() + set_size, true);
    
    do {
        std::vector<Combination> combination;
        for (size_t i = 0; i < items.size(); ++i) {
            if (mask[i]) {
                combination.push_back(items[i]);
            }
        }
        result.push_back(combination);
    } while (std::prev_permutation(mask.begin(), mask.end()));
    
    return result;
}

// Convert string to lowercase
std::string toLower(const std::string& str) {
    std::string result = str;
    std::transform(result.begin(), result.end(), result.begin(), ::tolower);
    return result;
}

// Convert string to integer safely
int safeStoi(const std::string& str) {
    try {
        return std::stoi(str);
    } catch (...) {
        return 0;
    }
}

// Process file function (equivalent to Python process_file)
std::vector<OutputRow> processFile(const DataFrame& input_df, const std::vector<Combination>& combination) {
    std::vector<OutputRow> output_rows;
    
    // Create column mapping for grouping
    std::map<std::string, std::vector<std::string>> grouped_data;
    
    // Process each row
    for (const auto& row : input_df) {
        if (row.size() < 8) continue; // Need at least 8 columns
        
        std::string player = row[0];
        std::string result = toLower(row[7]);
        
        // Create group key
        std::string group_key = player;
        for (const auto& combo : combination) {
            if (combo.col_index < row.size()) {
                group_key += "|" + row[combo.col_index];
            }
        }
        
        // Store data for grouping
        if (grouped_data.find(group_key) == grouped_data.end()) {
            grouped_data[group_key] = {player, result};
        } else {
            grouped_data[group_key].push_back(result);
        }
    }
    
    // Process grouped data
    for (const auto& group : grouped_data) {
        const std::string& group_key = group.first;
        const std::vector<std::string>& group_data = group.second;
        
        // Parse group key to get values
        std::vector<std::string> key_parts;
        std::stringstream ss(group_key);
        std::string part;
        while (std::getline(ss, part, '|')) {
            key_parts.push_back(part);
        }
        
        if (key_parts.size() < 2) continue;
        
        std::string player = key_parts[0];
        
        // Count results
        int over = 0, win = 0, under = 0, lose = 0;
        for (size_t i = 1; i < group_data.size(); ++i) {
            std::string result = group_data[i];
            if (result == "over") over++;
            else if (result == "win") win++;
            else if (result == "under") under++;
            else if (result == "lose") lose++;
        }
        
        over += win;
        under += lose;
        int total = over + under;
        
        if (total == 0) continue;
        
        // Create output row
        OutputRow output_row;
        output_row.player = player;
        output_row.match_total = total;
        output_row.win_total = over;
        output_row.win_percent_over = total > 0 ? std::round((double)over / total * 100.0) / 100.0 : 0.0;
        
        // Calculate count (sum of degree values)
        int row_sum = 0;
        for (size_t i = 1; i < key_parts.size(); ++i) {
            int col_id = i;  // Use 1-based indexing to match Python output
            std::string col_name = combination[i - 1].col_name;
            std::string val = key_parts[i];
            
            // Store column name and value (equivalent to Python logic)
            output_row.col_data["Col_" + std::to_string(col_id)] = col_name;
            output_row.val_data["Val_" + std::to_string(col_id)] = val;
            
            row_sum += safeStoi(val);
        }
        output_row.count = row_sum;
        
        output_rows.push_back(output_row);
    }
    
    return output_rows;
}

// Process file wrapper (equivalent to Python process_file_wrapper)
std::string processFileWrapper(const std::wstring& input_path, 
                              const std::vector<std::vector<Combination>>& combinations,
                              int set_size, 
                              const std::wstring& output_dir) {
    std::vector<OutputRow> all_results;
    
    try {
        std::wstring filename = std::filesystem::path(input_path).filename().wstring();
        std::wcout << L"→ " << filename << L" started" << std::endl;
        
        DataFrame input_df = CSVManager::read(input_path);
        
        for (size_t comb_id = 0; comb_id < combinations.size(); ++comb_id) {
            const auto& combination = combinations[comb_id];
                    std::wcout << L"  Combo " << (comb_id + 1) << L"/" << combinations.size() << L": ";
        for (const auto& combo : combination) {
            std::wcout << CSVManager::s2ws(combo.col_name) << L" ";
        }
        std::wcout << std::endl;
            
            auto results = processFile(input_df, combination);
            all_results.insert(all_results.end(), results.begin(), results.end());
        }
        
        if (!all_results.empty()) {
            // Create output filename
            std::wstring base_name = std::filesystem::path(input_path).stem().wstring();
            std::wstring output_name = base_name + L"_Size_" + std::to_wstring(set_size) + L"_Degree_YES.csv";
            std::wstring output_path = std::filesystem::path(output_dir) / output_name;
            
            // Convert OutputRow to DataFrame
            DataFrame output_df;
            
            // Add header row
            Row header = {"Player"};
                    for (size_t i = 1; i <= set_size; ++i) {
            header.push_back("Col_" + std::to_string(i));
            header.push_back("Val_" + std::to_string(i));
        }
            header.insert(header.end(), {"Count", "MATCH TOTAL", "WIN TOTAL", "WIN% OVER"});
            output_df.push_back(header);
            
            // Add data rows
            for (const auto& output_row : all_results) {
                Row data_row;
                data_row.push_back(output_row.player);
                
                for (size_t i = 1; i <= set_size; ++i) {
                    std::string col_key = "Col_" + std::to_string(i);
                    std::string val_key = "Val_" + std::to_string(i);
                    
                    auto col_it = output_row.col_data.find(col_key);
                    auto val_it = output_row.val_data.find(val_key);
                    
                    data_row.push_back(col_it != output_row.col_data.end() ? col_it->second : "");
                    data_row.push_back(val_it != output_row.val_data.end() ? val_it->second : "");
                }
                
                data_row.push_back(std::to_string(output_row.count));
                data_row.push_back(std::to_string(output_row.match_total));
                data_row.push_back(std::to_string(output_row.win_total));
                
                std::ostringstream oss;
                oss << std::fixed << std::setprecision(2) << output_row.win_percent_over;
                data_row.push_back(oss.str());
                
                output_df.push_back(data_row);
            }
            
            // Write output file
            CSVManager::write(output_df, output_path);
            std::wcout << L"✓ Saved to " << output_path << std::endl;
        }
        
        return CSVManager::ws2s(filename) + " completed";
    } catch (const std::exception& e) {
        return CSVManager::ws2s(input_path) + " failed: " + e.what();
    }
}

// Global handles for GUI controls
HWND hMainWindow;
HWND hInputEntry;
HWND hSetSizeVars[6]; // Radio buttons for set sizes 3-8
HWND hProcessButton;
HWND hStatusText;
HWND hProgressBar;
int selectedSetSize = 3; // Default set size is 3

// Function declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void OnBrowseInput();
void OnProcess();
std::wstring OpenFolderDialog();

// Main processing logic
void ProcessBulkFiles(const std::wstring& input_dir, int set_size) {
    try {
        SetWindowTextW(hStatusText, L"Generating combinations...");
        
        // Generate combinations
        auto combinations = generateCombinations(set_size);
        
        // Get list of files to process
        std::vector<std::wstring> files_to_process;
        for (const auto& entry : std::filesystem::directory_iterator(input_dir)) {
            if (!entry.is_regular_file()) continue;
            
            std::wstring filename = entry.path().filename().wstring();
            std::wstring ext = entry.path().extension().wstring();
            std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
            
            if ((ext == L".xlsx" && filename.find(L"~$") != 0) || ext == L".csv") {
                files_to_process.push_back(entry.path().wstring());
            }
        }
        
        if (files_to_process.empty()) {
            SetWindowTextW(hStatusText, L"No valid files found to process.");
            MessageBoxW(hMainWindow, L"No valid files found to process.", L"Warning", MB_OK | MB_ICONWARNING);
            return;
        }
        
        // Create output directory
        std::wstring output_dir = input_dir + L"_output";
        std::filesystem::create_directories(output_dir);
        
        // Setup progress bar
        SendMessageW(hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, files_to_process.size()));
        SendMessageW(hProgressBar, PBM_SETPOS, 0, 0);
        
        // Process files
        int processed = 0;
        for (const auto& file_path : files_to_process) {
            std::wstring filename = std::filesystem::path(file_path).filename().wstring();
            std::wstring status = L"Processing: " + filename;
            SetWindowTextW(hStatusText, status.c_str());
            
            std::string result = processFileWrapper(file_path, combinations, set_size, output_dir);
            std::wcout << CSVManager::s2ws(result) << std::endl;
            
            processed++;
            SendMessageW(hProgressBar, PBM_SETPOS, processed, 0);
        }
        
        SetWindowTextW(hStatusText, L"Processing finished successfully!");
        MessageBoxW(hMainWindow, (L"Processing finished. Output saved in: " + output_dir).c_str(), L"Success", MB_OK | MB_ICONINFORMATION);
        
    } catch (const std::exception& e) {
        std::wstring err = L"Error: ";
        err += CSVManager::s2ws(e.what());
        SetWindowTextW(hStatusText, err.c_str());
        MessageBoxW(hMainWindow, err.c_str(), L"Error", MB_OK | MB_ICONERROR);
    }
    
    SendMessageW(hProgressBar, PBM_SETPOS, 0, 0);
    EnableWindow(hProcessButton, TRUE);
}

void OnProcess() {
    wchar_t input_path[260];
    GetWindowTextW(hInputEntry, input_path, 260);
    
    if (wcslen(input_path) == 0) {
        MessageBoxW(hMainWindow, L"Please select an input folder.", L"Error", MB_OK | MB_ICONERROR);
        return;
    }
    
    EnableWindow(hProcessButton, FALSE);
    SetWindowTextW(hStatusText, L"Starting processing...");
    
            // Get selected set size
        for (size_t i = 0; i < 6; ++i) {
            if (SendMessageW(hSetSizeVars[i], BM_GETCHECK, 0, 0) == BST_CHECKED) {
                selectedSetSize = static_cast<int>(i + 3); // 3, 4, 5, 6, 7, 8
                break;
            }
        }
    
    // Start processing in separate thread
    std::thread([=]() {
        ProcessBulkFiles(input_path, selectedSetSize);
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
    wc.lpszClassName = L"BulkProcessorMainWindow";
    wc.hCursor = LoadCursorW(nullptr, (LPCWSTR)IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    RegisterClassW(&wc);

    hMainWindow = CreateWindowExW(0, wc.lpszClassName, L"Bulk File Processor", WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_THICKFRAME | WS_MAXIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 400, nullptr, nullptr, hInstance, nullptr);

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
        // Input Folder Label
        CreateWindowW(L"STATIC", L"Select Folder with Excel Files:", WS_VISIBLE | WS_CHILD,
            10, 20, 200, 20, hwnd, nullptr, nullptr, nullptr);
        
        // Input Folder Entry
        hInputEntry = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER,
            220, 20, 400, 20, hwnd, nullptr, nullptr, nullptr);
        
        // Browse Button
        CreateWindowW(L"BUTTON", L"Browse", WS_VISIBLE | WS_CHILD,
            640, 20, 80, 20, hwnd, (HMENU)1, nullptr, nullptr);
        
        // Set Size Label
        CreateWindowW(L"STATIC", L"Select Set Size:", WS_VISIBLE | WS_CHILD,
            10, 60, 200, 20, hwnd, nullptr, nullptr, nullptr);
        
        // Set Size Radio Buttons (3-8)
        const wchar_t* set_sizes[] = {L"3", L"4", L"5", L"6", L"7", L"8"};
        for (size_t i = 0; i < 6; ++i) {
            hSetSizeVars[i] = CreateWindowW(L"BUTTON", set_sizes[i], WS_VISIBLE | WS_CHILD | BS_RADIOBUTTON,
                220 + static_cast<int>(i * 60), 60, 50, 20, hwnd, (HMENU)(10 + static_cast<int>(i)), nullptr, nullptr);
        }
        SendMessageW(hSetSizeVars[0], BM_SETCHECK, BST_CHECKED, 0); // Default to 3
        
        // Process Button
        hProcessButton = CreateWindowW(L"BUTTON", L"Process", WS_VISIBLE | WS_CHILD,
            350, 100, 150, 30, hwnd, (HMENU)20, nullptr, nullptr);
        
        // Status Text
        hStatusText = CreateWindowW(L"STATIC", L"Ready to process files...", WS_VISIBLE | WS_CHILD | SS_LEFT,
            10, 150, 770, 50, hwnd, nullptr, nullptr, nullptr);
        
        // Progress Bar
        hProgressBar = CreateWindowW(PROGRESS_CLASSW, L"", WS_VISIBLE | WS_CHILD,
            10, 220, 770, 20, hwnd, nullptr, nullptr, nullptr);
        
        break;
    }
    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        if (wmId >= 10 && wmId <= 15) {
            // Set size radio button clicked
            for (size_t i = 0; i < 6; ++i) {
                SendMessageW(hSetSizeVars[i], BM_SETCHECK, (wParam == (10 + static_cast<int>(i))) ? BST_CHECKED : BST_UNCHECKED, 0);
            }
        }
        else if (wmId == 1) OnBrowseInput();
        else if (wmId == 20) OnProcess();
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

void OnBrowseInput() {
    std::wstring folder = OpenFolderDialog();
    if (!folder.empty()) {
        SetWindowTextW(hInputEntry, folder.c_str());
    }
}

std::wstring OpenFolderDialog() {
    BROWSEINFOW bi = { 0 };
    wchar_t szFolder[260] = { 0 };
    bi.hwndOwner = hMainWindow;
    bi.pszDisplayName = szFolder;
    bi.lpszTitle = L"Select Input Folder";
    LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
    if (pidl != nullptr) {
        SHGetPathFromIDListW(pidl, szFolder);
        return szFolder;
    }
    return L"";
}
