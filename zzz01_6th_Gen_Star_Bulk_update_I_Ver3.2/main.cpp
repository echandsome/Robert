#include <windows.h>
#include <commdlg.h>
#include <shlobj.h>
#include <thread>
#include <mutex>
#include <string>
#include <vector>
#include <map>
#include <filesystem>
#include <algorithm>
#include <iostream>
#include <fstream>
#include <sstream>
#include <iomanip>
#include <xlnt/xlnt.hpp>

namespace fs = std::filesystem;
using Row = std::vector<std::string>;
using DataFrame = std::vector<Row>;

// ==== Globals ====
HWND hInputEntry, hStatus, hProcessBtn, hDegreeCheck;
HWND hRadioBtns[6];
std::mutex cout_mutex;

// Correct column mapping matching Python version
std::map<std::string, int> COLUMN_MAPPING = {
    {"AP", 41}, {"AR", 43}, {"AT", 45}, {"AV", 47}, {"AX", 49},
    {"AZ", 51}, {"BB", 53}, {"BD", 55}, {"BF", 57}, {"BH", 59}, {"BJ", 61}
};

// ==== Excel/CSV Utilities ====
std::string get_next_column(const std::string &col) {
    char a = col[0], b = col[1];
    b++;
    if (b > 'Z') { 
        b = 'A'; 
        a++; 
    }
    return std::string(1, a) + std::string(1, b);
}

void write_csv(const DataFrame &data, const std::string &filename) {
    std::ofstream f(filename);
    if (!f.is_open()) {
        std::cerr << "Error: Cannot open file " << filename << std::endl;
        return;
    }
    
    for (const auto &row : data) {
        for (size_t i = 0; i < row.size(); i++) {
            if (i) f << ",";
            f << row[i];
        }
        f << "\n";
    }
}

DataFrame read_excel(const std::string &path) {
    DataFrame df;
    try {
        xlnt::workbook wb; 
        wb.load(path);
        auto ws = wb.active_sheet();
        for (auto r : ws.rows(false)) {
            Row row; 
            for (auto c : r) {
                row.push_back(c.to_string());
            }
            df.push_back(row);
        }
    } catch (const std::exception& e) {
        std::cerr << "Error reading Excel file: " << e.what() << std::endl;
    }
    return df;
}

// Helper function to create a key for grouping (matching Python's groupby logic)
std::string make_group_key(const Row &row, const std::vector<int> &col_indices) {
    std::string key;
    for (int idx : col_indices) {
        if (idx < (int)row.size()) {
            key += row[idx] + "|";
        } else {
            key += "|";
        }
    }
    return key;
}

// Process file function matching Python logic exactly
std::vector<std::map<std::string, std::string>> process_file(const DataFrame &df, bool is_degree,
                       const std::vector<std::pair<std::string,int>> &comb) {
    std::vector<std::map<std::string, std::string>> output_rows;
    
    // Build column selection like Python
    std::vector<std::string> selected_columns = {"Player"};
    std::vector<int> col_indexes = {0};
    
    for (auto &item : comb) {
        selected_columns.push_back(item.first);
        col_indexes.push_back(item.second);
        
        if (is_degree) {
            selected_columns.push_back(get_next_column(item.first));
            col_indexes.push_back(item.second + 1);
        }
    }
    
    // Group data by selected columns (matching Python's groupby)
    std::map<std::string, std::vector<Row>> groups;
    
    for (const auto &row : df) {
        if (row.size() <= 7) continue;
        
        // Create group key
        std::string group_key = make_group_key(row, col_indexes);
        
        // Store the row with its result
        Row group_row;
        for (int idx : col_indexes) {
            if (idx < (int)row.size()) {
                group_row.push_back(row[idx]);
            } else {
                group_row.push_back("");
            }
        }
        group_row.push_back(row[7]); // Result column
        groups[group_key].push_back(group_row);
    }
    
    // Process each group like Python
    for (auto &[group_key, group_rows] : groups) {
        int over = 0, under = 0;
        
        // Count results
        for (const auto &row : group_rows) {
            if (row.empty()) continue;
            std::string result = row.back();
            std::transform(result.begin(), result.end(), result.begin(), ::tolower);
            
            if (result == "over" || result == "win") over++;
            if (result == "under" || result == "lose") under++;
        }
        
        int total = over + under;
        if (total == 0) continue;
        
        // Create output row matching Python format exactly
        std::map<std::string, std::string> row_dict;
        
        // Get the first row's values for this group
        const auto &first_row = group_rows[0];
        
        int col_id = 0;
        for (size_t i = 0; i < selected_columns.size(); i++) {
            if (selected_columns[i] != "Player") {
                row_dict["Col_" + std::to_string(col_id)] = selected_columns[i];
            }
            row_dict["Col_" + std::to_string(col_id) + "_val"] = 
                (i < first_row.size()) ? first_row[i] : "";
            col_id++;
        }
        
        row_dict["Total"] = std::to_string(total);
        row_dict["WIN% OVER"] = std::to_string(round((double)over / total * 100.0) / 100.0);
        
        output_rows.push_back(row_dict);
    }
    
    return output_rows;
}

void combinations(const std::vector<std::pair<std::string,int>> &items, int k, int start,
                  std::vector<std::pair<std::string,int>> &cur,
                  std::vector<std::vector<std::pair<std::string,int>>> &res) {
    if ((int)cur.size() == k) {
        res.push_back(cur);
        return;
    }
    for (int i = start; i < (int)items.size(); i++) {
        cur.push_back(items[i]);
        combinations(items, k, i + 1, cur, res);
        cur.pop_back();
    }
}

// Convert map to CSV row (matching Python's DataFrame.to_csv behavior)
Row map_to_csv_row(const std::map<std::string, std::string> &row_dict) {
    Row csv_row;
    // Order: Col_0, Col_0_val, Col_1, Col_1_val, ..., Total, WIN% OVER
    std::vector<std::string> keys;
    for (const auto &[key, value] : row_dict) {
        keys.push_back(key);
    }
    std::sort(keys.begin(), keys.end());
    
    for (const auto &key : keys) {
        csv_row.push_back(row_dict.at(key));
    }
    return csv_row;
}

void process_excel_file(const fs::path &file, bool deg, int k, const fs::path &out) {
    try {
        std::cout << "→ " << file.filename().string() << " started" << std::endl;
        
        DataFrame df = read_excel(file.string());
        if (df.empty()) {
            std::cout << "Error: Empty or invalid Excel file" << std::endl;
            return;
        }
        
        std::vector<std::pair<std::string,int>> items(COLUMN_MAPPING.begin(), COLUMN_MAPPING.end());
        std::vector<std::vector<std::pair<std::string,int>>> combos; 
        std::vector<std::pair<std::string,int>> cur;
        combinations(items, k, 0, cur, combos);
        
        std::vector<std::map<std::string, std::string>> all_results;
        
        for (size_t i = 0; i < combos.size(); i++) {
            std::cout << "  Combo " << (i+1) << "/" << combos.size() << ": ";
            for (const auto& c : combos[i]) {
                std::cout << c.first << " ";
            }
            std::cout << std::endl;
            
            auto results = process_file(df, deg, combos[i]);
            all_results.insert(all_results.end(), results.begin(), results.end());
        }
        
        if (!all_results.empty()) {
            // Convert to DataFrame format for CSV writing
            DataFrame csv_data;
            for (const auto &row_dict : all_results) {
                csv_data.push_back(map_to_csv_row(row_dict));
            }
            
            std::string output_name = file.stem().string() + "_Size_" + std::to_string(k) + 
                                    "_Degree_" + (deg ? "YES" : "NO") + ".csv";
            fs::path output_path = out / output_name;
            write_csv(csv_data, output_path.string());
            std::cout << "✓ Saved to " << output_path.string() << std::endl;
        }
        
        std::cout << file.filename().string() << " completed" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << file.string() << " failed: " << e.what() << std::endl;
    }
}

// ==== GUI File Picker ====
std::wstring BrowseFolder() {
    BROWSEINFOW bi = {0}; 
    wchar_t buf[260];
    bi.pszDisplayName = buf; 
    bi.lpszTitle = L"Select Input Folder";
    LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
    if (pidl) {
        SHGetPathFromIDListW(pidl, buf);
        return buf;
    }
    return L"";
}

// ==== Threaded Bulk Processing ====
void RunProcessing(std::wstring folder, bool deg, int set_size) {
    try {
        fs::path in = folder; 
        fs::path out = in.string() + "_output"; 
        fs::create_directories(out);
        
        for (auto &e : fs::directory_iterator(in)) {
            if (e.path().extension() == ".xlsx" && 
                e.path().filename().string().substr(0, 2) != "~$") {
                process_excel_file(e.path(), deg, set_size, out);
            }
        }
        
        SetWindowTextW(hStatus, L"Processing Complete!");
        EnableWindow(hProcessBtn, TRUE);
        
    } catch (const std::exception& e) {
        std::wstring error_msg = L"Error: " + std::wstring(e.what(), e.what() + strlen(e.what()));
        SetWindowTextW(hStatus, error_msg.c_str());
        EnableWindow(hProcessBtn, TRUE);
    }
}

// ==== GUI Events ====
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
        case WM_CREATE: {
            CreateWindowW(L"STATIC", L"Input Folder:", WS_CHILD | WS_VISIBLE, 10, 20, 100, 20, hwnd, 0, 0, 0);
            hInputEntry = CreateWindowW(L"EDIT", L"", WS_CHILD | WS_VISIBLE | WS_BORDER, 120, 20, 300, 20, hwnd, 0, 0, 0);
            CreateWindowW(L"BUTTON", L"Browse", WS_CHILD | WS_VISIBLE, 430, 20, 80, 20, hwnd, (HMENU)1, 0, 0);
            hDegreeCheck = CreateWindowW(L"BUTTON", L"Include Degrees", WS_CHILD | WS_VISIBLE | BS_AUTOCHECKBOX, 10, 60, 150, 20, hwnd, 0, 0, 0);
            CreateWindowW(L"STATIC", L"Set Size:", WS_CHILD | WS_VISIBLE, 10, 100, 70, 20, hwnd, 0, 0, 0);
            int sizes[6] = {3, 4, 5, 6, 7, 8};
            for (int i = 0; i < 6; i++) {
                hRadioBtns[i] = CreateWindowW(L"BUTTON", (LPCWSTR)(std::to_wstring(sizes[i]).c_str()),
                    WS_CHILD | WS_VISIBLE | BS_RADIOBUTTON, 80 + i * 50, 100, 40, 20, hwnd, (HMENU)(100 + i), 0, 0);
            }
            SendMessageW(hRadioBtns[0], BM_SETCHECK, BST_CHECKED, 0);
            hProcessBtn = CreateWindowW(L"BUTTON", L"Process", WS_CHILD | WS_VISIBLE, 10, 140, 100, 30, hwnd, (HMENU)2, 0, 0);
            hStatus = CreateWindowW(L"STATIC", L"", WS_CHILD | WS_VISIBLE, 10, 180, 400, 40, hwnd, 0, 0, 0);
            break;
        }
        case WM_COMMAND: {
            if (LOWORD(wp) == 1) {
                auto f = BrowseFolder();
                if (!f.empty()) SetWindowTextW(hInputEntry, f.c_str());
            }
            if (LOWORD(wp) == 2) {
                wchar_t buf[260]; 
                GetWindowTextW(hInputEntry, buf, 260); 
                if (wcslen(buf) == 0) {
                    MessageBoxW(hwnd, L"Select folder", L"Error", 0);
                    break;
                }
                bool deg = (SendMessageW(hDegreeCheck, BM_GETCHECK, 0, 0) == BST_CHECKED);
                int set_size = 3; 
                for (int i = 0; i < 6; i++) {
                    if (SendMessageW(hRadioBtns[i], BM_GETCHECK, 0, 0) == BST_CHECKED) {
                        set_size = 3 + i;
                        break;
                    }
                }
                EnableWindow(hProcessBtn, FALSE); 
                SetWindowTextW(hStatus, L"Processing...");
                std::thread([=] { RunProcessing(buf, deg, set_size); }).detach();
            } 
            break;
        }
        case WM_DESTROY: 
            PostQuitMessage(0); 
            break;
        default: 
            return DefWindowProcW(hwnd, msg, wp, lp);
    } 
    return 0;
}

// ==== Entry Point ====
int WINAPI WinMain(HINSTANCE hInst, HINSTANCE, LPSTR, int nCmdShow) {
    WNDCLASSW wc = {}; 
    wc.lpszClassName = L"BulkProc"; 
    wc.lpfnWndProc = WndProc; 
    wc.hInstance = hInst;
    RegisterClassW(&wc);
    HWND hwnd = CreateWindowW(L"BulkProc", L"Bulk Excel Processor", WS_OVERLAPPEDWINDOW | WS_VISIBLE, 100, 100, 550, 300, 0, 0, hInst, 0);
    MSG msg; 
    while (GetMessageW(&msg, 0, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    return 0;
}
