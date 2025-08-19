#include <windows.h>
#include <commdlg.h>
#include <shlobj.h>
#include <commctrl.h>
#include <string>
#include <vector>
#include <iostream>
#include <fstream>
#include <sstream>
#include <algorithm>
#include <filesystem>
#include <thread>
#include <memory>
#include <xlnt/xlnt.hpp>
#include <cmath>
#include <chrono>
#define M_PI 3.14159265358979323846

// Forward declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
void OnBrowseFile();
void OnStartProcessing();
std::wstring OpenFileDialog();
std::wstring OpenFolderDialog();

// Global variables
HWND hMainWindow;
HWND hFileEntry;
HWND hBrowseButton;
HWND hDOBColumnDropdown;
HWND hDateColumnDropdown;
HWND hStartButton;
HWND hStatusText;
HWND hProgressBar;

// Biorhythm calculation structure
struct BiorhythmResult {
    double physical;
    double emotional;
    double intellectual;
    double spiritual;
    double awareness;
    double intuitive;
    double aesthetic;
};

// Data structures
struct BiorhythmData {
    std::vector<std::vector<std::string>> data;
    std::vector<std::string> headers;
    std::string filePath;
    std::string folderPath;
    std::string dobColumn;
    std::string dateColumn;
    std::vector<std::string> biorhythmColumns = {"Emotional", "Physical", "Intellectual", "Spiritual", "Awareness", "Intuitive", "Aesthetic"};
};

BiorhythmData biorhythmData;

// Biorhythm calculation function
BiorhythmResult calculateBiorhythm(const std::string& birthDateStr, const std::string& targetDateStr) {
    BiorhythmResult result = {0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0};
    
    try {
        // Parse dates with multiple format support
        std::tm birthDate = {}, targetDate = {};
        bool parseSuccess = false;
        
        // Try multiple date formats
        std::vector<std::pair<std::string, std::string>> dateFormats = {
            {"%Y-%m-%d", "YYYY-MM-DD"},           // 1991-12-14
            {"%m/%d/%Y", "M/D/YYYY"},             // 12/14/1991
            {"%d/%m/%Y", "DD/MM/YYYY"},           // 14/12/1991
            {"%Y/%m/%d", "YYYY/MM/DD"},           // 1991/12/14
            {"%m-%d-%Y", "MM-DD-YYYY"},           // 12-14-1991
            {"%d-%m-%Y", "DD-MM-YYYY"},           // 14-12-1991
            {"%Y.%m.%d", "YYYY.MM.DD"},           // 1991.12.14
            {"%m.%d.%Y", "MM.DD.YYYY"},           // 12.14.1991
            {"%d.%m.%Y", "DD.MM.YYYY"},           // 14.12.1991
            {"%Y %m %d", "YYYY MM DD"},           // 1991 12 14
            {"%m %d %Y", "MM DD YYYY"},           // 12 14 1991
            {"%d %m %Y", "DD MM YYYY"}            // 14 12 1991
        };
        
        for (const auto& format : dateFormats) {
            std::istringstream birthStream(birthDateStr);
            std::istringstream targetStream(targetDateStr);
            
            // Reset tm structures
            birthDate = {};
            targetDate = {};
            
            birthStream >> std::get_time(&birthDate, format.first.c_str());
            targetStream >> std::get_time(&targetDate, format.first.c_str());
            
            if (!birthStream.fail() && !targetStream.fail()) {
                parseSuccess = true;
                break;
            }
        }
        
        if (!parseSuccess) {
            return result; // Return zeros if all parsing attempts fail
        }
        
        // Convert to time_t
        std::time_t birthTime = std::mktime(&birthDate);
        std::time_t targetTime = std::mktime(&targetDate);
        
        if (birthTime == -1 || targetTime == -1) {
            return result;
        }
        
        // Calculate days since birth
        double daysSinceBirth = std::difftime(targetTime, birthTime) / (24 * 60 * 60);
        
        // Biorhythm cycles (constants)
        const double EMOTIONAL_CYCLE = 28.0;
        const double PHYSICAL_CYCLE = 23.0;
        const double INTELLECTUAL_CYCLE = 33.0;
        const double SPIRITUAL_CYCLE = 53.0;
        const double AWARENESS_CYCLE = 48.0;
        const double INTUITIVE_CYCLE = 38.0;
        const double AESTHETIC_CYCLE = 43.0;
        
        // Calculate each biorhythm value with rounding
        result.emotional = std::round(std::sin(2.0 * M_PI * daysSinceBirth / EMOTIONAL_CYCLE) * 100.0);
        result.physical = std::round(std::sin(2.0 * M_PI * daysSinceBirth / PHYSICAL_CYCLE) * 100.0);
        result.intellectual = std::round(std::sin(2.0 * M_PI * daysSinceBirth / INTELLECTUAL_CYCLE) * 100.0);
        result.spiritual = std::round(std::sin(2.0 * M_PI * daysSinceBirth / SPIRITUAL_CYCLE) * 100.0);
        result.awareness = std::round(std::sin(2.0 * M_PI * daysSinceBirth / AWARENESS_CYCLE) * 100.0);
        result.intuitive = std::round(std::sin(2.0 * M_PI * daysSinceBirth / INTUITIVE_CYCLE) * 100.0);
        result.aesthetic = std::round(std::sin(2.0 * M_PI * daysSinceBirth / AESTHETIC_CYCLE) * 100.0);
        
    } catch (...) {
        return result; // Return zeros if any error occurs
    }
    
    return result;
}

// Helper function to convert Excel column letter to 0-based index
int columnLetterToIndex(const std::string& columnLetter) {
    int result = 0;
    for (char c : columnLetter) {
        if (c >= 'A' && c <= 'Z') {
            result = result * 26 + (c - 'A' + 1);
        } else if (c >= 'a' && c <= 'z') {
            result = result * 26 + (c - 'a' + 1);
        }
    }
    return result - 1; // Convert to 0-based index
}

// CSV/Excel file manager
class FileManager {
public:
    static bool readFile(const std::string& filePath, BiorhythmData& data) {
        std::filesystem::path path(filePath);
        std::string extension = path.extension().string();
        std::transform(extension.begin(), extension.end(), extension.begin(), ::tolower);
        
        if (extension == ".csv") {
            return readCSVFile(filePath, data);
        } else if (extension == ".xlsx" || extension == ".xls") {
            return readExcelFile(filePath, data);
        }
        return false;
    }
    
    static void writeFile(const BiorhythmData& data, const std::string& filePath) {
        std::filesystem::path path(filePath);
        std::string extension = path.extension().string();
        std::transform(extension.begin(), extension.end(), extension.begin(), ::tolower);
        
        if (extension == ".csv") {
            writeCSVFile(data, filePath);
        } else if (extension == ".xlsx" || extension == ".xls") {
            writeExcelFile(data, filePath);
        }
    }

private:
    static bool readCSVFile(const std::string& filePath, BiorhythmData& data) {
        std::ifstream file(filePath);
        if (!file.is_open()) return false;
        
        data.data.clear();
        data.headers.clear();
        
        std::string line;
        
        while (std::getline(file, line)) {
            std::vector<std::string> row;
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
                data.data.push_back(row);
            }
        }
        
        // Generate default headers based on column count
        if (!data.data.empty()) {
            size_t maxCols = 0;
            for (const auto& row : data.data) {
                maxCols = std::max(maxCols, row.size());
            }
            
            for (size_t i = 0; i < maxCols; ++i) {
                std::string colName = "C_" + std::to_string(i + 1);
                data.headers.push_back(colName);
            }
        }
        
        return true;
    }
    
    static bool readExcelFile(const std::string& filePath, BiorhythmData& data) {
        try {
            xlnt::workbook wb;
            wb.load(filePath);
            auto ws = wb.active_sheet();
            
            data.data.clear();
            data.headers.clear();
            
            for (auto row : ws.rows(false)) {
                std::vector<std::string> rowData;
                for (auto cell : row) {
                    rowData.push_back(cell.to_string());
                }
                
                if (!rowData.empty()) {
                    data.data.push_back(rowData);
                }
            }
            
            // Generate default headers based on column count
            if (!data.data.empty()) {
                size_t maxCols = 0;
                for (const auto& row : data.data) {
                    maxCols = std::max(maxCols, row.size());
                }
                
                for (size_t i = 0; i < maxCols; ++i) {
                    std::string colName = "C_" + std::to_string(i + 1);
                    data.headers.push_back(colName);
                }
            }
            
            return true;
        } catch (...) {
            return false;
        }
    }
    
    static void writeCSVFile(const BiorhythmData& data, const std::string& filePath) {
        std::ofstream file(filePath);
        if (!file.is_open()) return;
        
        // Write headers
        for (size_t i = 0; i < data.headers.size(); ++i) {
            if (i > 0) file << ",";
            file << "\"" << data.headers[i] << "\"";
        }
        file << "\n";
        
        // Write data
        for (const auto& row : data.data) {
            for (size_t i = 0; i < row.size(); ++i) {
                if (i > 0) file << ",";
                file << "\"" << row[i] << "\"";
            }
            file << "\n";
        }
    }
    
    static void writeExcelFile(const BiorhythmData& data, const std::string& filePath) {
        try {
            xlnt::workbook wb;
            auto ws = wb.active_sheet();
            
            // Write headers
            for (size_t i = 0; i < data.headers.size(); ++i) {
                ws.cell(static_cast<uint32_t>(i + 1), 1).value(data.headers[i]);
            }
            
            // Write data
            for (size_t rowIdx = 0; rowIdx < data.data.size(); ++rowIdx) {
                const auto& row = data.data[rowIdx];
                for (size_t colIdx = 0; colIdx < row.size(); ++colIdx) {
                    ws.cell(static_cast<uint32_t>(colIdx + 1), static_cast<uint32_t>(rowIdx + 2)).value(row[colIdx]);
                }
            }
            
            wb.save(filePath);
        } catch (...) {
            // Handle error
        }
    }
};

// UI Helper functions
void updateColumnDropdowns() {
    // Clear existing items
    SendMessage(hDOBColumnDropdown, CB_RESETCONTENT, 0, 0);
    SendMessage(hDateColumnDropdown, CB_RESETCONTENT, 0, 0);
    
    if (biorhythmData.headers.empty()) return;
    
    // Add column options
    for (size_t i = 0; i < biorhythmData.headers.size(); ++i) {
        std::wstring colName = std::wstring(biorhythmData.headers[i].begin(), biorhythmData.headers[i].end());
        
        // Convert index to Excel column name (A, B, C, ..., Z, AA, AB, AC, ...)
        std::wstring colLetter;
        size_t temp = i;
        while (temp >= 0) {
            colLetter = std::wstring(1, L'A' + (temp % 26)) + colLetter;
            if (temp < 26) break;
            temp = temp / 26 - 1;
        }
        
        std::wstring displayText = colLetter + L" (" + colName + L")";
        
        SendMessage(hDOBColumnDropdown, CB_ADDSTRING, 0, (LPARAM)displayText.c_str());
        SendMessage(hDateColumnDropdown, CB_ADDSTRING, 0, (LPARAM)displayText.c_str());
    }
    
    // Set default selections: B (column 1) and P (column 15)
    if (biorhythmData.headers.size() > 1) {
        SendMessage(hDOBColumnDropdown, CB_SETCURSEL, 1, 0); // Column B (index 1)
    } else if (biorhythmData.headers.size() > 0) {
        SendMessage(hDOBColumnDropdown, CB_SETCURSEL, 0, 0); // Column A (index 0)
    }
    
    if (biorhythmData.headers.size() > 15) {
        SendMessage(hDateColumnDropdown, CB_SETCURSEL, 15, 0); // Column P (index 15)
    } else if (biorhythmData.headers.size() > 1) {
        SendMessage(hDateColumnDropdown, CB_SETCURSEL, 1, 0); // Column B (index 1)
    } else if (biorhythmData.headers.size() > 0) {
        SendMessage(hDateColumnDropdown, CB_SETCURSEL, 0, 0); // Column A (index 0)
    }
}

void OnBrowseFile() {
    std::wstring filePath = OpenFileDialog();
    if (!filePath.empty()) {
        // Convert wide string to regular string
        std::string narrowPath(filePath.begin(), filePath.end());
        
        // Set the file path in the entry field
        SetWindowTextA(hFileEntry, narrowPath.c_str());
        
        // Load the file
        if (FileManager::readFile(narrowPath, biorhythmData)) {
            biorhythmData.filePath = narrowPath;
            
            // Create output folder
            std::filesystem::path path(narrowPath);
            std::string fileName = path.stem().string();
            biorhythmData.folderPath = (path.parent_path() / fileName).string();
            
            if (!std::filesystem::exists(biorhythmData.folderPath)) {
                std::filesystem::create_directories(biorhythmData.folderPath);
            }
            
            // Update column dropdowns
            updateColumnDropdowns();
            
            // Update status
            std::string status = "File loaded: " + narrowPath + " (" + std::to_string(biorhythmData.data.size()) + " rows)";
            SetWindowTextA(hStatusText, status.c_str());
            
            // Enable start button
            EnableWindow(hStartButton, TRUE);
        } else {
            SetWindowTextA(hStatusText, "Error: Could not load file");
            EnableWindow(hStartButton, FALSE);
        }
    }
}

void OnStartProcessing() {
    if (biorhythmData.data.empty()) {
        SetWindowTextA(hStatusText, "Error: No file loaded");
        return;
    }
    
    // Get selected columns
    int dobIdx = SendMessage(hDOBColumnDropdown, CB_GETCURSEL, 0, 0);
    int dateIdx = SendMessage(hDateColumnDropdown, CB_GETCURSEL, 0, 0);
    
    if (dobIdx == CB_ERR || dateIdx == CB_ERR) {
        SetWindowTextA(hStatusText, "Error: Please select all columns");
        return;
    }
    
    biorhythmData.dobColumn = biorhythmData.headers[dobIdx];
    biorhythmData.dateColumn = biorhythmData.headers[dateIdx];
    
    // Disable start button during processing
    EnableWindow(hStartButton, FALSE);
    
    // Update status
    std::string status = "Processing started... DOB: " + biorhythmData.dobColumn + 
                        ", Date: " + biorhythmData.dateColumn;
    SetWindowTextA(hStatusText, status.c_str());
    
    // Set progress bar
    SendMessage(hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, biorhythmData.data.size()));
    SendMessage(hProgressBar, PBM_SETPOS, 0, 0);
    
    // Process biorhythms in a separate thread
    std::thread([=]() {
        try {
            // Add biorhythm column headers if they don't exist
            bool needHeaders = true;
            for (const auto& header : biorhythmData.headers) {
                if (header == "Emotional" || header == "Physical" || header == "Intellectual") {
                    needHeaders = false;
                    break;
                }
            }
            
            if (needHeaders) {
                biorhythmData.headers.insert(biorhythmData.headers.end(), 
                    biorhythmData.biorhythmColumns.begin(), biorhythmData.biorhythmColumns.end());
            }
            
            // Process each row
            for (size_t i = 0; i < biorhythmData.data.size(); ++i) {
                // Update progress
                SendMessage(hProgressBar, PBM_SETPOS, i + 1, 0);
                
                // Update status
                std::string progressStatus = "Processing row " + std::to_string(i + 1) + "/" + std::to_string(biorhythmData.data.size());
                SetWindowTextA(hStatusText, progressStatus.c_str());
                
                // Get DOB and target date from the row
                if (dobIdx < biorhythmData.data[i].size() && dateIdx < biorhythmData.data[i].size()) {
                    std::string birthDate = biorhythmData.data[i][dobIdx];
                    std::string targetDate = biorhythmData.data[i][dateIdx];

                    // Calculate biorhythm
                    BiorhythmResult biorhythm = calculateBiorhythm(birthDate, targetDate);
                    
                    // Add biorhythm values to the row
                    if (needHeaders) {
                        // Add new columns for biorhythm values
                        biorhythmData.data[i].push_back(std::to_string(static_cast<int>(biorhythm.emotional)));
                        biorhythmData.data[i].push_back(std::to_string(static_cast<int>(biorhythm.physical)));
                        biorhythmData.data[i].push_back(std::to_string(static_cast<int>(biorhythm.intellectual)));
                        biorhythmData.data[i].push_back(std::to_string(static_cast<int>(biorhythm.spiritual)));
                        biorhythmData.data[i].push_back(std::to_string(static_cast<int>(biorhythm.awareness)));
                        biorhythmData.data[i].push_back(std::to_string(static_cast<int>(biorhythm.intuitive)));
                        biorhythmData.data[i].push_back(std::to_string(static_cast<int>(biorhythm.aesthetic)));
                    } else {
                        // Update existing biorhythm columns
                        size_t baseIdx = biorhythmData.headers.size() - 7;
                        if (biorhythmData.data[i].size() <= baseIdx) {
                            biorhythmData.data[i].resize(baseIdx + 7);
                        }
                        biorhythmData.data[i][baseIdx] = std::to_string(static_cast<int>(biorhythm.emotional));
                        biorhythmData.data[i][baseIdx + 1] = std::to_string(static_cast<int>(biorhythm.physical));
                        biorhythmData.data[i][baseIdx + 2] = std::to_string(static_cast<int>(biorhythm.intellectual));
                        biorhythmData.data[i][baseIdx + 3] = std::to_string(static_cast<int>(biorhythm.spiritual));
                        biorhythmData.data[i][baseIdx + 4] = std::to_string(static_cast<int>(biorhythm.awareness));
                        biorhythmData.data[i][baseIdx + 5] = std::to_string(static_cast<int>(biorhythm.intuitive));
                        biorhythmData.data[i][baseIdx + 6] = std::to_string(static_cast<int>(biorhythm.aesthetic));
                    }
                }
                
                // Small delay to prevent UI freezing
                std::this_thread::sleep_for(std::chrono::milliseconds(10));
            }
            
            // Save the updated data
            std::filesystem::path inputPath(biorhythmData.filePath);
            std::string extension = inputPath.extension().string();
            std::transform(extension.begin(), extension.end(), extension.begin(), ::tolower);
            
            std::string outputPath;
            if (extension == ".csv") {
                outputPath = biorhythmData.filePath.substr(0, biorhythmData.filePath.find_last_of('.')) + "_with_biorhythms.csv";
            } else {
                outputPath = biorhythmData.filePath.substr(0, biorhythmData.filePath.find_last_of('.')) + "_with_biorhythms.xlsx";
            }
            
            FileManager::writeFile(biorhythmData, outputPath);
            
            // Processing complete
            std::string completeStatus = "Processing completed successfully! Output saved to: " + outputPath;
            SetWindowTextA(hStatusText, completeStatus.c_str());
            
        } catch (const std::exception& e) {
            std::string errorStatus = "Error during processing: " + std::string(e.what());
            SetWindowTextA(hStatusText, errorStatus.c_str());
        }
        
        // Re-enable start button
        EnableWindow(hStartButton, TRUE);
        SendMessage(hProgressBar, PBM_SETPOS, 0, 0);
    }).detach();
}

// File dialog functions
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
    bi.lpszTitle = L"Select Output Folder";
    
    LPITEMIDLIST pidl = SHBrowseForFolderW(&bi);
    if (pidl != nullptr) {
        SHGetPathFromIDListW(pidl, szFolder);
        return szFolder;
    }
    return L"";
}

// Main window procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE: {
        // File selection section
        CreateWindowW(L"STATIC", L"Select File:", WS_VISIBLE | WS_CHILD,
            10, 20, 100, 20, hwnd, nullptr, nullptr, nullptr);
        
        hFileEntry = CreateWindowW(L"EDIT", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_READONLY,
            120, 20, 400, 20, hwnd, nullptr, nullptr, nullptr);
        
        hBrowseButton = CreateWindowW(L"BUTTON", L"Browse", WS_VISIBLE | WS_CHILD,
            530, 20, 80, 20, hwnd, (HMENU)1, nullptr, nullptr);
        
        // Column selection section
        CreateWindowW(L"STATIC", L"Date of Birth Column:", WS_VISIBLE | WS_CHILD,
            10, 60, 120, 20, hwnd, nullptr, nullptr, nullptr);
        
        hDOBColumnDropdown = CreateWindowW(L"COMBOBOX", L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST | WS_VSCROLL,
            120, 60, 200, 200, hwnd, nullptr, nullptr, nullptr);
        
        CreateWindowW(L"STATIC", L"Date Column:", WS_VISIBLE | WS_CHILD,
            10, 90, 100, 20, hwnd, nullptr, nullptr, nullptr);
        
        hDateColumnDropdown = CreateWindowW(L"COMBOBOX", L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST | WS_VSCROLL,
            120, 90, 200, 200, hwnd, nullptr, nullptr, nullptr);
        
        // Start button
        hStartButton = CreateWindowW(L"BUTTON", L"Start Processing", WS_VISIBLE | WS_CHILD,
            200, 130, 150, 30, hwnd, (HMENU)2, nullptr, nullptr);
        EnableWindow(hStartButton, FALSE); // Disabled until file is loaded
        
        // Status and progress
        hStatusText = CreateWindowW(L"STATIC", L"Waiting for file selection...", WS_VISIBLE | WS_CHILD | SS_LEFT,
            10, 180, 600, 20, hwnd, nullptr, nullptr, nullptr);
        
        hProgressBar = CreateWindowW(PROGRESS_CLASSW, L"", WS_VISIBLE | WS_CHILD,
            10, 210, 600, 20, hwnd, nullptr, nullptr, nullptr);
        
        break;
    }
    
    case WM_COMMAND: {
        int wmId = LOWORD(wParam);
        switch (wmId) {
        case 1: // Browse button
            OnBrowseFile();
            break;
        case 2: // Start processing button
            OnStartProcessing();
            break;
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

// Main entry point
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // Initialize common controls
    INITCOMMONCONTROLSEX icex = { sizeof(INITCOMMONCONTROLSEX), ICC_WIN95_CLASSES };
    InitCommonControlsEx(&icex);
    
    // Register window class
    WNDCLASSW wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"BiorhythmCalculatorWindow";
    wc.hCursor = LoadCursorW(nullptr, (LPCWSTR)IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    
    if (!RegisterClassW(&wc)) {
        MessageBoxW(nullptr, L"Window registration failed!", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    // Create main window
    hMainWindow = CreateWindowExW(
        0, wc.lpszClassName, L"Biorhythm Calculator",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX | WS_THICKFRAME | WS_MAXIMIZEBOX,
        CW_USEDEFAULT, CW_USEDEFAULT, 650, 280,
        nullptr, nullptr, hInstance, nullptr
    );
    
    if (!hMainWindow) {
        MessageBoxW(nullptr, L"Window creation failed!", L"Error", MB_OK | MB_ICONERROR);
        return 1;
    }
    
    // Show and update window
    ShowWindow(hMainWindow, nCmdShow);
    UpdateWindow(hMainWindow);
    
    // Message loop
    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    
    return 0;
}
