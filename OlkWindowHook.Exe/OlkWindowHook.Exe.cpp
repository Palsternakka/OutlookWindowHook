/*
*   Outlook Window Hook
*   Keeps Outlook running when main window is closed
*   Copyright (C) 2024  Oliver Dalton
*
*   This program is free software: you can redistribute it and/or modify
*   it under the terms of the GNU General Public License as published by
*   the Free Software Foundation, either version 3 of the License, or
*   (at your option) any later version.
*
*   This program is distributed in the hope that it will be useful,
*   but WITHOUT ANY WARRANTY; without even the implied warranty of
*   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
*   GNU General Public License for more details.

*   You should have received a copy of the GNU General Public License
*   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*/

#include <windows.h>
#include <tlhelp32.h>
#include <tchar.h>
#include <string>
#include <sstream>
#include <thread>
#include <shellapi.h>
#include "resource.h"

#define IDC_OPEN_GITHUB     1000
#define ID_TRAY_APP_ICON    1001
#define ID_TRAY_EXIT        1002
#define ID_TRAY_ABOUT       1003
#define ID_TRAY_AUTOSTART   1004
#define WM_SYSICON          (WM_USER + 1)

typedef void(*SET_HOOK_PROC)();
typedef void(*REMOVE_HOOK_PROC)();

DWORD FindProcessId(const std::wstring& processName) {
    DWORD processId = 0;
    HANDLE hSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (hSnap != INVALID_HANDLE_VALUE) {
        PROCESSENTRY32 pe32{};
        pe32.dwSize = sizeof(PROCESSENTRY32);
        if (Process32First(hSnap, &pe32)) {
            do {
                if (std::wstring(pe32.szExeFile) == processName) {
                    processId = pe32.th32ProcessID;
                    break;
                }
            } while (Process32Next(hSnap, &pe32));
        }
        CloseHandle(hSnap);
    }
    return processId;
}

void InjectDLL(DWORD processID,
    const wchar_t* dllPath) {
    HANDLE hProcess = OpenProcess(PROCESS_ALL_ACCESS, FALSE, processID);
    if (!hProcess) {
        MessageBox(NULL, L"Failed to open target process", L"Outlook Window Hook", MB_ICONERROR);
        return;
    }

    size_t pathLength = (wcslen(dllPath) + 1) * sizeof(wchar_t);
    void* pLibRemote = VirtualAllocEx(hProcess, NULL, pathLength, MEM_COMMIT, PAGE_READWRITE);
    if (!pLibRemote) {
        MessageBox(NULL, L"Failed to allocate memory in target process", L"Outlook Window Hook", MB_ICONERROR);
        CloseHandle(hProcess);
        return;
    }

    WriteProcessMemory(hProcess, pLibRemote, (void*)dllPath, pathLength, NULL);

    HANDLE hThread = CreateRemoteThread(hProcess, NULL, 0, (LPTHREAD_START_ROUTINE)LoadLibraryW, pLibRemote, 0, NULL);
    if (!hThread) {
        MessageBox(NULL, L"Failed to create remote thread", L"Outlook Window Hook", MB_ICONERROR);
        VirtualFreeEx(hProcess, pLibRemote, 0, MEM_RELEASE);
        CloseHandle(hProcess);
        return;
    }

    WaitForSingleObject(hThread, INFINITE);

    VirtualFreeEx(hProcess, pLibRemote, 0, MEM_RELEASE);
    CloseHandle(hThread);
    CloseHandle(hProcess);
}

HINSTANCE hInst;
NOTIFYICONDATA notifyIconData;
HMENU hPopupMenu;
HWND hwnd;
HWND hAboutDlg = NULL;
HMODULE hModule = NULL;
bool appRunning = true;

INT_PTR CALLBACK AboutDlgProc(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
    switch (message) {
    case WM_INITDIALOG:
        hAboutDlg = hDlg;
        return (INT_PTR)TRUE;
    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) {
            EndDialog(hDlg, LOWORD(wParam));
            hAboutDlg = NULL;
            return (INT_PTR)TRUE;
        }
        else if (LOWORD(wParam) == IDC_OPEN_GITHUB) {
            ShellExecute(NULL, L"open", L"https://github.com/Palsternakka/OutlookWindowHook", NULL, NULL, SW_SHOWNORMAL);
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

void ShowAboutDialog(HWND hwnd, HINSTANCE hInst) {
    if (hAboutDlg) {
        SetForegroundWindow(hAboutDlg);
    }
    else {
        DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hwnd, AboutDlgProc);
    }
}

void ManageStartup(bool add) {
    TCHAR szPath[MAX_PATH];
    GetModuleFileName(NULL, szPath, MAX_PATH);

    HKEY hKey;
    LPCTSTR czStartupKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";
    LPCTSTR czValueName = L"Outlook Window Hook";

    if (RegOpenKeyEx(HKEY_CURRENT_USER, czStartupKey, 0, KEY_SET_VALUE, &hKey) == ERROR_SUCCESS) {
        if (add) {
            if (RegSetValueEx(hKey, czValueName, 0, REG_SZ, (LPBYTE)szPath, (_tcslen(szPath) + 1) * sizeof(TCHAR)) == ERROR_SUCCESS) {
                MessageBox(NULL, L"Successfully added to startup!", L"Outlook Window Hook", MB_OK | MB_ICONINFORMATION);
            }
            else {
                MessageBox(NULL, L"Failed to add to startup!", L"Outlook Window Hook", MB_OK | MB_ICONERROR);
            }
        }
        else {
            if (RegDeleteValue(hKey, czValueName) == ERROR_SUCCESS) {
                MessageBox(NULL, L"Successfully removed from startup!", L"Outlook Window Hook", MB_OK | MB_ICONINFORMATION);
            }
            else {
                MessageBox(NULL, L"Failed to remove from startup!", L"Outlook Window Hook", MB_OK | MB_ICONERROR);
            }
        }
        RegCloseKey(hKey);
    }
    else {
        MessageBox(NULL, L"Failed to open registry key!", L"Outlook Window Hook", MB_OK | MB_ICONERROR);
    }
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE:
        notifyIconData.cbSize = sizeof(NOTIFYICONDATA);
        notifyIconData.hWnd = hwnd;
        notifyIconData.uID = ID_TRAY_APP_ICON;
        notifyIconData.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
        notifyIconData.uCallbackMessage = WM_SYSICON;
        notifyIconData.hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_ICON));
        wcscpy_s(notifyIconData.szTip, L"Outlook Window Hook");
        Shell_NotifyIcon(NIM_ADD, &notifyIconData);
        break;
    case WM_SYSICON:
        if (lParam == WM_RBUTTONUP) {
            POINT curPoint;
            GetCursorPos(&curPoint);
            SetForegroundWindow(hwnd);
            TrackPopupMenu(hPopupMenu, TPM_LEFTALIGN | TPM_RIGHTBUTTON, curPoint.x, curPoint.y, 0, hwnd, NULL);
            PostMessage(hwnd, WM_NULL, 0, 0);
        }
        break;
    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_TRAY_EXIT:
            Shell_NotifyIcon(NIM_DELETE, &notifyIconData);
            DestroyWindow(hwnd);
            break;
        case ID_TRAY_ABOUT:
            ShowAboutDialog(hwnd, hInst);
            break;
        case ID_TRAY_AUTOSTART:
            MENUITEMINFO mii = {
              0
            };
            mii.cbSize = sizeof(MENUITEMINFO);
            mii.fMask = MIIM_STATE;
            GetMenuItemInfo(hPopupMenu, ID_TRAY_AUTOSTART, FALSE, &mii);

            if (mii.fState & MFS_CHECKED) {
                mii.fState &= ~MFS_CHECKED;
                ManageStartup(false);
            }
            else {
                mii.fState |= MFS_CHECKED;
                ManageStartup(true);
            }

            SetMenuItemInfo(hPopupMenu, ID_TRAY_AUTOSTART, FALSE, &mii);
            break;
        }
        break;
    case WM_DESTROY:
        Shell_NotifyIcon(NIM_DELETE, &notifyIconData);
        PostQuitMessage(0);
        appRunning = false;
        break;
    default:
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

bool IsInStartup() {
    HKEY hKey;
    LPCTSTR czStartupKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";
    LPCTSTR czValueName = L"Outlook Window Hook";

    if (RegOpenKeyEx(HKEY_CURRENT_USER, czStartupKey, 0, KEY_QUERY_VALUE, &hKey) == ERROR_SUCCESS) {
        DWORD dwType = 0;
        DWORD dwSize = 0;
        LONG lResult = RegQueryValueEx(hKey, czValueName, NULL, &dwType, NULL, &dwSize);
        RegCloseKey(hKey);

        if (lResult == ERROR_SUCCESS) {
            return true;
        }
    }

    return false;
}

void CreateTrayIconMenu() {
    hPopupMenu = CreatePopupMenu();
    AppendMenu(hPopupMenu, MF_STRING, ID_TRAY_ABOUT, L"About");

    bool inStartup = IsInStartup();
    UINT menuState = inStartup ? MF_STRING | MF_CHECKED : MF_STRING | MF_UNCHECKED;
    AppendMenu(hPopupMenu, menuState, ID_TRAY_AUTOSTART, L"Autostart");

    AppendMenu(hPopupMenu, MF_SEPARATOR, 0, NULL);
    AppendMenu(hPopupMenu, MF_STRING, ID_TRAY_EXIT, L"Exit");
}

void MonitorProcess() {
    wchar_t exePath[MAX_PATH];
    GetModuleFileName(NULL, exePath, MAX_PATH);
    std::wstring exeDir = exePath;
    exeDir = exeDir.substr(0, exeDir.find_last_of(L"\\"));

    std::wstringstream dllPathStream;
    dllPathStream << exeDir << L"\\OlkWindowHook.dll";
    std::wstring dllPath = dllPathStream.str();

    DWORD prevProcessID = 0;

    while (appRunning) {
        DWORD processID = FindProcessId(L"olk.exe");

        if (processID != 0 && processID != prevProcessID) {
            InjectDLL(processID, dllPath.c_str());

            if (!hModule) {
                hModule = LoadLibrary(dllPath.c_str());
            }

            if (hModule) {
                SET_HOOK_PROC SetHook = (SET_HOOK_PROC)GetProcAddress(hModule, "SetHook");
                if (SetHook) {
                    SetHook();
                }
                else {
                    MessageBox(NULL, L"Failed to find SetHook function", L"Outlook Window Hook", MB_ICONERROR);
                }
            }
            else {
                MessageBox(NULL, L"Failed to load OlkWindowHook.dll", L"Outlook Window Hook", MB_ICONERROR);
            }

            prevProcessID = processID;
        }

        if (processID == 0) {
            prevProcessID = 0;
        }

        Sleep(500);
    }

    if (hModule) {
        REMOVE_HOOK_PROC RemoveHook = (REMOVE_HOOK_PROC)GetProcAddress(hModule, "RemoveHook");
        if (RemoveHook) {
            RemoveHook();
        }
        FreeLibrary(hModule);
    }
}

int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    const wchar_t* mutexName = L"OlkWindowHook";
    HANDLE hMutex = CreateMutex(NULL, TRUE, mutexName);

    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        MessageBox(NULL, L"Outlook Window Hook is already running", L"Outlook Window Hook", MB_OK | MB_ICONEXCLAMATION);

        CloseHandle(hMutex);

        return 0;
    }
    
    hInst = hInstance;

    WNDCLASSEX wcex{};
    wcex.cbSize = sizeof(WNDCLASSEX);
    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WindowProc;
    wcex.cbClsExtra = 0;
    wcex.cbWndExtra = 0;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON));
    wcex.hCursor = LoadCursor(NULL, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszMenuName = NULL;
    wcex.lpszClassName = L"OlkWindowHookClass";
    wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_ICON));

    RegisterClassEx(&wcex);

    hwnd = CreateWindowEx(0, L"OlkWindowHookClass", L"Outlook Window Hook", 0, 0, 0, 0, 0, HWND_MESSAGE, NULL, hInstance, NULL);

    CreateTrayIconMenu();

    std::thread monitorThread(MonitorProcess);
    monitorThread.detach();

    MSG msg;
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    CloseHandle(hMutex);

    return (int)msg.wParam;
}