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
#include <psapi.h>
#include <tchar.h>
#include <unordered_map>
#include <thread>
#include <mutex>
#include <commctrl.h>

#pragma comment(lib, "Psapi.lib")
#pragma comment(lib, "Comctl32.lib")

HHOOK hHookCallWndProc = NULL;
HINSTANCE hInstance;
std::unordered_map < HWND, bool > ignoreCloseMessage;
std::unordered_map < DWORD, HWND > firstOlkWindowMap;
std::mutex mapMutex;
bool hookSet = false;

BOOL IsOlkExeProcess(HWND hwnd) {
    DWORD procId;
    GetWindowThreadProcessId(hwnd, &procId);

    HANDLE hProcess = OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, FALSE, procId);
    if (hProcess) {
        TCHAR processName[MAX_PATH];
        if (GetModuleBaseName(hProcess, NULL, processName, sizeof(processName) / sizeof(TCHAR))) {
            CloseHandle(hProcess);
            return _tcsicmp(processName, L"olk.exe") == 0;
        }
        CloseHandle(hProcess);
    }
    return FALSE;
}

BOOL IsFirstOlkExeWindow(HWND hwnd) {
    DWORD procId;
    GetWindowThreadProcessId(hwnd, &procId);

    std::lock_guard < std::mutex > lock(mapMutex);
    if (firstOlkWindowMap.find(procId) == firstOlkWindowMap.end()) {
        firstOlkWindowMap[procId] = hwnd;
        return TRUE;
    }
    return firstOlkWindowMap[procId] == hwnd;
}

LRESULT CALLBACK SubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
    if (uMsg == WM_CLOSE) {
        std::lock_guard < std::mutex > lock(mapMutex);
        if (ignoreCloseMessage[hwnd]) {
            ignoreCloseMessage[hwnd] = false;
            return 0;
        }
    }
    return DefSubclassProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK CallWndProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0) {
        CWPSTRUCT* pCwp = (CWPSTRUCT*)lParam;
        if (pCwp->message == WM_CLOSE && IsOlkExeProcess(pCwp->hwnd) && IsFirstOlkExeWindow(pCwp->hwnd)) {
            {
                std::lock_guard < std::mutex > lock(mapMutex);
                if (!ignoreCloseMessage.count(pCwp->hwnd)) {
                    SetWindowSubclass(pCwp->hwnd, SubclassProc, 1, 0);
                }
            }
            ShowWindow(pCwp->hwnd, SW_HIDE);
            std::lock_guard < std::mutex > lock(mapMutex);
            ignoreCloseMessage[pCwp->hwnd] = true;
            return 0;
        }
    }
    return CallNextHookEx(hHookCallWndProc, nCode, wParam, lParam);
}

void KeepDLLLoaded() {
    while (true) {
        Sleep(1000);
    }
}

extern "C"
__declspec(dllexport) void SetHook() {
    std::lock_guard < std::mutex > lock(mapMutex);
    if (!hookSet) {
        hHookCallWndProc = SetWindowsHookEx(WH_CALLWNDPROC, CallWndProc, hInstance, 0);
        if (!hHookCallWndProc) {
            DWORD dwError = GetLastError();
            TCHAR errorMessage[256];
            _stprintf_s(errorMessage, L"Failed to install WH_CALLWNDPROC hook! Error: %d\n", dwError);
            MessageBox(NULL, errorMessage, L"Outlook Window Hook", MB_ICONERROR);
        }
        else {
            hookSet = true;
        }
        std::thread(KeepDLLLoaded).detach();
    }
}

extern "C"
__declspec(dllexport) void RemoveHook() {
    std::lock_guard < std::mutex > lock(mapMutex);
    if (hHookCallWndProc) {
        UnhookWindowsHookEx(hHookCallWndProc);
        hHookCallWndProc = NULL;
        hookSet = false;
    }
}

BOOL APIENTRY DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpReserved) {
    switch (fdwReason) {
    case DLL_PROCESS_ATTACH:
        hInstance = hinstDLL;
        InitCommonControls(); // Initialize common controls
        break;
    }
    return TRUE;
}