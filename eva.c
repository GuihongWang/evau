#include <stdio.h>
#include <stdlib.h>
#include <windows.h>
#include <shlobj.h>
#include <string.h>

#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "advapi32.lib")

// 函数声明
BOOL IsUserAdmin();
void ShowMenu();
void ExportEventLogs();
void OpenWindowsUpdate();
void ClearUpdateCache();
BOOL MyGetTempPath(LPSTR tempPath);
BOOL GetDesktopPath(LPSTR desktopPath);
void WaitForAnyKey();
void ClearInputBuffer();
int GetIntInput(const char* prompt);
void DeleteDirectoryRecursive(LPCSTR dirPath);

int main() {
    int choice;
    
    // 检查管理员权限
    if (!IsUserAdmin()) {
        printf("错误: 此程序需要管理员权限才能运行!\n");
        printf("请右键点击程序并选择'以管理员身份运行'\n");
        WaitForAnyKey();
        return 1;
    }
    
    printf("欢迎使用Windows系统管理工具\n");
    printf("==============================\n");
    
    do {
        ShowMenu();
        choice = GetIntInput("请选择操作然后按两下回车");
        
        switch(choice) {
            case 1:
                ExportEventLogs();
                break;
            case 2:
                OpenWindowsUpdate();
                break;
            case 3:
                ClearUpdateCache();
                break;
            case 4:
                printf("感谢使用，再见!\n");
                break;
            default:
                printf("无效选择，请重新输入!\n");
                WaitForAnyKey();
        }
    } while(choice != 4);
    
    return 0;
}

// 显示主菜单
void ShowMenu() {
    printf("\n===== 系统管理菜单 =====\n");
    printf("1. 导出事件日志文件\n");
    printf("2. 打开Windows更新设置\n");
    printf("3. 清理Windows更新缓存\n");
    printf("4. 退出\n");
    printf("========================\n");
}

// 导出事件日志文件
void ExportEventLogs() {
    char systemRoot[MAX_PATH];
    char sourcePath[MAX_PATH];
    char tempPath[MAX_PATH];
    char desktopPath[MAX_PATH];
    char zipPath[MAX_PATH];
    char tempFile1[MAX_PATH];
    char tempFile2[MAX_PATH];
    
    // 获取系统根目录
    if (!GetEnvironmentVariableA("SystemRoot", systemRoot, MAX_PATH)) {
        printf("无法获取系统根目录\n");
        WaitForAnyKey();
        return;
    }
    
    // 获取临时目录路径
    if (!MyGetTempPath(tempPath)) {
        printf("无法获取临时目录路径\n");
        WaitForAnyKey();
        return;
    }
    
    // 获取桌面路径
    if (!GetDesktopPath(desktopPath)) {
        printf("无法获取桌面路径\n");
        WaitForAnyKey();
        return;
    }
    
    // 构建源文件路径
    snprintf(sourcePath, MAX_PATH, "%s\\System32\\Winevt\\Logs\\Application.evtx", systemRoot);
    
    // 构建临时文件路径
    snprintf(tempFile1, MAX_PATH, "%s\\Application.evtx", tempPath);
    snprintf(tempFile2, MAX_PATH, "%s\\System.evtx", tempPath);
    
    // 复制Application.evtx到临时目录
    if (!CopyFileA(sourcePath, tempFile1, FALSE)) {
        printf("无法复制Application.evtx到临时目录\n");
        WaitForAnyKey();
        return;
    }
    printf("已复制Application.evtx到临时目录\n");
    
    // 复制System.evtx到临时目录
    snprintf(sourcePath, MAX_PATH, "%s\\System32\\Winevt\\Logs\\System.evtx", systemRoot);
    if (!CopyFileA(sourcePath, tempFile2, FALSE)) {
        printf("无法复制System.evtx到临时目录\n");
        WaitForAnyKey();
        return;
    }
    printf("已复制System.evtx到临时目录\n");
    
    // 构建ZIP文件路径
    snprintf(zipPath, MAX_PATH, "%s\\EventLogs.zip", desktopPath);
    
    // 创建ZIP文件 - 使用PowerShell
    char command[MAX_PATH * 3];
    snprintf(command, sizeof(command), 
             "powershell -Command \"Compress-Archive -Path '%s','%s' -DestinationPath '%s' -Force\"", 
             tempFile1, tempFile2, zipPath);
    
    int result = system(command);
    if (result == 0) {
        printf("已创建ZIP文件到桌面: %s\n", zipPath);
    } else {
        printf("无法创建ZIP文件\n");
    }
    
    // 清理临时文件
    DeleteFileA(tempFile1);
    DeleteFileA(tempFile2);
    
    WaitForAnyKey();
}

// 打开Windows更新设置 - 使用ShellExecute避免WSL问题
void OpenWindowsUpdate() {
    // 检查Windows版本
    OSVERSIONINFOEX osvi;
    ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));
    osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
    
    if (GetVersionEx((OSVERSIONINFO*)&osvi)) {
        if (osvi.dwMajorVersion == 10) {
            printf("您正在使用Windows 10\n");
            printf("Windows 10只检查Windows 10的更新，无需更新到Windows 11\n");
        } else if (osvi.dwMajorVersion == 11) {
            printf("您正在使用Windows 11\n");
        } else {
            printf("您正在使用Windows %lu.%lu\n", osvi.dwMajorVersion, osvi.dwMinorVersion);
        }
        printf("按任意键继续打开Windows更新设置...\n");
        WaitForAnyKey();
    }
    
    // 使用ShellExecuteA打开Windows更新设置（兼容WSL和原生Windows）
    HINSTANCE result = ShellExecuteA(NULL, "open", "ms-settings:windowsupdate", NULL, NULL, SW_SHOWNORMAL);
    
    // 正确的检查方式：使用INT_PTR进行转换以避免警告
    if ((INT_PTR)result > 32) {
        printf("已打开Windows更新设置\n");
    } else {
        DWORD error = (DWORD)(INT_PTR)result;
        printf("打开Windows更新设置失败，错误代码: %lu\n", error);
        printf("尝试使用备用方法...\n");
        
        // 备用方法：使用start命令
        system("start ms-settings:windowsupdate");
    }
}

// 递归删除目录及其内容
void DeleteDirectoryRecursive(LPCSTR dirPath) {
    char searchPath[MAX_PATH * 2];
    WIN32_FIND_DATAA findData;
    HANDLE hFind;
    
    // 构建搜索路径
    snprintf(searchPath, sizeof(searchPath), "%s\\*", dirPath);
    
    // 查找第一个文件
    hFind = FindFirstFileA(searchPath, &findData);
    if (hFind == INVALID_HANDLE_VALUE) {
        return;
    }
    
    do {
        // 跳过 . 和 ..
        if (strcmp(findData.cFileName, ".") == 0 || strcmp(findData.cFileName, "..") == 0) {
            continue;
        }
        
        char fullPath[MAX_PATH * 2];
        snprintf(fullPath, sizeof(fullPath), "%s\\%s", dirPath, findData.cFileName);
        
        if (findData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) {
            // 递归删除子目录
            DeleteDirectoryRecursive(fullPath);
            RemoveDirectoryA(fullPath);
        } else {
            // 删除文件
            DeleteFileA(fullPath);
        }
    } while (FindNextFileA(hFind, &findData));
    
    FindClose(hFind);
}

// 清理Windows更新缓存
void ClearUpdateCache() {
    SC_HANDLE scManager;
    SC_HANDLE service;
    SERVICE_STATUS serviceStatus;
    char windir[MAX_PATH];
    char downloadPath[MAX_PATH];
    BOOL serviceWasRunning = FALSE; // 记录服务原本是否在运行
    
    // 获取Windows目录
    if (!GetEnvironmentVariableA("windir", windir, MAX_PATH)) {
        printf("无法获取Windows目录\n");
        WaitForAnyKey();
        return;
    }
    
    // 构建SoftwareDistribution路径
    snprintf(downloadPath, MAX_PATH, "%s\\SoftwareDistribution\\Download", windir);
    
    // 打开服务控制管理器
    scManager = OpenSCManager(NULL, NULL, SC_MANAGER_CONNECT);
    if (scManager == NULL) {
        printf("无法打开服务控制管理器\n");
        WaitForAnyKey();
        return;
    }
    
    // 打开wuauserv服务
    service = OpenService(scManager, "wuauserv", SERVICE_QUERY_STATUS | SERVICE_STOP | SERVICE_START);
    if (service == NULL) {
        printf("无法打开wuauserv服务\n");
        CloseServiceHandle(scManager);
        WaitForAnyKey();
        return;
    }
    
    // 查询服务状态
    if (!QueryServiceStatus(service, &serviceStatus)) {
        printf("无法查询服务状态\n");
        CloseServiceHandle(service);
        CloseServiceHandle(scManager);
        WaitForAnyKey();
        return;
    }
    
    // 记录服务原本是否在运行
    if (serviceStatus.dwCurrentState == SERVICE_RUNNING) {
        serviceWasRunning = TRUE;
    }
    
    // 如果服务正在运行，停止它
    if (serviceWasRunning) {
        printf("wuauserv服务正在运行，正在停止服务...\n");
        if (!ControlService(service, SERVICE_CONTROL_STOP, &serviceStatus)) {
            printf("无法停止服务\n");
            CloseServiceHandle(service);
            CloseServiceHandle(scManager);
            WaitForAnyKey();
            return;
        }
        
        // 等待服务停止
        int timeout = 10; // 10秒超时
        while (timeout > 0) {
            QueryServiceStatus(service, &serviceStatus);
            if (serviceStatus.dwCurrentState == SERVICE_STOPPED) {
                break;
            }
            Sleep(1000);
            timeout--;
        }
        
        if (timeout == 0) {
            printf("服务停止超时\n");
            CloseServiceHandle(service);
            CloseServiceHandle(scManager);
            WaitForAnyKey();
            return;
        }
        printf("服务已停止\n");
    } else {
        printf("wuauserv服务未运行，直接清理缓存\n");
    }
    
    // 删除Download目录中的所有文件和子目录
    printf("正在清理更新缓存...\n");
    
    // 使用递归删除函数
    DeleteDirectoryRecursive(downloadPath);
    
    // 重新创建Download目录
    if (!CreateDirectoryA(downloadPath, NULL)) {
        // 如果目录已存在，也没关系
        DWORD error = GetLastError();
        if (error != ERROR_ALREADY_EXISTS) {
            printf("无法创建Download目录，错误代码: %lu\n", error);
        }
    }
    
    // 只有在服务原本运行的情况下才重新启动服务
    if (serviceWasRunning) {
        printf("正在重新启动wuauserv服务...\n");
        if (!StartService(service, 0, NULL)) {
            DWORD error = GetLastError();
            if (error != ERROR_SERVICE_ALREADY_RUNNING) {
                printf("无法启动服务，错误代码: %lu\n", error);
            } else {
                printf("服务已经在运行\n");
            }
        } else {
            printf("服务已成功启动\n");
        }
    } else {
        printf("服务原本未运行，清理完成后保持停止状态\n");
    }
    
    printf("更新缓存清理完成\n");
    
    // 清理资源
    CloseServiceHandle(service);
    CloseServiceHandle(scManager);
    
    WaitForAnyKey();
}

// 检查用户是否具有管理员权限
BOOL IsUserAdmin() {
    BOOL isAdmin = FALSE;
    SID_IDENTIFIER_AUTHORITY ntAuthority = SECURITY_NT_AUTHORITY;
    PSID administratorsGroup = NULL;
    
    if (AllocateAndInitializeSid(
        &ntAuthority,
        2,
        SECURITY_BUILTIN_DOMAIN_RID,
        DOMAIN_ALIAS_RID_ADMINS,
        0, 0, 0, 0, 0, 0,
        &administratorsGroup)) {
        
        if (!CheckTokenMembership(NULL, administratorsGroup, &isAdmin)) {
            isAdmin = FALSE;
        }
        
        FreeSid(administratorsGroup);
    }
    
    return isAdmin;
}

// 获取临时目录路径
BOOL MyGetTempPath(LPSTR tempPath) {
    DWORD result = GetTempPathA(MAX_PATH, tempPath);
    return result > 0 && result <= MAX_PATH;
}

// 获取桌面路径
BOOL GetDesktopPath(LPSTR desktopPath) {
    return SHGetFolderPathA(NULL, CSIDL_DESKTOPDIRECTORY, NULL, 0, desktopPath) == S_OK;
}

// 等待任意键按下
void WaitForAnyKey() {
    printf("\n按任意键继续...\n");
    
    HANDLE hStdin = GetStdHandle(STD_INPUT_HANDLE);
    if (hStdin == INVALID_HANDLE_VALUE) {
        getchar();
        return;
    }
    
    FlushConsoleInputBuffer(hStdin);
    
    INPUT_RECORD ir;
    DWORD count;
    while (ReadConsoleInput(hStdin, &ir, 1, &count)) {
        if (ir.EventType == KEY_EVENT && ir.Event.KeyEvent.bKeyDown) {
            break;
        }
    }
}

// 清空输入缓冲区
void ClearInputBuffer() {
    int c;
    while ((c = getchar()) != '\n' && c != EOF);
}

// 获取整数输入
int GetIntInput(const char* prompt) {
    int value;
    char buffer[100];
    
    while (1) {
        printf("%s: ", prompt);
        if (fgets(buffer, sizeof(buffer), stdin) == NULL) {
            continue;
        }
        
        if (sscanf(buffer, "%d", &value) == 1) {
            ClearInputBuffer();
            return value;
        }
        
        printf("无效输入，请输入整数!\n");
    }
}
