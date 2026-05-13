/**
 * comctl32v6hook.dll
 *
 * AppInit_DLLs hook that activates a Common Controls v6 side-by-side
 * activation context around classic Win32 UI creation APIs.
 *
 * Injection mechanism: AppInit_DLLs
 * Hooking library:      Microsoft Detours
 * ActCtx source:        sibling comctl32v6hook.manifest file
 */

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <detours/detours.h>
#include <string>

static HANDLE  g_hActCtx = INVALID_HANDLE_VALUE;
static HMODULE g_hSelf = nullptr;
static bool    g_hooksAttached = false;

struct ActCtxGuard {
    ULONG_PTR cookie = 0;
    bool active = false;

    explicit ActCtxGuard(HANDLE hCtx)
    {
        if (hCtx != INVALID_HANDLE_VALUE && hCtx != nullptr) {
            active = (ActivateActCtx(hCtx, &cookie) != FALSE);
        }
    }

    ~ActCtxGuard()
    {
        if (active) {
            DeactivateActCtx(0, cookie);
        }
    }

    ActCtxGuard(const ActCtxGuard&) = delete;
    ActCtxGuard& operator=(const ActCtxGuard&) = delete;
};

static HANDLE CreateActCtxFromSiblingManifest()
{
    wchar_t dllPath[MAX_PATH] = {};
    if (!GetModuleFileNameW(g_hSelf, dllPath, MAX_PATH)) {
        return INVALID_HANDLE_VALUE;
    }

    std::wstring manifestPath(dllPath);
    const auto dot = manifestPath.rfind(L'.');
    if (dot != std::wstring::npos) {
        manifestPath.resize(dot);
    }
    manifestPath += L".manifest";

    ACTCTXW ctx = {};
    ctx.cbSize = sizeof(ctx);
    ctx.lpSource = manifestPath.c_str();

    HANDLE h = CreateActCtxW(&ctx);
    if (h == INVALID_HANDLE_VALUE) {
        OutputDebugStringW(L"[comctl32v6hook] CreateActCtxW failed for: ");
        OutputDebugStringW(manifestPath.c_str());
    } else {
        OutputDebugStringW(L"[comctl32v6hook] ActCtx created successfully.");
    }

    return h;
}

static decltype(&CreateWindowExW) Real_CreateWindowExW = nullptr;
static decltype(&CreateWindowExA) Real_CreateWindowExA = nullptr;

static decltype(&DialogBoxParamW) Real_DialogBoxParamW = nullptr;
static decltype(&DialogBoxParamA) Real_DialogBoxParamA = nullptr;

static decltype(&DialogBoxIndirectParamW) Real_DialogBoxIndirectParamW = nullptr;
static decltype(&DialogBoxIndirectParamA) Real_DialogBoxIndirectParamA = nullptr;

static decltype(&CreateDialogParamW) Real_CreateDialogParamW = nullptr;
static decltype(&CreateDialogParamA) Real_CreateDialogParamA = nullptr;

static decltype(&CreateDialogIndirectParamW) Real_CreateDialogIndirectParamW = nullptr;
static decltype(&CreateDialogIndirectParamA) Real_CreateDialogIndirectParamA = nullptr;

static decltype(&MessageBoxW) Real_MessageBoxW = nullptr;
static decltype(&MessageBoxA) Real_MessageBoxA = nullptr;
static decltype(&MessageBoxExW) Real_MessageBoxExW = nullptr;
static decltype(&MessageBoxExA) Real_MessageBoxExA = nullptr;
static decltype(&MessageBoxIndirectW) Real_MessageBoxIndirectW = nullptr;
static decltype(&MessageBoxIndirectA) Real_MessageBoxIndirectA = nullptr;

using MessageBoxTimeoutW_t = int (WINAPI *)(HWND, LPCWSTR, LPCWSTR, UINT, WORD, DWORD);
using MessageBoxTimeoutA_t = int (WINAPI *)(HWND, LPCSTR, LPCSTR, UINT, WORD, DWORD);
static MessageBoxTimeoutW_t Real_MessageBoxTimeoutW = nullptr;
static MessageBoxTimeoutA_t Real_MessageBoxTimeoutA = nullptr;

static bool ResolveUser32Functions()
{
    HMODULE hUser32 = GetModuleHandleW(L"user32.dll");
    if (!hUser32) {
        hUser32 = LoadLibraryW(L"user32.dll");
        if (!hUser32) {
            OutputDebugStringW(L"[comctl32v6hook] Failed to get user32.dll handle.");
            return false;
        }
    }

    Real_CreateWindowExW = reinterpret_cast<decltype(Real_CreateWindowExW)>(
        GetProcAddress(hUser32, "CreateWindowExW"));
    Real_CreateWindowExA = reinterpret_cast<decltype(Real_CreateWindowExA)>(
        GetProcAddress(hUser32, "CreateWindowExA"));

    Real_DialogBoxParamW = reinterpret_cast<decltype(Real_DialogBoxParamW)>(
        GetProcAddress(hUser32, "DialogBoxParamW"));
    Real_DialogBoxParamA = reinterpret_cast<decltype(Real_DialogBoxParamA)>(
        GetProcAddress(hUser32, "DialogBoxParamA"));

    Real_DialogBoxIndirectParamW = reinterpret_cast<decltype(Real_DialogBoxIndirectParamW)>(
        GetProcAddress(hUser32, "DialogBoxIndirectParamW"));
    Real_DialogBoxIndirectParamA = reinterpret_cast<decltype(Real_DialogBoxIndirectParamA)>(
        GetProcAddress(hUser32, "DialogBoxIndirectParamA"));

    Real_CreateDialogParamW = reinterpret_cast<decltype(Real_CreateDialogParamW)>(
        GetProcAddress(hUser32, "CreateDialogParamW"));
    Real_CreateDialogParamA = reinterpret_cast<decltype(Real_CreateDialogParamA)>(
        GetProcAddress(hUser32, "CreateDialogParamA"));

    Real_CreateDialogIndirectParamW = reinterpret_cast<decltype(Real_CreateDialogIndirectParamW)>(
        GetProcAddress(hUser32, "CreateDialogIndirectParamW"));
    Real_CreateDialogIndirectParamA = reinterpret_cast<decltype(Real_CreateDialogIndirectParamA)>(
        GetProcAddress(hUser32, "CreateDialogIndirectParamA"));

    Real_MessageBoxW = reinterpret_cast<decltype(Real_MessageBoxW)>(
        GetProcAddress(hUser32, "MessageBoxW"));
    Real_MessageBoxA = reinterpret_cast<decltype(Real_MessageBoxA)>(
        GetProcAddress(hUser32, "MessageBoxA"));

    Real_MessageBoxExW = reinterpret_cast<decltype(Real_MessageBoxExW)>(
        GetProcAddress(hUser32, "MessageBoxExW"));
    Real_MessageBoxExA = reinterpret_cast<decltype(Real_MessageBoxExA)>(
        GetProcAddress(hUser32, "MessageBoxExA"));

    Real_MessageBoxIndirectW = reinterpret_cast<decltype(Real_MessageBoxIndirectW)>(
        GetProcAddress(hUser32, "MessageBoxIndirectW"));
    Real_MessageBoxIndirectA = reinterpret_cast<decltype(Real_MessageBoxIndirectA)>(
        GetProcAddress(hUser32, "MessageBoxIndirectA"));

    Real_MessageBoxTimeoutW = reinterpret_cast<MessageBoxTimeoutW_t>(
        GetProcAddress(hUser32, "MessageBoxTimeoutW"));
    Real_MessageBoxTimeoutA = reinterpret_cast<MessageBoxTimeoutA_t>(
        GetProcAddress(hUser32, "MessageBoxTimeoutA"));

    if (!Real_CreateWindowExW || !Real_CreateWindowExA) {
        OutputDebugStringW(L"[comctl32v6hook] Failed to resolve CreateWindowEx functions.");
        return false;
    }
    if (!Real_DialogBoxParamW || !Real_DialogBoxParamA) {
        OutputDebugStringW(L"[comctl32v6hook] Failed to resolve DialogBoxParam functions.");
        return false;
    }
    if (!Real_DialogBoxIndirectParamW || !Real_DialogBoxIndirectParamA) {
        OutputDebugStringW(L"[comctl32v6hook] Failed to resolve DialogBoxIndirectParam functions.");
        return false;
    }
    if (!Real_CreateDialogParamW || !Real_CreateDialogParamA) {
        OutputDebugStringW(L"[comctl32v6hook] Failed to resolve CreateDialogParam functions.");
        return false;
    }
    if (!Real_CreateDialogIndirectParamW || !Real_CreateDialogIndirectParamA) {
        OutputDebugStringW(L"[comctl32v6hook] Failed to resolve CreateDialogIndirectParam functions.");
        return false;
    }
    if (!Real_MessageBoxW || !Real_MessageBoxA) {
        OutputDebugStringW(L"[comctl32v6hook] Failed to resolve MessageBox functions.");
        return false;
    }
    if (!Real_MessageBoxExW || !Real_MessageBoxExA) {
        OutputDebugStringW(L"[comctl32v6hook] Failed to resolve MessageBoxEx functions.");
        return false;
    }
    if (!Real_MessageBoxIndirectW || !Real_MessageBoxIndirectA) {
        OutputDebugStringW(L"[comctl32v6hook] Failed to resolve MessageBoxIndirect functions.");
        return false;
    }

    OutputDebugStringW(L"[comctl32v6hook] All user32 functions resolved successfully.");
    return true;
}

HWND WINAPI Hook_CreateWindowExW(
    DWORD dwExStyle, LPCWSTR lpClassName, LPCWSTR lpWindowName,
    DWORD dwStyle, int X, int Y, int nWidth, int nHeight,
    HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_CreateWindowExW(dwExStyle, lpClassName, lpWindowName,
        dwStyle, X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam);
}

HWND WINAPI Hook_CreateWindowExA(
    DWORD dwExStyle, LPCSTR lpClassName, LPCSTR lpWindowName,
    DWORD dwStyle, int X, int Y, int nWidth, int nHeight,
    HWND hWndParent, HMENU hMenu, HINSTANCE hInstance, LPVOID lpParam)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_CreateWindowExA(dwExStyle, lpClassName, lpWindowName,
        dwStyle, X, Y, nWidth, nHeight, hWndParent, hMenu, hInstance, lpParam);
}

INT_PTR WINAPI Hook_DialogBoxParamW(
    HINSTANCE hInstance, LPCWSTR lpTemplateName,
    HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_DialogBoxParamW(hInstance, lpTemplateName,
        hWndParent, lpDialogFunc, dwInitParam);
}

INT_PTR WINAPI Hook_DialogBoxParamA(
    HINSTANCE hInstance, LPCSTR lpTemplateName,
    HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_DialogBoxParamA(hInstance, lpTemplateName,
        hWndParent, lpDialogFunc, dwInitParam);
}

INT_PTR WINAPI Hook_DialogBoxIndirectParamW(
    HINSTANCE hInstance, LPCDLGTEMPLATEW hDialogTemplate,
    HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_DialogBoxIndirectParamW(hInstance, hDialogTemplate,
        hWndParent, lpDialogFunc, dwInitParam);
}

INT_PTR WINAPI Hook_DialogBoxIndirectParamA(
    HINSTANCE hInstance, LPCDLGTEMPLATEA hDialogTemplate,
    HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_DialogBoxIndirectParamA(hInstance, hDialogTemplate,
        hWndParent, lpDialogFunc, dwInitParam);
}

HWND WINAPI Hook_CreateDialogParamW(
    HINSTANCE hInstance, LPCWSTR lpTemplateName,
    HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_CreateDialogParamW(hInstance, lpTemplateName,
        hWndParent, lpDialogFunc, dwInitParam);
}

HWND WINAPI Hook_CreateDialogParamA(
    HINSTANCE hInstance, LPCSTR lpTemplateName,
    HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_CreateDialogParamA(hInstance, lpTemplateName,
        hWndParent, lpDialogFunc, dwInitParam);
}

HWND WINAPI Hook_CreateDialogIndirectParamW(
    HINSTANCE hInstance, LPCDLGTEMPLATEW lpTemplate,
    HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_CreateDialogIndirectParamW(hInstance, lpTemplate,
        hWndParent, lpDialogFunc, dwInitParam);
}

HWND WINAPI Hook_CreateDialogIndirectParamA(
    HINSTANCE hInstance, LPCDLGTEMPLATEA lpTemplate,
    HWND hWndParent, DLGPROC lpDialogFunc, LPARAM dwInitParam)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_CreateDialogIndirectParamA(hInstance, lpTemplate,
        hWndParent, lpDialogFunc, dwInitParam);
}

int WINAPI Hook_MessageBoxW(HWND hWnd, LPCWSTR lpText, LPCWSTR lpCaption, UINT uType)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_MessageBoxW(hWnd, lpText, lpCaption, uType);
}

int WINAPI Hook_MessageBoxA(HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_MessageBoxA(hWnd, lpText, lpCaption, uType);
}

int WINAPI Hook_MessageBoxExW(
    HWND hWnd, LPCWSTR lpText, LPCWSTR lpCaption, UINT uType, WORD wLanguageId)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_MessageBoxExW(hWnd, lpText, lpCaption, uType, wLanguageId);
}

int WINAPI Hook_MessageBoxExA(
    HWND hWnd, LPCSTR lpText, LPCSTR lpCaption, UINT uType, WORD wLanguageId)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_MessageBoxExA(hWnd, lpText, lpCaption, uType, wLanguageId);
}

int WINAPI Hook_MessageBoxIndirectW(const MSGBOXPARAMSW* lpmbp)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_MessageBoxIndirectW(lpmbp);
}

int WINAPI Hook_MessageBoxIndirectA(const MSGBOXPARAMSA* lpmbp)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_MessageBoxIndirectA(lpmbp);
}

int WINAPI Hook_MessageBoxTimeoutW(
    HWND hWnd, LPCWSTR lpText, LPCWSTR lpCaption,
    UINT uType, WORD wLanguageId, DWORD dwMilliseconds)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_MessageBoxTimeoutW(
        hWnd, lpText, lpCaption, uType, wLanguageId, dwMilliseconds);
}

int WINAPI Hook_MessageBoxTimeoutA(
    HWND hWnd, LPCSTR lpText, LPCSTR lpCaption,
    UINT uType, WORD wLanguageId, DWORD dwMilliseconds)
{
    ActCtxGuard guard(g_hActCtx);
    return Real_MessageBoxTimeoutA(
        hWnd, lpText, lpCaption, uType, wLanguageId, dwMilliseconds);
}

#define ATTACH(Real, Hook) DetourAttach(reinterpret_cast<PVOID*>(&Real), reinterpret_cast<PVOID>(Hook))
#define DETACH(Real, Hook) DetourDetach(reinterpret_cast<PVOID*>(&Real), reinterpret_cast<PVOID>(Hook))
#define ATTACH_IF(Real, Hook) do { if (Real) { ATTACH(Real, Hook); } } while (0)
#define DETACH_IF(Real, Hook) do { if (Real) { DETACH(Real, Hook); } } while (0)

static LONG AttachHooks()
{
    LONG error = DetourTransactionBegin();
    if (error != NO_ERROR) {
        OutputDebugStringW(L"[comctl32v6hook] DetourTransactionBegin failed.");
        return error;
    }

    DetourUpdateThread(GetCurrentThread());

    ATTACH(Real_CreateWindowExW,            Hook_CreateWindowExW);
    ATTACH(Real_CreateWindowExA,            Hook_CreateWindowExA);
    ATTACH(Real_DialogBoxParamW,            Hook_DialogBoxParamW);
    ATTACH(Real_DialogBoxParamA,            Hook_DialogBoxParamA);
    ATTACH(Real_DialogBoxIndirectParamW,    Hook_DialogBoxIndirectParamW);
    ATTACH(Real_DialogBoxIndirectParamA,    Hook_DialogBoxIndirectParamA);
    ATTACH(Real_CreateDialogParamW,         Hook_CreateDialogParamW);
    ATTACH(Real_CreateDialogParamA,         Hook_CreateDialogParamA);
    ATTACH(Real_CreateDialogIndirectParamW, Hook_CreateDialogIndirectParamW);
    ATTACH(Real_CreateDialogIndirectParamA, Hook_CreateDialogIndirectParamA);
    ATTACH(Real_MessageBoxW,                Hook_MessageBoxW);
    ATTACH(Real_MessageBoxA,                Hook_MessageBoxA);
    ATTACH(Real_MessageBoxExW,              Hook_MessageBoxExW);
    ATTACH(Real_MessageBoxExA,              Hook_MessageBoxExA);
    ATTACH(Real_MessageBoxIndirectW,        Hook_MessageBoxIndirectW);
    ATTACH(Real_MessageBoxIndirectA,        Hook_MessageBoxIndirectA);
    ATTACH_IF(Real_MessageBoxTimeoutW,      Hook_MessageBoxTimeoutW);
    ATTACH_IF(Real_MessageBoxTimeoutA,      Hook_MessageBoxTimeoutA);

    error = DetourTransactionCommit();
    if (error != NO_ERROR) {
        OutputDebugStringW(L"[comctl32v6hook] DetourTransactionCommit (attach) failed.");
    } else {
        OutputDebugStringW(L"[comctl32v6hook] All hooks attached successfully.");
    }

    return error;
}

static LONG DetachHooks()
{
    LONG error = DetourTransactionBegin();
    if (error != NO_ERROR) {
        return error;
    }

    DetourUpdateThread(GetCurrentThread());

    DETACH(Real_CreateWindowExW,            Hook_CreateWindowExW);
    DETACH(Real_CreateWindowExA,            Hook_CreateWindowExA);
    DETACH(Real_DialogBoxParamW,            Hook_DialogBoxParamW);
    DETACH(Real_DialogBoxParamA,            Hook_DialogBoxParamA);
    DETACH(Real_DialogBoxIndirectParamW,    Hook_DialogBoxIndirectParamW);
    DETACH(Real_DialogBoxIndirectParamA,    Hook_DialogBoxIndirectParamA);
    DETACH(Real_CreateDialogParamW,         Hook_CreateDialogParamW);
    DETACH(Real_CreateDialogParamA,         Hook_CreateDialogParamA);
    DETACH(Real_CreateDialogIndirectParamW, Hook_CreateDialogIndirectParamW);
    DETACH(Real_CreateDialogIndirectParamA, Hook_CreateDialogIndirectParamA);
    DETACH(Real_MessageBoxW,                Hook_MessageBoxW);
    DETACH(Real_MessageBoxA,                Hook_MessageBoxA);
    DETACH(Real_MessageBoxExW,              Hook_MessageBoxExW);
    DETACH(Real_MessageBoxExA,              Hook_MessageBoxExA);
    DETACH(Real_MessageBoxIndirectW,        Hook_MessageBoxIndirectW);
    DETACH(Real_MessageBoxIndirectA,        Hook_MessageBoxIndirectA);
    DETACH_IF(Real_MessageBoxTimeoutW,      Hook_MessageBoxTimeoutW);
    DETACH_IF(Real_MessageBoxTimeoutA,      Hook_MessageBoxTimeoutA);

    return DetourTransactionCommit();
}

BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID)
{
    if (DetourIsHelperProcess()) {
        return TRUE;
    }

    switch (fdwReason) {
    case DLL_PROCESS_ATTACH:
        g_hSelf = hinstDLL;
        DisableThreadLibraryCalls(hinstDLL);

        if (!ResolveUser32Functions()) {
            OutputDebugStringW(L"[comctl32v6hook] Function resolution failed; hooks not installed.");
            break;
        }

        g_hActCtx = CreateActCtxFromSiblingManifest();
        g_hooksAttached = (AttachHooks() == NO_ERROR);
        break;

    case DLL_PROCESS_DETACH:
        if (g_hooksAttached) {
            DetachHooks();
            g_hooksAttached = false;
        }

        if (g_hActCtx != INVALID_HANDLE_VALUE && g_hActCtx != nullptr) {
            ReleaseActCtx(g_hActCtx);
            g_hActCtx = INVALID_HANDLE_VALUE;
        }
        break;
    }

    return TRUE;
}
