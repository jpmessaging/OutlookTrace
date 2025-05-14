#include "pch.h"

BOOL APIENTRY DllMain(
    HMODULE hModule,
    DWORD  ul_reason_for_call,
    [[maybe_unused]] LPVOID lpReserved)
{
    switch (ul_reason_for_call)
    {
    case DLL_PROCESS_ATTACH:
    {
        DisableThreadLibraryCalls(hModule);
        break;
    }
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }

    return TRUE;
}

namespace winrt
{
    using namespace winrt::Windows::Foundation;
    using namespace winrt::Windows::Security::Authentication::Web::Core;
    using namespace winrt::Windows::Security::Credentials;
}

extern "C"
__declspec(dllexport)
HRESULT RequestToken(
    const HWND hwnd,
    IInspectable* request,
    void** result)
{
    if (hwnd == nullptr || request == nullptr || result == nullptr)
    {
        return E_INVALIDARG;
    }

    auto webAuthCoreMgrInterop = winrt::get_activation_factory<winrt::WebAuthenticationCoreManager, IWebAuthenticationCoreManagerInterop>();
    auto asyncOp = winrt::IAsyncOperation<winrt::WebTokenRequestResult>{};

    auto hr = webAuthCoreMgrInterop->RequestTokenForWindowAsync(
        hwnd,
        request,
        winrt::guid_of<decltype(asyncOp)>(),
        winrt::put_abi(asyncOp)
    );

    if (FAILED(hr))
    {
        return hr;
    }

    auto asyncResult = asyncOp.get();
    *result = winrt::detach_abi(asyncResult);

    return S_OK;
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;

    default:
        return DefWindowProc(hwnd, msg, wParam, lParam);
    }
}

extern "C"
__declspec(dllexport)
HWND CreateAnchorWindow()
{
    WNDCLASSEX wndclass{
    .cbSize = sizeof(WNDCLASSEX),
    .style = CS_HREDRAW | CS_VREDRAW,
    .lpfnWndProc = WndProc,
    .cbClsExtra = 0,
    .cbWndExtra = 0,
    .hInstance = GetModuleHandleW(nullptr),
    .hIcon = LoadIcon(NULL, IDI_APPLICATION),
    .hCursor = LoadCursor(NULL, IDC_ARROW),
    .hbrBackground = GetSysColorBrush(COLOR_3DFACE),
    .lpszMenuName = nullptr,
    .lpszClassName = L"WndClass",
    .hIconSm = nullptr
    };

    RegisterClassExW(&wndclass);

    // Place at center of desktop
    auto rect = RECT{};
    GetClientRect(GetDesktopWindow(), &rect);
    auto x = rect.right / 2;
    auto y = rect.bottom / 2;
    auto width = 0;
    auto height = 0;

    auto hwndConsole = GetAncestor(GetConsoleWindow(), GA_ROOTOWNER);

    auto hwnd = CreateWindowExW(
        0,
        wndclass.lpszClassName,
        L"Anchor Window",
        WS_POPUP,
        x, y, width, height,
        hwndConsole, // hWndParent
        nullptr,
        GetModuleHandle(nullptr),
        nullptr);

    return hwnd;
}
