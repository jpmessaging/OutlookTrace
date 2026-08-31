#include <winrt/windows.foundation.h>
#include <winrt/windows.security.authentication.web.core.h>
#include <WebAuthenticationCoreManagerInterop.h>
#include <Windows.h>

BOOL APIENTRY DllMain(
    HMODULE hModule,
    DWORD  reason,
    LPVOID /*lpReserved*/)
{
    switch (reason)
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
    IInspectable* const request,
    void** const result)
{
    if (result != nullptr)
    {
        *result = nullptr;
    }

    if (hwnd == nullptr || request == nullptr || result == nullptr)
    {
        return E_INVALIDARG;
    }

    auto webAuthCoreMgrInterop = winrt::get_activation_factory<winrt::WebAuthenticationCoreManager, IWebAuthenticationCoreManagerInterop>();
    auto asyncOp = winrt::IAsyncOperation<winrt::WebTokenRequestResult>{};

    HRESULT hr = webAuthCoreMgrInterop->RequestTokenForWindowAsync(
        hwnd,
        request,
        winrt::guid_of<decltype(asyncOp)>(),
        winrt::put_abi(asyncOp));

    if (SUCCEEDED(hr))
    {
        auto asyncResult = asyncOp.get();
        *result = winrt::detach_abi(asyncResult);
    }

    return hr;
}

/** Create a hidden window that can be used for RequestToken function above. 
Window created by this function must be closed by WM_CLOSE or DestroyWindow() (Use either one, but not both).
The window procedure of this window DOES NOT post WM_QUIT.
 */
extern "C"
__declspec(dllexport)
HWND CreateAnchorWindow()
{
    // Use Immediately Invoked Lambda to ensure the window class is registered only once
    static const ATOM wndClassAtom = []() {
        WNDCLASSEXW wndClass{
        .cbSize = sizeof(WNDCLASSEXW),
        .style = CS_HREDRAW | CS_VREDRAW,
        .lpfnWndProc = DefWindowProcW,
        .cbClsExtra = 0,
        .cbWndExtra = 0,
        .hInstance = GetModuleHandleW(nullptr),
        .hIcon = LoadIcon(NULL, IDI_APPLICATION),
        .hCursor = LoadCursor(NULL, IDC_ARROW),
        .hbrBackground = GetSysColorBrush(COLOR_3DFACE),
        .lpszMenuName = nullptr,
        .lpszClassName = L"AnchorWndClass",
        .hIconSm = nullptr
        };

        return RegisterClassExW(&wndClass);
    }();

    if (0 == wndClassAtom)
    {
        return nullptr;
    }

    // Place at center of desktop
    auto rect = RECT{};
    GetClientRect(GetDesktopWindow(), &rect);
    auto x = static_cast<int>(rect.right / 2);
    auto y = static_cast<int>(rect.bottom / 2);
    auto width = 0;
    auto height = 0;

    auto hwndParent = GetAncestor(GetConsoleWindow(), GA_ROOTOWNER);

    if (nullptr == hwndParent)
    {
        hwndParent = GetForegroundWindow();
    }

    auto hwnd = CreateWindowExW(
        0,
        MAKEINTATOM(wndClassAtom),
        L"Anchor Window",
        WS_POPUP,
        x, y, width, height,
        hwndParent,
        nullptr,
        GetModuleHandleW(nullptr),
        nullptr);

    return hwnd;
}
