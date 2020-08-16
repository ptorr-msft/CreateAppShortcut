#include <ShlObj.h>
#include <ShObjIdl_core.h>
#include <propvarutil.h>
#include <propkey.h>
#include <memory>
#include <string>
#include <iostream>
#include <wrl/client.h>
#include <conio.h>

using namespace Microsoft::WRL;

static const int exit_success{ 0 };
static const int exit_usage{ 1 };
static const int exit_user_abort{ 2 };
static const int exit_no_app{ 3 };
static const int exit_runtime_error{ 4 };

void check(HRESULT hr)
{
    if (SUCCEEDED(hr))
    {
        return;
    }

    std::wcerr << L"Error " << hr << L" (" << GetLastError() << L")." << std::endl;
    exit(exit_runtime_error);
}

const wchar_t* GetPropertyStringRef(IPropertyStore* store, const PROPERTYKEY& key)
{
    static wchar_t buffer[MAX_PATH];
    PROPVARIANT var{};
    PropVariantInit(&var);

    store->GetValue(key, &var);
    PropVariantToString(var, buffer, ARRAYSIZE(buffer));
    PropVariantClear(&var);

    return buffer;
}

struct FreeIdList
{
    void operator()(void* p)
    {
        auto idlist{ static_cast<LPITEMIDLIST>(p) };
        ILFree(idlist);
    }
};

typedef std::unique_ptr<__unaligned ITEMIDLIST, FreeIdList> IDListPtr;

int wmain(int argc, wchar_t* argv[])
{

    std::wstring appName{};
    if (argc > 1)
    {
        if (*argv[1] == L'/' || *argv[1] == L'-')
        {
            std::wcout << L"Usage: mkapplnk [appname-substring [output-file] ]" << std::endl;
            std::wcout << L"Creates shortcuts to any app in the apps folder, including UWP / MSIX apps." << std::endl;
            return exit_usage;
        }

        appName = argv[1];
    }
    else
    {
        std::wcout << L"App name to search for: ";
        std::getline(std::wcin, appName);
    }

    if (appName.length() == 0)
    {
        return exit_user_abort;
    }

    std::wcout << L"Searching for '" << appName << "' . . ." << std::endl;

    check(RoInitialize(RO_INIT_SINGLETHREADED));

    // Get the Desktop folder; this is the root of all shell folders.
    ComPtr<IShellFolder> desktopFolder{ nullptr };
    check(SHGetDesktopFolder(&desktopFolder));

    // Get the ITEMIDLIST of the AppsFolder. We use a temp IDLIST and then copy the pointers into our RAII type.
    LPITEMIDLIST tempIDList{ nullptr };
    check(SHGetKnownFolderIDList(FOLDERID_AppsFolder, 0, nullptr, &tempIDList));
    IDListPtr appsKnownFolderIDList{ tempIDList };

    // Convert the AppsFolder ITEMIDLIST into an IShellFolder2.
    ComPtr<IShellFolder2> appsKnownFolderShellFolder{ nullptr };
    check(desktopFolder->BindToObject(appsKnownFolderIDList.get(), nullptr, IID_PPV_ARGS(&appsKnownFolderShellFolder)));

    wchar_t foundAppName[MAX_PATH]{ 0 };
    IDListPtr foundAppAbsoluteIDList{ nullptr };

    // Enumerate all the children of the AppsFolder.
    ComPtr<IEnumIDList> appsFolderChildEnumerator{ nullptr };
    check(appsKnownFolderShellFolder->EnumObjects(nullptr, SHCONTF_FASTITEMS | SHCONTF_NONFOLDERS, &appsFolderChildEnumerator));
    while (S_OK == appsFolderChildEnumerator->Next(1, &tempIDList, nullptr))
    {
        IDListPtr childItemIDList{ tempIDList };

        // Ask the AppsFolder for a PropertyStore factory for the current item.
        ComPtr<IPropertyStoreFactory> propertyStoreFactoryForChild{ nullptr };
        check(appsKnownFolderShellFolder->BindToObject(childItemIDList.get(), nullptr, IID_PPV_ARGS(&propertyStoreFactoryForChild)));

        // Get a PropertyStore from that factory.
        ComPtr<IPropertyStore> propertyStoreForChild{ nullptr };
        check(propertyStoreFactoryForChild->GetPropertyStore(GPS_DEFAULT, nullptr, IID_PPV_ARGS(&propertyStoreForChild)));

        // Get the ItemNameDisplay property to see if it matches what the user asked for.
        auto thisAppDisplayName{ GetPropertyStringRef(propertyStoreForChild.Get(), PKEY_ItemNameDisplay) };
        if (StrStrIW(thisAppDisplayName, appName.c_str()) != nullptr)
        {
            std::wcout << std::endl << L"Found '" << thisAppDisplayName << L"' (";
            std::wcout << GetPropertyStringRef(propertyStoreForChild.Get(), PKEY_AppUserModel_ID) << L")." << std::endl;
            std::wcout << L"Use this app? [Y/N; blank for Y] ";
            auto answer{ toupper(_getwch()) };
            std::wcout << static_cast<wchar_t>(answer) << std::endl;
            if (answer != L'N')
            {
                // The ITIDLIST we got is relative to the AppsFolder; we need an absolute ITEMIDLIST to save in the shortcut.
                tempIDList = ILCombine(appsKnownFolderIDList.get(), childItemIDList.get());
                foundAppAbsoluteIDList.reset(tempIDList);
                wcscpy_s(foundAppName, thisAppDisplayName);
                break;
            }
        }
    }

    std::wcout << std::endl;

    if (wcslen(foundAppName) == 0)
    {
        std::wcout << L"Did not find an app with that name." << std::endl;
        return exit_no_app;
    }

    std::wstring fileName{};
    if (argc > 2)
    {
        fileName = argv[2];
    }
    else
    {
        std::wcout << L"Filename to save as (<enter> to skip): ";
        std::getline(std::wcin, fileName);
    }

    if (fileName.length() == 0)
    {
        return exit_user_abort;
    }

    // Create the IShellLink and set the ITEMIDLIST and name.
    ComPtr<IShellLinkW> shellLink{ nullptr };
    check(CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(shellLink.GetAddressOf())));
    check(shellLink->SetDescription(foundAppName));
    check(shellLink->SetIDList(foundAppAbsoluteIDList.get()));

    // Save the shortcut to the provided filename.
    ComPtr<IPersistFile> file{ nullptr };
    check(shellLink->QueryInterface(file.GetAddressOf()));

    check(file->Save(fileName.c_str(), FALSE));

    std::wcout << std::endl;
    std::wcout << L"Success. Created shortcut to '" << foundAppName << L"' at " << fileName << L"." << std::endl;
    return exit_success;
}
