/*
 * Chargeur natif du complément « Tableaux Word vers Excel ».
 *
 * Pourquoi : Windows n'active pas un composant COM .NET (serveur « mscoree.dll ») enregistré
 * pour le seul utilisateur courant (HKCU). Ce chargeur, une DLL native enregistrée dans HKCU
 * (sans droits administrateur), démarre le CLR .NET 4 et crée l'objet du complément
 * (WordTableToExcel.AddIn.Connect) contenu dans WordTableToExcel.dll, située dans le même dossier.
 *
 * Fonctionnement : ICLRRuntimeHost::ExecuteInDefaultAppDomain appelle la méthode statique
 * WordTableToExcel.AddIn.ShimEntryPoint.CreateAddIn(string), à laquelle on passe l'adresse
 * (hexadécimale) d'une variable IUnknown* ; la méthode y écrit le pointeur IUnknown de l'objet.
 *
 * Compilé en 32 et 64 bits avec MinGW-w64 (voir build/build-installation.sh), sans dépendance
 * autre que les DLL système.
 */
#define WIN32_LEAN_AND_MEAN
#define COBJMACROS
#include <windows.h>
#include <objbase.h>
#include <unknwn.h>

/* --- Identifiants ------------------------------------------------------------------------- */

static const CLSID CLSID_WordTableToExcel =
    { 0x03F63233, 0xF2FE, 0x4A75, { 0xAF, 0x7A, 0xF9, 0x9C, 0xBE, 0xDC, 0x80, 0x30 } };
static const CLSID CLSID_CLRMetaHost_ =
    { 0x9280188D, 0x0E8E, 0x4867, { 0xB3, 0x0C, 0x7F, 0xA8, 0x38, 0x84, 0xE8, 0xDE } };
static const IID IID_ICLRMetaHost_ =
    { 0xD332DB9E, 0xB9B3, 0x4125, { 0x82, 0x07, 0xA1, 0x48, 0x84, 0xF5, 0x32, 0x16 } };
static const IID IID_ICLRRuntimeInfo_ =
    { 0xBD39D1D2, 0xBA2F, 0x486A, { 0x89, 0xB0, 0xB4, 0xB0, 0xCB, 0x46, 0x68, 0x91 } };
static const CLSID CLSID_CLRRuntimeHost_ =
    { 0x90F1A06E, 0x7712, 0x4762, { 0x86, 0xB5, 0x7A, 0x5E, 0xBA, 0x6B, 0xDB, 0x02 } };
static const IID IID_ICLRRuntimeHost_ =
    { 0x90F1A06C, 0x7712, 0x4762, { 0x86, 0xB5, 0x7A, 0x5E, 0xBA, 0x6B, 0xDB, 0x02 } };
static const IID IID_IUnknown_ =
    { 0x00000000, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };
static const IID IID_IClassFactory_ =
    { 0x00000001, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };

/* --- Interfaces d'hébergement du CLR (metahost.h / mscoree.h), déclarées ici --------------- */

typedef struct HostMetaHost HostMetaHost;
typedef struct HostMetaHostVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(HostMetaHost *self, REFIID riid, void **ppv);
    ULONG (STDMETHODCALLTYPE *AddRef)(HostMetaHost *self);
    ULONG (STDMETHODCALLTYPE *Release)(HostMetaHost *self);
    HRESULT (STDMETHODCALLTYPE *GetRuntime)(HostMetaHost *self, LPCWSTR version, REFIID riid, void **runtime);
    /* méthodes suivantes non utilisées */
} HostMetaHostVtbl;
struct HostMetaHost { const HostMetaHostVtbl *lpVtbl; };

typedef struct HostRuntimeInfo HostRuntimeInfo;
typedef struct HostRuntimeInfoVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(HostRuntimeInfo *self, REFIID riid, void **ppv);
    ULONG (STDMETHODCALLTYPE *AddRef)(HostRuntimeInfo *self);
    ULONG (STDMETHODCALLTYPE *Release)(HostRuntimeInfo *self);
    void *GetVersionString;
    void *GetRuntimeDirectory;
    void *IsLoaded;
    void *LoadErrorString;
    void *LoadLibrary_;
    void *GetProcAddress_;
    HRESULT (STDMETHODCALLTYPE *GetInterface)(HostRuntimeInfo *self, REFCLSID clsid, REFIID riid, void **ppv);
    /* méthodes suivantes non utilisées */
} HostRuntimeInfoVtbl;
struct HostRuntimeInfo { const HostRuntimeInfoVtbl *lpVtbl; };

typedef struct HostRuntimeHost HostRuntimeHost;
typedef struct HostRuntimeHostVtbl {
    HRESULT (STDMETHODCALLTYPE *QueryInterface)(HostRuntimeHost *self, REFIID riid, void **ppv);
    ULONG (STDMETHODCALLTYPE *AddRef)(HostRuntimeHost *self);
    ULONG (STDMETHODCALLTYPE *Release)(HostRuntimeHost *self);
    HRESULT (STDMETHODCALLTYPE *Start)(HostRuntimeHost *self);
    void *Stop;
    void *SetHostControl;
    void *GetCLRControl;
    void *UnloadAppDomain;
    void *ExecuteInAppDomain;
    void *GetCurrentAppDomainId;
    void *ExecuteApplication;
    HRESULT (STDMETHODCALLTYPE *ExecuteInDefaultAppDomain)(HostRuntimeHost *self, LPCWSTR assemblyPath,
        LPCWSTR typeName, LPCWSTR methodName, LPCWSTR argument, DWORD *returnValue);
} HostRuntimeHostVtbl;
struct HostRuntimeHost { const HostRuntimeHostVtbl *lpVtbl; };

typedef HRESULT (WINAPI *CLRCreateInstanceFn)(REFCLSID clsid, REFIID riid, void **ppv);

/* --- Journal (même fichier que le complément) --------------------------------------------- */

static HMODULE g_module;

static void AppendLog(const char *step, HRESULT hr)
{
    WCHAR path[MAX_PATH + 64];
    DWORD len = GetEnvironmentVariableW(L"LOCALAPPDATA", path, MAX_PATH);
    if (len == 0 || len >= MAX_PATH) return;
    lstrcatW(path, L"\\WordTableToExcel");
    CreateDirectoryW(path, NULL);
    lstrcatW(path, L"\\WordTableToExcel.log");

    HANDLE file = CreateFileW(path, FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_ALWAYS,
                              FILE_ATTRIBUTE_NORMAL, NULL);
    if (file == INVALID_HANDLE_VALUE) return;

    SYSTEMTIME t;
    GetLocalTime(&t);
    char line[256];
    int n = wsprintfA(line, "%04d-%02d-%02d %02d:%02d:%02d ERROR Chargeur natif (%d bits) : %s, HRESULT 0x%08lX\r\n",
                      t.wYear, t.wMonth, t.wDay, t.wHour, t.wMinute, t.wSecond, (int)(sizeof(void *) * 8),
                      step, (unsigned long)hr);
    DWORD written;
    WriteFile(file, line, (DWORD)n, &written, NULL);
    CloseHandle(file);
}

/* --- Création de l'objet .NET ------------------------------------------------------------- */

static void ToHex(ULONG_PTR value, WCHAR *buffer)
{
    static const WCHAR digits[] = L"0123456789ABCDEF";
    WCHAR tmp[2 * sizeof(ULONG_PTR) + 1];
    int i = 0;
    do {
        tmp[i++] = digits[value & 0xF];
        value >>= 4;
    } while (value != 0);
    int j = 0;
    while (i > 0) buffer[j++] = tmp[--i];
    buffer[j] = 0;
}

static HRESULT CreateManagedAddIn(IUnknown **result)
{
    HRESULT hr;
    HostMetaHost *metaHost = NULL;
    HostRuntimeInfo *runtimeInfo = NULL;
    HostRuntimeHost *runtimeHost = NULL;
    WCHAR assemblyPath[MAX_PATH * 2];
    WCHAR argument[2 * sizeof(ULONG_PTR) + 1];
    IUnknown *created = NULL;
    DWORD returnValue = 0;

    *result = NULL;

    /* WordTableToExcel.dll se trouve dans le dossier de ce chargeur. */
    DWORD len = GetModuleFileNameW(g_module, assemblyPath, MAX_PATH);
    if (len == 0 || len >= MAX_PATH) { AppendLog("chemin du chargeur", E_FAIL); return E_FAIL; }
    WCHAR *slash = assemblyPath + len;
    while (slash > assemblyPath && *(slash - 1) != L'\\' && *(slash - 1) != L'/') slash--;
    *slash = 0;
    lstrcatW(assemblyPath, L"WordTableToExcel.dll");
    if (GetFileAttributesW(assemblyPath) == INVALID_FILE_ATTRIBUTES) {
        AppendLog("WordTableToExcel.dll introuvable", HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND));
        return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
    }

    HMODULE mscoree = LoadLibraryW(L"mscoree.dll");
    CLRCreateInstanceFn createInstance = mscoree ? (CLRCreateInstanceFn)(void *)GetProcAddress(mscoree, "CLRCreateInstance") : NULL;
    if (!createInstance) { AppendLog(".NET Framework 4 absent (CLRCreateInstance)", E_NOINTERFACE); return E_NOINTERFACE; }

    hr = createInstance(&CLSID_CLRMetaHost_, &IID_ICLRMetaHost_, (void **)&metaHost);
    if (FAILED(hr)) { AppendLog("CLRCreateInstance", hr); goto done; }

    hr = metaHost->lpVtbl->GetRuntime(metaHost, L"v4.0.30319", &IID_ICLRRuntimeInfo_, (void **)&runtimeInfo);
    if (FAILED(hr)) { AppendLog(".NET Framework 4 introuvable (GetRuntime)", hr); goto done; }

    hr = runtimeInfo->lpVtbl->GetInterface(runtimeInfo, &CLSID_CLRRuntimeHost_, &IID_ICLRRuntimeHost_, (void **)&runtimeHost);
    if (FAILED(hr)) { AppendLog("GetInterface(ICLRRuntimeHost)", hr); goto done; }

    hr = runtimeHost->lpVtbl->Start(runtimeHost);   /* S_FALSE si le CLR est déjà démarré */
    if (FAILED(hr)) { AppendLog("démarrage du CLR", hr); goto done; }

    ToHex((ULONG_PTR)&created, argument);
    hr = runtimeHost->lpVtbl->ExecuteInDefaultAppDomain(runtimeHost, assemblyPath,
        L"WordTableToExcel.AddIn.ShimEntryPoint", L"CreateAddIn", argument, &returnValue);
    if (FAILED(hr)) { AppendLog("chargement de WordTableToExcel.dll (ExecuteInDefaultAppDomain)", hr); goto done; }
    if (returnValue != 0 || created == NULL) {
        hr = returnValue != 0 ? (HRESULT)returnValue : E_FAIL;
        AppendLog("création du complément", hr);
        goto done;
    }

    *result = created;
    created = NULL;
    hr = S_OK;

done:
    if (created) created->lpVtbl->Release(created);
    if (runtimeHost) runtimeHost->lpVtbl->Release(runtimeHost);
    if (runtimeInfo) runtimeInfo->lpVtbl->Release(runtimeInfo);
    if (metaHost) metaHost->lpVtbl->Release(metaHost);
    return hr;
}

/* --- Fabrique de classe COM --------------------------------------------------------------- */

static HRESULT STDMETHODCALLTYPE Factory_QueryInterface(IClassFactory *self, REFIID riid, void **ppv)
{
    if (!ppv) return E_POINTER;
    if (IsEqualIID(riid, &IID_IUnknown_) || IsEqualIID(riid, &IID_IClassFactory_)) {
        *ppv = self;
        return S_OK;
    }
    *ppv = NULL;
    return E_NOINTERFACE;
}

static ULONG STDMETHODCALLTYPE Factory_AddRef(IClassFactory *self) { (void)self; return 2; }
static ULONG STDMETHODCALLTYPE Factory_Release(IClassFactory *self) { (void)self; return 1; }

static HRESULT STDMETHODCALLTYPE Factory_CreateInstance(IClassFactory *self, IUnknown *outer, REFIID riid, void **ppv)
{
    (void)self;
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (outer) return CLASS_E_NOAGGREGATION;

    IUnknown *addIn = NULL;
    HRESULT hr = CreateManagedAddIn(&addIn);
    if (FAILED(hr)) return hr;
    hr = addIn->lpVtbl->QueryInterface(addIn, riid, ppv);
    addIn->lpVtbl->Release(addIn);
    return hr;
}

static HRESULT STDMETHODCALLTYPE Factory_LockServer(IClassFactory *self, BOOL lock)
{
    (void)self;
    (void)lock;
    return S_OK;
}

static IClassFactoryVtbl g_factoryVtbl = {
    Factory_QueryInterface, Factory_AddRef, Factory_Release, Factory_CreateInstance, Factory_LockServer
};
static IClassFactory g_factory = { &g_factoryVtbl };

/* --- Exports ------------------------------------------------------------------------------ */

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv)
{
    if (!ppv) return E_POINTER;
    *ppv = NULL;
    if (!IsEqualCLSID(rclsid, &CLSID_WordTableToExcel)) return CLASS_E_CLASSNOTAVAILABLE;
    return Factory_QueryInterface(&g_factory, riid, ppv);
}

/* Le CLR ne peut pas être déchargé : la DLL reste chargée jusqu'à la fermeture de Word. */
STDAPI DllCanUnloadNow(void)
{
    return S_FALSE;
}

BOOL WINAPI DllMain(HINSTANCE instance, DWORD reason, LPVOID reserved)
{
    (void)reserved;
    if (reason == DLL_PROCESS_ATTACH) {
        g_module = instance;
        DisableThreadLibraryCalls(instance);
    }
    return TRUE;
}
