/* --------------------------------------------------------------------------------
 #
 #	4DPlugin.cpp
 #	source generated by 4D Plugin Wizard
 #	Project : Message file drop
 #	author : miyako
 #	2018/06/21
 #
 # --------------------------------------------------------------------------------*/

#include "4DPluginAPI.h"
#include "4DPlugin.h"

std::mutex globalMutex; /* WATCH_PATHS,CALLBACK_EVENT_PATHS */
std::mutex globalMutex1;/* for METHOD_PROCESS_ID */
std::mutex globalMutex2;/* for LISTENER_METHOD,WATCH_CONTEXT */
std::mutex globalMutex3;/* PROCESS_SHOULD_TERMINATE */
std::mutex globalMutex4;/* PROCESS_SHOULD_RESUME */

namespace MFD
{
    //constants
    const process_stack_size_t MONITOR_PROCESS_STACK_SIZE = 0;
	const process_name_t MONITOR_PROCESS_NAME = (PA_Unichar *)L"$MFD";
    
	C_TEXT WATCH_CONTEXT;

    //context management
	std::vector<CUTF16String> CALLBACK_MSG;
	std::vector<CUTF16String> CALLBACK_MHT;
	std::vector<CUTF16String> CALLBACK_HTMLBODY;

     //callback management
    C_TEXT LISTENER_METHOD;
    process_number_t METHOD_PROCESS_ID = 0;
    bool MONITOR_PROCESS_SHOULD_TERMINATE = false;
	bool PROCESS_SHOULD_RESUME = false;
}

#if VERSIONWIN
void generateUuid(std::wstring &uuidstr)
{
	RPC_WSTR str;
	UUID uuid;

	if (UuidCreate(&uuid) == RPC_S_OK) {
		if (UuidToString(&uuid, &str) == RPC_S_OK) {
			size_t len = wcslen((const wchar_t *)str);

			std::vector<wchar_t>buf(len);
			memcpy(&buf[0], str, len * sizeof(wchar_t));
			_wcsupr((wchar_t *)&buf[0]);

			uuidstr = std::wstring((const wchar_t *)&buf[0], len);

			RpcStringFree(&str);
		}
	}
}
#endif

#pragma mark -

void PluginMain(PA_long32 selector, PA_PluginParameters params)
{
	try
	{
		PA_long32 pProcNum = selector;
		sLONG_PTR *pResult = (sLONG_PTR *)params->fResult;
		PackagePtr pParams = (PackagePtr)params->fParameters;

		CommandDispatcher(pProcNum, pResult, pParams); 
	}
	catch(...)
	{

	}
}

void CommandDispatcher (PA_long32 pProcNum, sLONG_PTR *pResult, PackagePtr pParams)
{
	switch(pProcNum)
	{
// --- Message file drop
		case kInitPlugin :
		case kServerInitPlugin :
			OnStartup();
			break;
		/* this is too late to break listenerloop */	
        //case kDeinitPlugin :
        //OnExit();
        //break;
        
		case kCloseProcess:
			OnCloseProcess();
			break;
			
		case 1 :
			ACCEPT_MESSAGE_FILES(pResult, pParams);
			break;

	}
}

#pragma mark -

bool IsProcessOnExit()
{
	C_TEXT name;
	PA_long32 state, time;
	PA_GetProcessInfo(PA_GetCurrentProcessNumber(), name, &state, &time);
	CUTF16String procName(name.getUTF16StringPtr());
	CUTF16String exitProcName((PA_Unichar *)L"$xx");
	return (!procName.compare(exitProcName));
}

void OnStartup()
{
	
}

void OnExit()
{
    listenerLoopFinish();
}

void OnCloseProcess()
{
	if(IsProcessOnExit())
	{
		OnExit();
	}
}

#pragma mark -

#if VERSIONWIN
IDispatch *get_outlook_application() {

    IDispatch *pDispatch = NULL;

    CLSID CLSID_outlook_application;
    HRESULT hr = CLSIDFromProgID(L"Outlook.Application", &CLSID_outlook_application);
    if (SUCCEEDED(hr)) {
        IUnknown *pUnknown;
        hr = GetActiveObject(CLSID_outlook_application, NULL, (IUnknown **)&pUnknown);
        if (FAILED(hr)) {
            hr = CoCreateInstance(CLSID_outlook_application, NULL, CLSCTX_LOCAL_SERVER, IID_PPV_ARGS(&pDispatch));
            if (FAILED(hr)) {
                pDispatch = NULL;
            }
        }
        else {
            hr = pUnknown->QueryInterface(IID_PPV_ARGS(&pDispatch));
            if (FAILED(hr)) {
                pDispatch = NULL;
            }
        }
    }

    return pDispatch;
}

bool invoke_DISPATCH_METHOD_VT_DISPATCH(DISPID dispid, IDispatch *pDispatch, IDispatch **returnValue) {

    bool        success = false;
    HRESULT        hr;

    /*
    *    return value
    */

    VARIANT        varResult;
    VariantInit(&varResult);
    varResult.vt = VT_DISPATCH;

    /*
    *    arguments
    */

    size_t cArgs = 0;
    std::vector<VARIANT>args(cArgs);

    DISPPARAMS    dispParams;
    dispParams.cArgs = cArgs;
    dispParams.rgvarg = cArgs ? &args[0] : NULL;
    dispParams.cNamedArgs = 0;
    dispParams.rgdispidNamedArgs = NULL;

    /*
    *    error
    */

    EXCEPINFO    excepInfo = { 0 };

    hr = pDispatch->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &varResult, &excepInfo, NULL);

    if ((hr == S_OK) && (varResult.vt == VT_DISPATCH)) {
        *returnValue = varResult.pdispVal;
        success = true;
    }

    return success;
}

bool invoke_DISPATCH_METHOD_VT_I4_VT_DISPATCH(DISPID dispid, IDispatch *pDispatch, long arg1, IDispatch **returnValue) {

    bool        success = false;
    HRESULT        hr;

    /*
    *    return value
    */

    VARIANT        varResult;
    VariantInit(&varResult);
    varResult.vt = VT_DISPATCH;

    /*
    *    arguments
    */

    size_t cArgs = 1;
    std::vector<VARIANT>args(cArgs);

    VariantInit(&args[0]);
    args[0].vt = VT_I4;
    args[0].lVal = arg1;

    DISPPARAMS    dispParams;
    dispParams.cArgs = cArgs;
    dispParams.rgvarg = cArgs ? &args[0] : NULL;
    dispParams.cNamedArgs = 0;
    dispParams.rgdispidNamedArgs = NULL;

    /*
    *    error
    */

    EXCEPINFO    excepInfo = { 0 };

    hr = pDispatch->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &varResult, &excepInfo, NULL);

    if ((hr == S_OK) && (varResult.vt == VT_DISPATCH)) {
        *returnValue = varResult.pdispVal;
        success = true;
    }

    return success;
}

bool invoke_DISPATCH_PROPERTYGET_VT_I4(DISPID dispid, IDispatch *pDispatch, long *returnValue) {

    bool        success = false;
    HRESULT        hr;

    /*
    *    return value
    */

    VARIANT        varResult;
    VariantInit(&varResult);
    varResult.vt = VT_I4;

    /*
    *    arguments
    */

    size_t cArgs = 0;
    std::vector<VARIANT>args(cArgs);

    DISPPARAMS    dispParams;
    dispParams.cArgs = cArgs;
    dispParams.rgvarg = cArgs ? &args[0] : NULL;
    dispParams.cNamedArgs = 0;
    dispParams.rgdispidNamedArgs = NULL;

    /*
    *    error
    */

    EXCEPINFO    excepInfo = { 0 };

    hr = pDispatch->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &varResult, &excepInfo, NULL);

    if ((hr == S_OK) && (varResult.vt == VT_I4)) {
        *returnValue = varResult.lVal;
        success = true;
    }

    return success;
}

bool invoke_DISPATCH_METHOD_VT_I4_VT_BSTR_VT_EMPTY(DISPID dispid, IDispatch *pDispatch, long arg1, const wchar_t *arg2) {

    bool        success = false;
    HRESULT        hr;

    /*
    *    return value
    */

    VARIANT        varResult;
    VariantInit(&varResult);
    varResult.vt = VT_EMPTY;

    /*
    *    arguments
    */

    size_t cArgs = 2;
    std::vector<VARIANT>args(cArgs);

    VariantInit(&args[1]);
    args[1].vt = VT_BSTR;
    BSTR str1 = SysAllocString((const OLECHAR *)arg2);
    args[1].bstrVal = str1;

    VariantInit(&args[0]);
    args[0].vt = VT_UINT;
    args[0].uintVal = arg1;

    DISPPARAMS    dispParams;
    dispParams.cArgs = cArgs;
    dispParams.rgvarg = cArgs ? &args[0] : NULL;
    dispParams.cNamedArgs = 0;
    dispParams.rgdispidNamedArgs = NULL;

    /*
    *    error
    */

    EXCEPINFO    excepInfo = { 0 };

    hr = pDispatch->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispParams, &varResult, &excepInfo, NULL);

    if (hr == S_OK) {
        success = true;
    }

    SysFreeString(str1);

    return success;
}

bool invoke_DISPATCH_PROPERTYGET_VT_BSTR(DISPID dispid, IDispatch *pDispatch, CUTF16String& returnValue) {

    bool        success = false;
    HRESULT        hr;

    /*
    *    return value
    */

    VARIANT        varResult;
    VariantInit(&varResult);
    varResult.vt = VT_BSTR;

    /*
    *    arguments
    */

    size_t cArgs = 0;
    std::vector<VARIANT>args(cArgs);

    DISPPARAMS    dispParams;
    dispParams.cArgs = cArgs;
    dispParams.rgvarg = cArgs ? &args[0] : NULL;
    dispParams.cNamedArgs = 0;
    dispParams.rgdispidNamedArgs = NULL;

    /*
    *    error
    */

    EXCEPINFO    excepInfo = { 0 };

    hr = pDispatch->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &varResult, &excepInfo, NULL);

    if ((hr == S_OK) && (varResult.vt == VT_BSTR)) {
        if (varResult.bstrVal) {
            returnValue = (const PA_Unichar *)varResult.bstrVal;
            SysFreeString(varResult.bstrVal);
        }
        else {
            returnValue = (const PA_Unichar *)L"";

        }
        success = true;
    }

    return success;
}

bool invoke_DISPATCH_PROPERTYGET_VT_DISPATCH(DISPID dispid, IDispatch *pDispatch, IDispatch **returnValue) {

	bool		success = false;
	HRESULT		hr;

	/*
	*	return value
	*/

	VARIANT		varResult;
	VariantInit(&varResult);
	varResult.vt = VT_DISPATCH;

	/*
	*	arguments
	*/

	size_t cArgs = 0;
	std::vector<VARIANT>args(cArgs);

	DISPPARAMS	dispParams;
	dispParams.cArgs = cArgs;
	dispParams.rgvarg = cArgs ? &args[0] : NULL;
	dispParams.cNamedArgs = 0;
	dispParams.rgdispidNamedArgs = NULL;

	/*
	*	error
	*/

	EXCEPINFO	excepInfo = { 0 };

	hr = pDispatch->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &dispParams, &varResult, &excepInfo, NULL);

	if ((hr == S_OK) && (varResult.vt == VT_DISPATCH)) {
		*returnValue = varResult.pdispVal;
		success = true;
	}

	return success;
}

namespace CApplication {

    bool ActiveExplorer(IDispatch *pDispatch, IDispatch **pExplorer) {
        
        return invoke_DISPATCH_METHOD_VT_DISPATCH(0x111, pDispatch, pExplorer);
    }

}

namespace CExplorer {

    bool get_Selection(IDispatch *pDispatch, IDispatch **pSelection) {
        
        return invoke_DISPATCH_PROPERTYGET_VT_DISPATCH(0x2202, pDispatch, pSelection);
    }

}

namespace CItems {

    bool get_Count(IDispatch *pDispatch, long *count) {
        return invoke_DISPATCH_PROPERTYGET_VT_I4(0x50, pDispatch, count);
    }

    bool Item(IDispatch *pDispatch, long index, IDispatch **pItem) {
        return invoke_DISPATCH_METHOD_VT_I4_VT_DISPATCH(0x51, pDispatch, index, pItem);
    }

}

namespace CMailItem {

    bool get_Class(IDispatch *pDispatch, long *olObjClass) {
        return invoke_DISPATCH_PROPERTYGET_VT_I4(0xf00a, pDispatch, olObjClass);
    }

    bool get_HTMLBody(IDispatch *pDispatch, CUTF16String& htmlBody) {

        return invoke_DISPATCH_PROPERTYGET_VT_BSTR(0xf404, pDispatch, htmlBody);
    }

    bool SaveAs(IDispatch *pDispatch, const wchar_t *path, long type) {

        return invoke_DISPATCH_METHOD_VT_I4_VT_BSTR_VT_EMPTY(0xf051, pDispatch, type, path);
    }
}

void getTempFolder(std::wstring& path) {
	wchar_t	fDrive[_MAX_DRIVE], fDir[_MAX_DIR], fName[_MAX_FNAME], fExt[_MAX_EXT];
	std::vector<wchar_t>_buf(_MAX_DRIVE + _MAX_DIR + _MAX_FNAME + _MAX_EXT);
	GetTempPath(_buf.size(), (LPWSTR)&_buf[0]);
	path = (LPWSTR)&_buf[0];

	std::wstring uuid;
	generateUuid(uuid);

	path += uuid;
	path += L"\\";

	SHCreateDirectory(NULL, (PCWSTR)path.c_str());
}

void getEnumFile(const wchar_t *folder, long index, const wchar_t *extension, CUTF16String& path) {
#define LENGTH_OF_LONG 21
	wchar_t digits[LENGTH_OF_LONG];
	ZeroMemory(digits, LENGTH_OF_LONG * sizeof(wchar_t));
	_itow(index, digits, LENGTH_OF_LONG);
	path  = (const PA_Unichar *)folder;
	path += (const PA_Unichar *)digits;
	path += (const PA_Unichar *)extension;
}

HRESULT outlook_export_selected_messages(){
	bool didReceiveItem = false;
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (SUCCEEDED(hr)) {
        IDispatch *pDispatch = get_outlook_application();
        if (pDispatch) {
            
            IDispatch *pExplorer = NULL;
            if (CApplication::ActiveExplorer(pDispatch, &pExplorer) && (pExplorer != NULL)){
                IDispatch *pSelection = NULL;
                if (CExplorer::get_Selection(pExplorer, &pSelection) && (pSelection != NULL)){
                    long count = 0;
                    if (CItems::get_Count(pSelection, &count)) {
                        long index = 0;

						std::wstring tempFolder;
						getTempFolder(tempFolder);

                        while (index < count) {
                            index++;
                            IDispatch *pItem = NULL;
                            if (CItems::Item(pSelection, index, &pItem) && (pItem != NULL)) {
                                long olObjClass = 0;
                                if (CMailItem::get_Class(pItem, &olObjClass)) {
									if (olObjClass == 43 /* olMail */) {

										std::lock_guard<std::mutex> lock(globalMutex);

										CUTF16String htmlBody;
										CMailItem::get_HTMLBody(pItem, htmlBody);
										MFD::CALLBACK_HTMLBODY.push_back(htmlBody);

										CUTF16String mhtPath, msgPath;

										getEnumFile(tempFolder.c_str(), index + 1, L".mht", mhtPath);
										getEnumFile(tempFolder.c_str(), index + 1, L".msg", msgPath);

										CMailItem::SaveAs(pItem, (const wchar_t *)mhtPath.c_str(), 10 /* olMHTML */);
										CMailItem::SaveAs(pItem, (const wchar_t *)msgPath.c_str(),  9 /* olMSGUnicode */);

										MFD::CALLBACK_MHT.push_back(mhtPath);
										MFD::CALLBACK_MSG.push_back(msgPath);

										didReceiveItem = true;
									}
                                }
                            }
                        }
                    }
                }
            }

            pDispatch->Release();
        }
        CoUninitialize();
    }
	return didReceiveItem;
}
#endif

#if VERSIONWIN
unsigned __stdcall doIt(void *p)
{
	Params *params = (Params *)p;
	
	HANDLE h = params->h;
	
	HANDLE b = CreateEvent(NULL,
												 TRUE,
												 FALSE,
												 (const wchar_t *)params->b);
	
	if (b)
	{
		DWORD notifyFilter =
			FILE_NOTIFY_CHANGE_SIZE
			| FILE_NOTIFY_CHANGE_CREATION
			| FILE_NOTIFY_CHANGE_LAST_WRITE
			| FILE_NOTIFY_CHANGE_FILE_NAME 
			/* | FILE_NOTIFY_CHANGE_DIR_NAME */
			/* | FILE_NOTIFY_CHANGE_SECURITY */
			/* | FILE_NOTIFY_CHANGE_ATTRIBUTES */
			| FILE_NOTIFY_CHANGE_LAST_ACCESS;
		
		std::vector<wchar_t>buf(BUF_SIZE);
		DWORD bytesReturned;
		BOOL exit = FALSE;
		
		do {
			if (ReadDirectoryChangesW(h,
				(FILE_NOTIFY_INFORMATION *)&buf[0],
				BUF_SIZE,
				TRUE,
				notifyFilter,
				&bytesReturned,
				NULL,
				NULL))
			{
				if (bytesReturned)
				{
					DWORD data_len = bytesReturned;
					DWORD len = sizeof(data_len) + data_len;
					HANDLE fmOut = CreateFileMapping(
						INVALID_HANDLE_VALUE,
						NULL,
						PAGE_READWRITE,
						0, len,
						L"PARAM_OUT");
					if (fmOut)
					{
						LPVOID bufOut = MapViewOfFile(fmOut, FILE_MAP_WRITE, 0, 0, len);
						if (bufOut)
						{
							unsigned char *p = (unsigned char *)bufOut;
							try
							{
								CopyMemory(p, &data_len, sizeof(data_len));
								p += sizeof(data_len);
								if (data_len)
								{
									CopyMemory(p, (void *)&buf[0], data_len);
								}
							}
							catch (...)
							{

							}
							UnmapViewOfFile(bufOut);
						}//bufOut

						HANDLE a = OpenEvent(EVENT_ALL_ACCESS, FALSE, (const wchar_t *)params->a);
						
						if (a)
						{
							SetEvent(a);
							CloseHandle(a);
						}
						
						WaitForSingleObject(b, INFINITE);
						ResetEvent(b);
						
						CloseHandle(fmOut);
					}
					
				}
			}
			
			if (MFD::MONITOR_PROCESS_SHOULD_TERMINATE || bytesReturned == 0) exit = TRUE;
			
		} while (!exit);
		
		CloseHandle(b);
	}
	_endthreadex(0);
	return 0;
}
#endif

void listenerLoop()
{
#if VERSIONWIN
	std::vector<HANDLE>::iterator it;
	std::vector<HANDLE>watch_path_handles;
#endif

	std::vector<CUTF16String>watch_path_handle_paths;

	if (1)
	{
#if VERSIONWIN
        if(1)
        {
            std::lock_guard<std::mutex> lock(globalMutex3);
            
            MFD::MONITOR_PROCESS_SHOULD_TERMINATE = false;
        }

		/* Current process returns 0 for PA_NewProcess */
		PA_long32 currentProcessNumber = PA_GetCurrentProcessNumber();

		while (!PA_IsProcessDying())
		{
			PA_YieldAbsolute();

			bool PROCESS_SHOULD_RESUME;
			bool PROCESS_SHOULD_TERMINATE;

			if (1)
			{
				PROCESS_SHOULD_RESUME = MFD::PROCESS_SHOULD_RESUME;
				PROCESS_SHOULD_TERMINATE = MFD::MONITOR_PROCESS_SHOULD_TERMINATE;
			}

			if (PROCESS_SHOULD_RESUME)
			{
				size_t TYPES;

				if (1)
				{
					std::lock_guard<std::mutex> lock(globalMutex);

					TYPES = MFD::CALLBACK_HTMLBODY.size();
				}

				while (TYPES != 0)
				{
					PA_YieldAbsolute();

					if (true)
					{
						std::wstring processName;
						generateUuid(processName);
						PA_NewProcess((void *)listenerLoopExecuteMethod,
							MFD::MONITOR_PROCESS_STACK_SIZE,
							(PA_Unichar *)processName.c_str());
					}
					else {
						listenerLoopExecuteMethod();
					}

					if (PROCESS_SHOULD_TERMINATE)
						break;

					if (1)
					{
						std::lock_guard<std::mutex> lock(globalMutex);

						TYPES = MFD::CALLBACK_HTMLBODY.size();
						PROCESS_SHOULD_TERMINATE = MFD::MONITOR_PROCESS_SHOULD_TERMINATE;
					}
				}

				if (TYPES == 0)
				{
					std::lock_guard<std::mutex> lock(globalMutex4);

					MFD::PROCESS_SHOULD_RESUME = false;
				}

		}
			else
			{
				PA_PutProcessToSleep(currentProcessNumber, CALLBACK_SLEEP_TIME);
			}

		}
#endif
	}
	
	if (1)
	{
		std::lock_guard<std::mutex> lock(globalMutex);

		MFD::CALLBACK_HTMLBODY.clear();
		MFD::CALLBACK_MHT.clear();
		MFD::CALLBACK_MSG.clear();
	}

    if(1)
    {
        std::lock_guard<std::mutex> lock(globalMutex1);
        
        MFD::METHOD_PROCESS_ID = 0;
    }
	
	PA_KillProcess();
}

void listenerLoopStart()
{
	if (!MFD::METHOD_PROCESS_ID)
	{
        std::lock_guard<std::mutex> lock(globalMutex1);
        
        MFD::METHOD_PROCESS_ID = PA_NewProcess((void *)listenerLoop,
                                                MFD::MONITOR_PROCESS_STACK_SIZE,
                                                MFD::MONITOR_PROCESS_NAME);
	}
	else
	{
		listenerLoopFinish();
	}

}

void listenerLoopFinishGracefully() {
	if (MFD::METHOD_PROCESS_ID)
	{
		if (1)
		{
			std::lock_guard<std::mutex> lock(globalMutex4);

			MFD::PROCESS_SHOULD_RESUME = true;
		}
		PA_YieldAbsolute();
	}
}

void listenerLoopFinish()
{
	if(MFD::METHOD_PROCESS_ID)
	{
        if(1)
        {
            std::lock_guard<std::mutex> lock(globalMutex3);
            
            MFD::MONITOR_PROCESS_SHOULD_TERMINATE = true;
        }

		PA_YieldAbsolute();
		
        if(1)
        {
            std::lock_guard<std::mutex> lock(globalMutex4);

            MFD::PROCESS_SHOULD_RESUME = true;
        }
	}
}

void listenerLoopExecute()
{
    if(1)
    {
        std::lock_guard<std::mutex> lock(globalMutex3);
        
        MFD::MONITOR_PROCESS_SHOULD_TERMINATE = false;
    }

    if(1)
    {
        std::lock_guard<std::mutex> lock(globalMutex4);
        
        MFD::PROCESS_SHOULD_RESUME = true;
    }
}

void listenerLoopExecuteMethod()
{
	CUTF16String htmlBody, mhtPath, msgPath;

    if(1)
    {
        std::lock_guard<std::mutex> lock(globalMutex);
        
		std::vector<CUTF16String>::iterator p = MFD::CALLBACK_HTMLBODY.begin();

		htmlBody = *p;
		MFD::CALLBACK_HTMLBODY.erase(p);

		p = MFD::CALLBACK_MHT.begin();

		mhtPath = *p;
		MFD::CALLBACK_MHT.erase(p);

		p = MFD::CALLBACK_MSG.begin();

		msgPath = *p;
		MFD::CALLBACK_MSG.erase(p);
    }

	method_id_t methodId = PA_GetMethodID((PA_Unichar *)MFD::LISTENER_METHOD.getUTF16StringPtr());
	
	if(methodId)
	{
		PA_Variable	params[4];
		params[0] = PA_CreateVariable(eVK_Unistring);
		params[1] = PA_CreateVariable(eVK_Unistring);
		params[2] = PA_CreateVariable(eVK_Unistring);
		params[3] = PA_CreateVariable(eVK_Unistring);

		PA_Unistring body = PA_CreateUnistring((PA_Unichar *)htmlBody.c_str());
		PA_Unistring mht = PA_CreateUnistring((PA_Unichar *)mhtPath.c_str());
		PA_Unistring msg = PA_CreateUnistring((PA_Unichar *)msgPath.c_str());
		PA_Unistring context = PA_CreateUnistring((PA_Unichar *)MFD::WATCH_CONTEXT.getUTF16StringPtr());

		PA_SetStringVariable(&params[0], &body);
		PA_SetStringVariable(&params[1], &mht);
		PA_SetStringVariable(&params[2], &msg);
		PA_SetStringVariable(&params[3], &context);

		PA_ExecuteMethodByID(methodId, params, 4);

		PA_ClearVariable(&params[0]);
		PA_ClearVariable(&params[1]);
		PA_ClearVariable(&params[2]);

	}else
	{
		PA_Variable	params[5];
		params[1] = PA_CreateVariable(eVK_Unistring);
		params[2] = PA_CreateVariable(eVK_Unistring);
		params[3] = PA_CreateVariable(eVK_Unistring);
		params[4] = PA_CreateVariable(eVK_Unistring);

		PA_Unistring body = PA_CreateUnistring((PA_Unichar *)htmlBody.c_str());
		PA_Unistring mht = PA_CreateUnistring((PA_Unichar *)mhtPath.c_str());
		PA_Unistring msg = PA_CreateUnistring((PA_Unichar *)msgPath.c_str());
		PA_Unistring context = PA_CreateUnistring((PA_Unichar *)MFD::WATCH_CONTEXT.getUTF16StringPtr());

		PA_SetStringVariable(&params[1], &body);
		PA_SetStringVariable(&params[2], &mht);
		PA_SetStringVariable(&params[3], &msg);
		PA_SetStringVariable(&params[4], &context);

		params[0] = PA_CreateVariable(eVK_Unistring);
		PA_Unistring method = PA_CreateUnistring((PA_Unichar *)MFD::LISTENER_METHOD.getUTF16StringPtr());
		PA_SetStringVariable(&params[0], &method);

		PA_ExecuteCommandByID(1007, params, 4);

		PA_ClearVariable(&params[0]);
		PA_ClearVariable(&params[1]);
		PA_ClearVariable(&params[2]);
		PA_ClearVariable(&params[3]);
		PA_ClearVariable(&params[4]);
	}
}

#pragma mark MyDropTarget

#if VERSIONWIN

class MyDropTarget : public IDropTarget
{
	
private:
	
	ULONG m_cRefCount; // Reference count
	HWND m_hWnd;
	UINT m_cfFormatFileDescriptor;
	UINT m_cfFormatRPSourceFolder;
	UINT m_cfFormatRPItem;
	UINT m_cfFormatRPMessages;
	UINT m_cfFormatRPLatestMessages;
	UINT m_cfFormatFileContents;
	UINT m_cfFormatWebCustomData;
		
	bool isFileDrop(IDataObject *pDataObj, FORMATETC *formatetc) {
			
		formatetc->cfFormat = CF_HDROP;
		formatetc->ptd = NULL;
		formatetc->dwAspect = DVASPECT_CONTENT;
		formatetc->lindex = -1;
		formatetc->tymed = TYMED_HGLOBAL;
		
		if (S_OK == pDataObj->QueryGetData(formatetc))
			return true;
		
		return false;
	}
	
	bool isFilePromiseDrop(IDataObject *pDataObj, FORMATETC *formatetc) {
				
		formatetc->cfFormat = m_cfFormatFileDescriptor;
		formatetc->ptd = NULL;
		formatetc->dwAspect = DVASPECT_CONTENT;
		formatetc->lindex = -1;
		formatetc->tymed = TYMED_HGLOBAL;
		
		if (S_OK == pDataObj->QueryGetData(formatetc))
			return true;
		
		return false;
	}
	
	bool isOutlookDrop(IDataObject *pDataObj) {
	
		FORMATETC formatetc;
	
		formatetc.ptd = NULL;
		formatetc.dwAspect = DVASPECT_CONTENT;
		formatetc.lindex = -1;
	
		/* look for outlook private formats */
		formatetc.tymed = TYMED_HGLOBAL;
		formatetc.cfFormat = m_cfFormatRPItem;
		
		if (S_OK == pDataObj->QueryGetData(&formatetc))
			return true;
		
		formatetc.tymed = TYMED_ISTREAM;
		formatetc.cfFormat = m_cfFormatRPMessages;
		
		if (S_OK == pDataObj->QueryGetData(&formatetc))
			return true;

		/* no need to check for these... */
		//formatetc.cfFormat = m_cfFormatRPSourceFolder;
		//formatetc.cfFormat = m_cfFormatRPLatestMessages;

		return false;
	}
	
	bool isNewOutlookDrop(IDataObject* pDataObj, FORMATETC* formatetc) {

		formatetc->cfFormat = CF_HDROP;
		formatetc->ptd = NULL;
		formatetc->dwAspect = DVASPECT_CONTENT;
		formatetc->lindex = -1;
		formatetc->tymed = TYMED_HGLOBAL;
		formatetc->cfFormat = m_cfFormatWebCustomData;

		if (S_OK == pDataObj->QueryGetData(formatetc))
			return true;

		return false;
	}


	public:
	
		MyDropTarget::MyDropTarget() : m_cRefCount(1), m_hWnd(NULL) {
			m_cfFormatFileDescriptor = RegisterClipboardFormat(CFSTR_FILEDESCRIPTOR);
			m_cfFormatRPSourceFolder = RegisterClipboardFormat(L"RenPrivateSourceFolder");
			m_cfFormatRPItem = RegisterClipboardFormat(L"RenPrivateItem");
			m_cfFormatRPMessages = RegisterClipboardFormat(L"RenPrivateMessages");
			m_cfFormatRPLatestMessages = RegisterClipboardFormat(L"RenPrivateLatestMessages");
			m_cfFormatFileContents = RegisterClipboardFormat(CFSTR_FILECONTENTS);
			m_cfFormatWebCustomData = RegisterClipboardFormat(L"Chromium Web Custom MIME Data Format");

		}
	
		MyDropTarget::~MyDropTarget() {}
	
		// IUnknown methods...
		virtual HRESULT _stdcall QueryInterface(REFIID riid, void **ppvObject) {
			*ppvObject = NULL;
			return E_NOTIMPL;
		}
	
		virtual ULONG _stdcall AddRef(void) {
			return InterlockedIncrement(&m_cRefCount);
		}
		
		virtual ULONG _stdcall Release(void) {
			if (!InterlockedDecrement(&m_cRefCount))
			{
				delete this;
				return 0;
			}
			return m_cRefCount;
		}
	
	// IDropTarget methods...
		virtual HRESULT __stdcall DragEnter(
			IDataObject *pDataObj,
			DWORD       grfKeyState,
			POINTL      pt,
			DWORD       *pdwEffect
		)
		{
		FORMATETC formatetc;
		if (this->isFileDrop(pDataObj, &formatetc))
		{
			//*pdwEffect = DROPEFFECT_COPY;
			return S_OK;
		}

		if (this->isFilePromiseDrop(pDataObj, &formatetc))
		{
			//*pdwEffect = DROPEFFECT_COPY;
			return S_OK;
		}

		if (this->isOutlookDrop(pDataObj))
		{
			//*pdwEffect = DROPEFFECT_COPY;
			return S_OK;
		}

		if (this->isNewOutlookDrop(pDataObj, &formatetc))
		{
			//*pdwEffect = DROPEFFECT_COPY;
			return S_OK;
		}
		
		*pdwEffect = DROPEFFECT_NONE;
		return DRAGDROP_S_CANCEL;
	}
	
	virtual HRESULT __stdcall DragLeave(){return S_OK;}
	
	virtual HRESULT __stdcall DragOver(
		DWORD  grfKeyState,
		POINTL pt,
		DWORD  *pdwEffect
	)
	{
		return S_OK;
	}
	
	virtual HRESULT __stdcall Drop(
		IDataObject *pDataObj,
		DWORD       grfKeyState,
		POINTL      pt,
		DWORD       *pdwEffect
	)
	{		
		bool didReceiveItem = false;

		std::wstring tempFolder;
		getTempFolder(tempFolder);

		FORMATETC formatetc;

		if (this->isFilePromiseDrop(pDataObj, &formatetc))
		{
			if (S_OK == pDataObj->QueryGetData(&formatetc))
			{
				STGMEDIUM stm = {};
				if (SUCCEEDED(pDataObj->GetData(&formatetc, &stm)))
				{
					LPFILEGROUPDESCRIPTOR pFileGroupDescriptor = static_cast<LPFILEGROUPDESCRIPTOR>(GlobalLock(stm.hGlobal));
					if (pFileGroupDescriptor != NULL)
					{
						UINT nFiles = pFileGroupDescriptor->cItems;

						if (nFiles != 0)
						{
							LPFILEDESCRIPTOR pFileDescriptor = pFileGroupDescriptor->fgd;

							WCHAR *fileName = pFileDescriptor->cFileName;
							std::wstring tempFile = tempFolder + fileName;

							FORMATETC formatetc0;

							formatetc0.cfFormat = m_cfFormatRPMessages;
							formatetc0.ptd = NULL;
							formatetc0.dwAspect = DVASPECT_CONTENT;
							formatetc0.lindex = -1;
							formatetc0.tymed = TYMED_ISTREAM;

							STGMEDIUM stm0 = {};

							if (S_OK != pDataObj->GetData(&formatetc0, &stm0))
							{
								for (UINT i = 0; i < nFiles; ++i)
								{
									WCHAR *fileName = pFileDescriptor->cFileName;
									std::wstring tempFile = tempFolder + fileName;

									FORMATETC formatetc2;

									formatetc2.cfFormat = m_cfFormatFileContents;
									formatetc2.ptd = NULL;
									formatetc2.dwAspect = DVASPECT_CONTENT;
									formatetc2.lindex = -1;// i;
									formatetc2.tymed = TYMED_ISTREAM;

									STGMEDIUM stm2 = {};

									if (SUCCEEDED(pDataObj->GetData(&formatetc2, &stm2)))
									{
										if (stm2.pstm != NULL)
										{
											size_t bufSize = 8192;
											ULONG bytesRead = 0;
											std::vector<unsigned char>buf(bufSize);
											FILE *f = _wfopen(tempFile.c_str(), L"wb");
											if (f != NULL)
											{
												while (1)
												{
													stm2.pstm->Read(&buf[0], bufSize, &bytesRead);
													fwrite(&buf[0], bytesRead, sizeof(unsigned char), f);
													if (bytesRead < bufSize) break;
												}
												fclose(f);

												std::lock_guard<std::mutex> lock(globalMutex);

												MFD::CALLBACK_HTMLBODY.push_back(CUTF16String((const PA_Unichar *)tempFile.c_str(), tempFile.length()));

												CUTF16String msgPath, mhtPath;
												MFD::CALLBACK_MHT.push_back(mhtPath);//empty
												MFD::CALLBACK_MSG.push_back(msgPath);//empty

												didReceiveItem = true;
											}
										}
									}

									pFileDescriptor++;
								}

							}
						}
					}
					GlobalUnlock(stm.hGlobal);
				}
			}
		}

		if ((!didReceiveItem) && (this->isOutlookDrop(pDataObj)))
		{
			didReceiveItem = outlook_export_selected_messages();
		}
		
		if ((!didReceiveItem) && (this->isFileDrop(pDataObj, &formatetc)))
		{
			if (S_OK == pDataObj->QueryGetData(&formatetc))
			{
				STGMEDIUM stm = {};
				if (SUCCEEDED(pDataObj->GetData(&formatetc, &stm)))
				{
					HDROP hDrop = static_cast<HDROP>(GlobalLock(stm.hGlobal));
					if (hDrop != NULL)
					{
						UINT nFiles = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
						if (nFiles != 0)
						{
							for (UINT i = 0; i < nFiles; ++i)
							{
								size_t len = DragQueryFile(hDrop, i, NULL, 0);
								len++;
								std::vector<unsigned char>buf((len) * sizeof(wchar_t));
								if (DragQueryFile(hDrop, i, (LPTSTR)&buf[0], len))
								{
									std::lock_guard<std::mutex> lock(globalMutex);

									MFD::CALLBACK_HTMLBODY.push_back(CUTF16String((const PA_Unichar *)&buf[0], len));

									CUTF16String msgPath, mhtPath;
									MFD::CALLBACK_MHT.push_back(mhtPath);//empty
									MFD::CALLBACK_MSG.push_back(msgPath);//empty

									didReceiveItem = true;
								}
							}
						}
					}
					GlobalUnlock(stm.hGlobal);
				}
			}
		}

		if ((!didReceiveItem) && (this->isNewOutlookDrop(pDataObj, &formatetc)))
		{
			if (S_OK == pDataObj->QueryGetData(&formatetc))
			{
				STGMEDIUM stm = {};
				if (SUCCEEDED(pDataObj->GetData(&formatetc, &stm)))
				{
                    uint32_t readBytes = 0;
                    uint32_t bufferSize = 0;
                    
					unsigned char *ptr = static_cast<unsigned char*>(GlobalLock(stm.hGlobal));

					if (ptr != NULL)
                    {
                        uint32_t size = 0;
                        bufferSize = sizeof(uint32_t);
						memcpy(&size, ptr, bufferSize);
						size += sizeof(uint32_t);
                        readBytes += bufferSize;
                        ptr += bufferSize;
                        
						uint32_t count = 0;//not used
						bufferSize = sizeof(uint32_t);
						memcpy(&count, ptr, bufferSize);//count of kvp; not used
						readBytes += bufferSize;
						ptr += bufferSize;
						bool isValue = false;
						CUTF16String key;

						do
						{
							uint32_t len = 0;
							bufferSize = sizeof(uint32_t);
							memcpy(&len, ptr, bufferSize);
							readBytes += bufferSize;
							ptr += bufferSize;

							bufferSize = sizeof(PA_Unichar) * len;

							if (readBytes + bufferSize > size) {
								//bust
							}
							else {
								std::vector<unsigned char>buf(bufferSize);
								memcpy(&buf[0], ptr, bufferSize);
								readBytes += bufferSize;
								ptr += bufferSize;
								if (readBytes % sizeof(uint32_t) != 0) {
									readBytes += sizeof(PA_Unichar);
									ptr += sizeof(PA_Unichar);
								}
								
								if (!isValue) {
									key = CUTF16String((const PA_Unichar*)& buf[0], len);
									isValue = true;
								}
								else {
									CUTF16String value((const PA_Unichar*)&buf[0], len);
									if ((key == CUTF16String((const PA_Unichar*)L"maillistrow"))
									|| (key == CUTF16String((const PA_Unichar*)L"multimaillistconversationrows"))) {
									
										if (1)
										{
											std::lock_guard<std::mutex> lock(globalMutex);

											MFD::CALLBACK_HTMLBODY.push_back(value);

											CUTF16String msgPath, mhtPath;
											MFD::CALLBACK_MHT.push_back(mhtPath);//empty
											MFD::CALLBACK_MSG.push_back(msgPath);//empty

											didReceiveItem = true;
										}

									}
									//multimaillistconversationrows
									isValue = false;
								}
							}

						} while (readBytes + sizeof(uint32_t) <= size);

					}
					GlobalUnlock(stm.hGlobal);
				}
			}

		}

		if (didReceiveItem)
		{
			std::lock_guard<std::mutex> lock(globalMutex4);

			MFD::PROCESS_SHOULD_RESUME = true;
		}

		return S_OK;
	}
	
	virtual HRESULT __stdcall Register(HWND hwnd)
	{
		m_hWnd = hwnd;
		/* need to revoke standard 4D implementation */
		RevokeDragDrop(hwnd);
		return RegisterDragDrop(hwnd, this);
	}
	
	virtual HRESULT __stdcall Unregister()
	{
		HWND hwnd = m_hWnd;
		m_hWnd = 0;
		return RevokeDragDrop(hwnd);
	}
	
};

#endif

#pragma mark OleInitClass

#if VERSIONWIN
class OleInitClass {
	
public:
	OleInitClass() {
		OleInitialize(NULL);
	}
	
	~OleInitClass() {
		OleUninitialize();
	}
};

OleInitClass g_OleInitClass;
MyDropTarget g_MyDropTarget;

#endif

// ------------------------------- Message file drop ------------------------------

void ACCEPT_MESSAGE_FILES(sLONG_PTR *pResult, PackagePtr pParams)
{
	C_LONGINT Param1_window;
	C_TEXT Param2_method;

	Param1_window.fromParamAtIndex(pParams, 1);
	
#if VERSIONWIN
	if(!IsProcessOnExit())
	{
		int arg = Param1_window.getIntValue();

		if (arg == 0)
		{
			if (1)
			{
				std::lock_guard<std::mutex> lock(globalMutex2);

				MFD::LISTENER_METHOD.fromParamAtIndex(pParams, 2);
				MFD::WATCH_CONTEXT.fromParamAtIndex(pParams, 3);
			}

			g_MyDropTarget.Unregister();

			listenerLoopFinish();
		}else

		if (arg == -1)
		{
			g_MyDropTarget.Unregister();

			listenerLoopFinishGracefully();
		}
		else {
			
			HWND hwnd = (HWND)PA_GetHWND(reinterpret_cast<PA_WindowRef>(arg));
			if (hwnd)
			{
				if (1)
				{
					std::lock_guard<std::mutex> lock(globalMutex2);

					MFD::LISTENER_METHOD.fromParamAtIndex(pParams, 2);
					MFD::WATCH_CONTEXT.fromParamAtIndex(pParams, 3);
				}
				g_MyDropTarget.Register(hwnd);

				listenerLoopStart();
			}
		}

	}/* IsProcessOnExit */
#endif
}
