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

std::mutex globalMutex;
std::mutex globalMutex0;

namespace MFD
{
	const process_name_t MONITOR_PROCESS_NAME = (PA_Unichar *)L"$MFD";
	process_number_t MONITOR_PROCESS_ID = 0;
	const process_stack_size_t MONITOR_PROCESS_STACK_SIZE = 0;
	bool MONITOR_PROCESS_SHOULD_TERMINATE = false;
	C_TEXT WATCH_METHOD;
	C_TEXT WATCH_CONTEXT;
	ARRAY_TEXT WATCH_PATHS;
	std::vector<CUTF16String>CALLBACK_EVENT_PATHS;
	double MONITOR_LATENCY;

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
			
		case kCloseProcess :
			OnCloseProcess();
			break;
			
		case 1 :
			ACCEPT_MESSAGE_FILES(pResult, pParams);
			break;

	}
}

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

void OnCloseProcess()
{
	if(IsProcessOnExit())
	{
		listenerLoopFinish();
	}
}

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
		std::lock_guard<std::mutex> lock(globalMutex);
		
#if VERSIONWIN
		MFD::MONITOR_PROCESS_SHOULD_TERMINATE = false;
		for (UINT i = 0; i < MFD::WATCH_PATHS.getSize(); ++i)
		{
			CUTF16String path;
			MFD::WATCH_PATHS.copyUTF16StringAtIndex(&path, i);
			HANDLE h = CreateFileW(
				(LPCWSTR)path.c_str(),
				FILE_LIST_DIRECTORY,
				FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
				NULL,
				OPEN_EXISTING,
				FILE_FLAG_BACKUP_SEMANTICS | FILE_FLAG_OVERLAPPED,
				NULL);
			if (h != INVALID_HANDLE_VALUE)
			{
				watch_path_handles.push_back(h);
				watch_path_handle_paths.push_back(path);
			}
		}
#endif
	}
	
#if VERSIONWIN
	std::vector<HANDLE>handles;
	std::vector<HANDLE>signals;
#endif
	
	std::vector<CUTF16String>antisignalnames;
	std::vector<CUTF16String>folderpaths;
	
#if VERSIONWIN
	std::vector<Params>params(watch_path_handles.size());
	
	static const __int64 SECS_BETWEEN_1601_AND_1970_EPOCHS = 116444736000000000LL;
	
	if(watch_path_handles.size() <= MAXIMUM_WAIT_OBJECTS)
	{
		size_t pos = 0;
		for(it = watch_path_handles.begin(); it != watch_path_handles.end(); ++it)
		{
			HANDLE h = *it;
			
			if(h != INVALID_HANDLE_VALUE)
			{
				Params *param = &params.at(pos);
				CUTF16String folderpath = watch_path_handle_paths.at(pos);
				param->h = h;
				std::wstring _a;
				generateUuid(_a);
				std::wstring _b;
				generateUuid(_b);
				unsigned char *p;
				DWORD data_len = _a.length() * sizeof(wchar_t);
				memset(param->a, 0x0, sizeof(param->a));
				memset(param->b, 0x0, sizeof(param->b));
				try
				{
					p = (unsigned char *)_a.c_str();
					CopyMemory(param->a, p, data_len);
					p = (unsigned char *)_b.c_str();
					CopyMemory(param->b, p, data_len);
				}
				catch (...)
				{
					
				}
				
				HANDLE a = CreateEvent(NULL,
					TRUE,
					FALSE,
					(const wchar_t *)param->a);

				if(a)
				{
					HANDLE h = (HANDLE)_beginthreadex(NULL /* security: handle not inherited */,
						0 /* stack size:default */,
						doIt,
						param /* arguments */,
						0 /* init flags:execute immediately */,
						NULL /* thread id */);
					if(h)
					{
						signals.push_back(a);
						handles.push_back(h);
						folderpaths.push_back(folderpath);
						antisignalnames.push_back(CUTF16String((PA_Unichar *)_b.c_str(), _b.length()));
					}
				}
			}
			pos++;
		}
		
		BOOL exit = FALSE;
		
		do {
			DWORD count = WaitForMultipleObjects(signals.size(), &signals[0], FALSE, 16);//1 tick = 0.0166666
			switch (count)
			{
				case WAIT_TIMEOUT:
					PA_YieldAbsolute();
					/* process (queued) file drop events */
					if (1)
					{
						std::lock_guard<std::mutex> lock(globalMutex0);/* =>listenerLoopExecuteMethod */
						
						if (MFD::CALLBACK_EVENT_PATHS.size())
						{
							listenerLoopExecuteMethod();
						}
					}
					PA_PutProcessToSleep(PA_GetCurrentProcessNumber(), CALLBACK_SLEEP_TIME);//59 ticks
					break;
				case WAIT_FAILED:
					exit = TRUE;
					break;
				default:
					for(DWORD i = count-WAIT_OBJECT_0; i < signals.size();++i)
					{
						HANDLE h = signals[i];
						CUTF16String path = folderpaths[i];
						ResetEvent(h);
						
						DWORD data_len = 0;
						DWORD len = sizeof(data_len);
						BOOL success = FALSE;
						
						HANDLE fmOut = CreateFileMapping(
							INVALID_HANDLE_VALUE,
							NULL,
							PAGE_READWRITE,
							0, len,
							L"PARAM_OUT");

						if (fmOut)
						{
							LPVOID bufOut = MapViewOfFile(fmOut, FILE_MAP_READ, 0, 0, len);
							if (bufOut)
							{
								unsigned char *p = (unsigned char *)bufOut;
								try
								{
									CopyMemory(&data_len, p, sizeof(data_len));
									success = TRUE;
								}
								catch (...)
								{
									
								}
								UnmapViewOfFile(bufOut);
							}
							CloseHandle(fmOut);
						}
						if ((success) && (data_len))
						{
							len = len + data_len;
							
							fmOut = CreateFileMapping(
								INVALID_HANDLE_VALUE,
								NULL,
								PAGE_READWRITE,
								0, len,
								L"PARAM_OUT");
							if (fmOut)
							{
								LPVOID bufOut = MapViewOfFile(fmOut, FILE_MAP_READ, 0, 0, len);
								if (bufOut)
								{
									unsigned char *p = (unsigned char *)bufOut;
									p = p + sizeof(data_len);
									try
									{
										std::vector<uint8_t>buf(data_len);
										CopyMemory(&buf[0], p, data_len);
										size_t pos = 0;
										FILE_NOTIFY_INFORMATION *fni = (FILE_NOTIFY_INFORMATION *)&buf[0];
										size_t nextEntryOffset = fni->NextEntryOffset;
										bool should_execute_method = false;

										do
										{
											/* get full path */
											DWORD fileNameLength = fni->FileNameLength;/* bytes */
											std::vector<PA_Unichar>ubuf(fileNameLength + sizeof(PA_Unichar));
											unsigned char *up = (unsigned char *)fni->FileName;
											CopyMemory(&ubuf[0], up, fileNameLength);
											CUTF16String event_path = path + CUTF16String((const PA_Unichar *)&ubuf[0]);

											/* get file time */
											if (0)
											{
												double t = 0;
												FILETIME ft;
												GetSystemTimeAsFileTime(&ft);
												ULARGE_INTEGER ul;
												ul.LowPart = ft.dwLowDateTime;
												ul.HighPart = ft.dwHighDateTime;
												double ts = (double)((ul.QuadPart - SECS_BETWEEN_1601_AND_1970_EPOCHS) / 10000);
											}

											bool is_good_event = false;
											switch (fni->Action)
											{
											case FILE_ACTION_RENAMED_NEW_NAME:
												is_good_event = true;
												break;
											case FILE_ACTION_MODIFIED:
											case FILE_ACTION_ADDED:
											case FILE_ACTION_REMOVED:
											case FILE_ACTION_RENAMED_OLD_NAME:
												break;
											default:
												break;
											}

											if (is_good_event)
											{
												if (!PathIsDirectory((LPCTSTR)event_path.c_str()))
												{
													bool is_good_extension = false;
													wchar_t	fDrive[_MAX_DRIVE], fDir[_MAX_DIR], fName[_MAX_FNAME], fExt[_MAX_EXT];
													HMODULE hplugin = GetModuleHandleW(THIS_BUNDLE_NAME);
													_wsplitpath_s((wchar_t *)event_path.c_str(), fDrive, fDir, fName, fExt);
													is_good_extension = !_wcsicmp(fExt, L".mht");

													if (is_good_extension)
													{
														bool is_good_name = false;
														is_good_name = (event_path.find((PA_Unichar *)L"~$") != 0);
														if (is_good_name)
														{
															should_execute_method = true;

															std::lock_guard<std::mutex> lock(globalMutex);

															MFD::CALLBACK_EVENT_PATHS.push_back(event_path);
														}
													}
												}
											}

											nextEntryOffset = fni->NextEntryOffset;
											pos += nextEntryOffset;

											fni = (FILE_NOTIFY_INFORMATION *)&buf.at(pos);

										} while (nextEntryOffset);

										if (should_execute_method)
										{
											/* execute callback method */
											listenerLoopExecuteMethod();
										}
										
									}
									catch (...)
									{
										
									}
									UnmapViewOfFile(bufOut);
								}
								CloseHandle(fmOut);
							}
							
						}
						
						CUTF16String antisignalname = antisignalnames.at(i);
						
						HANDLE antisignal = OpenEvent(EVENT_ALL_ACCESS, FALSE, (LPCWSTR)antisignalname.c_str());
						if(antisignal)
						{
							SetEvent(antisignal);
							CloseHandle(antisignal);
						}
					}
					break;
			}
			
			if (MFD::MONITOR_PROCESS_SHOULD_TERMINATE) exit = TRUE;
			
		} while (!exit);
		
		/* release signal handles */
		for(it = signals.begin(); it != signals.end(); ++it)
		{
			HANDLE h = *it;
			if(h != INVALID_HANDLE_VALUE)
			{
				CloseHandle(h);
			}
		}
		
		/* release thread handles */
		for (it = handles.begin(); it != handles.end(); ++it)
		{
			HANDLE h = *it;
			if (h != INVALID_HANDLE_VALUE)
			{
				CloseHandle(h);
			}
		}
	}
	
	/* release path handles */
	for(it = watch_path_handles.begin(); it != watch_path_handles.end(); ++it)
	{
		HANDLE h = *it;
		if(h != INVALID_HANDLE_VALUE)
		{
			CloseHandle(h);/* this will end the thread via bytesReturned == 0 */
		}
	}
#endif
	
	if(1)
	{
		std::lock_guard<std::mutex> lock(globalMutex);

		MFD::CALLBACK_EVENT_PATHS.clear();
		MFD::MONITOR_PROCESS_ID = 0;
	}
	
	PA_KillProcess();
}

void listenerLoopStart()
{
	std::lock_guard<std::mutex> lock(globalMutex0);/* =>listenerLoop */
	
	if (!MFD::MONITOR_PROCESS_ID)
	{
		MFD::MONITOR_PROCESS_ID = PA_NewProcess((void *)listenerLoop,
																						MFD::MONITOR_PROCESS_STACK_SIZE,
																						MFD::MONITOR_PROCESS_NAME);
	}
	else
	{
		listenerLoopFinish();
	}

}

void listenerLoopFinish()
{
	std::lock_guard<std::mutex> lock(globalMutex);
	
	if(MFD::MONITOR_PROCESS_ID)
	{
		MFD::MONITOR_PROCESS_SHOULD_TERMINATE = true;
		
		PA_YieldAbsolute();
		
		MFD::PROCESS_SHOULD_RESUME = true;
	}
}

void listenerLoopExecute()
{
	std::lock_guard<std::mutex> lock(globalMutex);
	
	MFD::MONITOR_PROCESS_SHOULD_TERMINATE = false;
	MFD::PROCESS_SHOULD_RESUME = true;
}

void listenerLoopExecuteMethod()
{
	std::lock_guard<std::mutex> lock(globalMutex);
	
	std::vector<CUTF16String>::iterator p = MFD::CALLBACK_EVENT_PATHS.begin();
	CUTF16String event_path = *p;
	
	method_id_t methodId = PA_GetMethodID((PA_Unichar *)MFD::WATCH_METHOD.getUTF16StringPtr());
	
	if(methodId)
	{
		PA_Variable	params[2];
		params[0] = PA_CreateVariable(eVK_Unistring);
		params[1] = PA_CreateVariable(eVK_Unistring);
		
		PA_Unistring path = PA_CreateUnistring((PA_Unichar *)event_path.c_str());
		PA_Unistring context = PA_CreateUnistring((PA_Unichar *)MFD::WATCH_CONTEXT.getUTF16StringPtr());

		PA_SetStringVariable(&params[0], &path);
		PA_SetStringVariable(&params[1], &context);

		MFD::CALLBACK_EVENT_PATHS.erase(p);
		
		PA_ExecuteMethodByID(methodId, params, 2);
		
		PA_ClearVariable(&params[0]);
		PA_ClearVariable(&params[1]);

	}else
	{
		PA_Variable	params[3];
		params[1] = PA_CreateVariable(eVK_Unistring);
		params[2] = PA_CreateVariable(eVK_Unistring);

		PA_Unistring path = PA_CreateUnistring((PA_Unichar *)event_path.c_str());
		PA_Unistring context = PA_CreateUnistring((PA_Unichar *)MFD::WATCH_CONTEXT.getUTF16StringPtr());

		PA_SetStringVariable(&params[1], &path);
		PA_SetStringVariable(&params[2], &context);
		
		params[0] = PA_CreateVariable(eVK_Unistring);
		PA_Unistring method = PA_CreateUnistring((PA_Unichar *)MFD::WATCH_METHOD.getUTF16StringPtr());
		PA_SetStringVariable(&params[0], &method);

		MFD::CALLBACK_EVENT_PATHS.erase(p);
		
		PA_ExecuteCommandByID(1007, params, 2);
		
		PA_ClearVariable(&params[0]);
		PA_ClearVariable(&params[1]);
		PA_ClearVariable(&params[2]);
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
	
		std::wstring m_ScriptPath;
		std::wstring m_ExportPath;
	
	void getTempPath(std::wstring &path)
	{
		std::vector<wchar_t>buf(MAX_PATH + 1);
		DWORD len = GetTempPath((DWORD)buf.size(), (LPTSTR)&buf[0]);
		if(len)
			path = std::wstring((const wchar_t *)&buf[0], len);
	}
		
	void getScriptPath(std::wstring &path)
	{
		wchar_t	fDrive[_MAX_DRIVE], fDir[_MAX_DIR], fName[_MAX_FNAME], fExt[_MAX_EXT];
		wchar_t thisPath[_MAX_PATH] = {0};
		
		HMODULE hplugin = GetModuleHandleW(THIS_BUNDLE_NAME);
		GetModuleFileNameW(hplugin, thisPath, _MAX_PATH);
		_wsplitpath_s(thisPath, fDrive, fDir, fName, fExt);

		path = fDrive;
		path += fDir;
		
		/* remove delimiter to go one level up the hierarchy */
		if(path.at(path.size() - 1) == L'\\')
			path = path.substr(0, path.size() - 1);
		
		_wsplitpath_s(path.c_str(), fDrive, fDir, fName, fExt);
		
		/* ../Resources/outlook_export_selected_messages.vbs */
		path = fDrive;
		path+= fDir;
		path+= L"Resources\\";
		path+= VBS_FILE_NAME;
	}
	
	void getExportRootPath(std::wstring &path)
	{
		getTempPath(path);
		
		std::wstring uuid;
		generateUuid(uuid);
		path += uuid;
		path += L"\\";
	}
	
	void getExportPath(std::wstring &path)
	{
		std::wstring uuid;
		generateUuid(uuid);
		
		SHCreateDirectory(NULL, (PCWSTR)m_ExportPath.c_str());

		path = L"\"" + m_ScriptPath;
		path += L"\" \"";
		path += m_ExportPath;
		path+= uuid;
		path+= L"\"";
	}
	
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
	
	bool isFilePromiseDrop(IDataObject *pDataObj) {
		
		FORMATETC formatetc;
		
		formatetc.cfFormat = m_cfFormatFileDescriptor;
		formatetc.ptd = NULL;
		formatetc.dwAspect = DVASPECT_CONTENT;
		formatetc.lindex = -1;
		formatetc.tymed = TYMED_HGLOBAL;
		
		if (S_OK == pDataObj->QueryGetData(&formatetc))
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
	
	public:
	
		MyDropTarget::MyDropTarget() : m_cRefCount(1), m_hWnd(NULL) {
			m_cfFormatFileDescriptor = RegisterClipboardFormat(CFSTR_FILEDESCRIPTOR);
			m_cfFormatRPSourceFolder = RegisterClipboardFormat(L"RenPrivateSourceFolder");
			m_cfFormatRPItem = RegisterClipboardFormat(L"RenPrivateItem");
			m_cfFormatRPMessages = RegisterClipboardFormat(L"RenPrivateMessages");
			m_cfFormatRPLatestMessages = RegisterClipboardFormat(L"RenPrivateLatestMessages");
			
			getScriptPath(m_ScriptPath);

			/* create root folder here (script creates only 1 level for each drag and drop) */
			getExportRootPath(m_ExportPath);			
			SHCreateDirectory(NULL, (PCWSTR)m_ExportPath.c_str());

			MFD::WATCH_PATHS.setSize(0);
			MFD::WATCH_PATHS.appendUTF16String((const PA_Unichar *)m_ExportPath.c_str());
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

		if (this->isOutlookDrop(pDataObj))
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
		FORMATETC formatetc;
		if (this->isFileDrop(pDataObj, &formatetc))
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
								std::vector<unsigned char>buf((len)*sizeof(wchar_t));
								if (DragQueryFile(hDrop, i, (LPTSTR)&buf[0], len))
								{
									std::lock_guard<std::mutex> lock(globalMutex);

									MFD::CALLBACK_EVENT_PATHS.push_back(CUTF16String((const PA_Unichar *)&buf[0], len));
								}
							}
						}
						return S_OK;
					}
				}
			}
		}

		if (this->isOutlookDrop(pDataObj))
		{
			std::wstring path;
			getExportPath(path);
			ShellExecute(NULL, L"open", L"cscript.exe", path.c_str(), NULL, SW_HIDE);
			return S_OK;
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
		MFD::WATCH_METHOD.fromParamAtIndex(pParams, 2);
		MFD::WATCH_CONTEXT.fromParamAtIndex(pParams, 3);

		HWND hwnd = (HWND)PA_GetHWND((PA_WindowRef)Param1_window.getIntValue());
		
		if(hwnd)
		{
			g_MyDropTarget.Register(hwnd);
			
			listenerLoopStart();
		}
		else
		{
			g_MyDropTarget.Unregister();

			//listenerLoopFinish(); 
		}
	
	}/* IsProcessOnExit */
#endif
}