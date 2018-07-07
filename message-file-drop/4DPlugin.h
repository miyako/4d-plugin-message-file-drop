/* --------------------------------------------------------------------------------
 #
 #	4DPlugin.h
 #	source generated by 4D Plugin Wizard
 #	Project : Message file drop
 #	author : miyako
 #	2018/06/21
 #
 # --------------------------------------------------------------------------------*/

#if VERSIONWIN
#include <shlobj.h> /* SHCreateDirectory */
#include <Shlwapi.h>
#include <process.h>
#define BUF_SIZE 32768 /* max=64KB */
#define CALLBACK_SLEEP_TIME 59
#define THIS_BUNDLE_NAME L"Message file drop.4DX"
#define VBS_FILE_NAME L"outlook_export_selected_messages.vbs"
#endif

#include <mutex>
#include <algorithm>

// --- Message file drop
void ACCEPT_MESSAGE_FILES(sLONG_PTR *pResult, PackagePtr pParams);

bool IsProcessOnExit();
void OnStartup();
void OnCloseProcess();

void listenerLoop(void);
void listenerLoopStart(void);
void listenerLoopFinish(void);
void listenerLoopExecute(void);
void listenerLoopExecuteMethod(void);

typedef PA_long32 process_number_t;
typedef PA_long32 process_stack_size_t;
typedef PA_long32 method_id_t;
typedef PA_Unichar* process_name_t;

#if VERSIONWIN
typedef struct
{
	HANDLE h;
	wchar_t a[40];//UUID
	wchar_t b[40];//UUID
}Params;
#endif
