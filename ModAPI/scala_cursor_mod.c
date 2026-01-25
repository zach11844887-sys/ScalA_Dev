/*
 * ScalA Cursor Mod Implementation v1.0
 *
 * Lightweight cursor transformation mod for Astonia client.
 * Connects to ScalA's shared memory and provides coordinate transformation.
 */

#define SCALA_MOD_EXPORTS
#define WIN32_LEAN_AND_MEAN
#define _CRT_SECURE_NO_WARNINGS

#include <windows.h>
#include <stdio.h>
#include "scala_cursor_mod.h"

/*
 * =============================================================================
 * SHARED MEMORY STRUCTURES (must match ScalA's ZoomStateIPC)
 * =============================================================================
 */

#define STRUCT_VERSION 2
#define DEFAULT_MAX_INSTANCES 32
#define MAX_MAPPED_INSTANCES 256
#define HEADER_SIZE 64
#define ENTRY_SIZE 64

#pragma pack(push, 1)

typedef struct {
    int version;
    int count;
    int versionMismatch;
    int maxInstances;
    int reserved[12];
} ScalAZoomHeader;

typedef struct {
    int scalaPID;
    int viewportX;
    int viewportY;
    int viewportW;
    int viewportH;
    int clientW;
    int clientH;
    int enabled;
    int reserved[8];
} ScalAZoomEntry;

#pragma pack(pop)

/*
 * =============================================================================
 * GLOBAL STATE
 * =============================================================================
 */

static HANDLE g_hMapFile = NULL;
static void* g_pMappedMem = NULL;
static uint32_t g_clientPID = 0;
static CRITICAL_SECTION g_lock;
static int g_initialized = 0;

/*
 * =============================================================================
 * HELPER FUNCTIONS
 * =============================================================================
 */

static ScalAZoomHeader* GetHeader(void) {
    return (ScalAZoomHeader*)g_pMappedMem;
}

static ScalAZoomEntry* GetEntry(int index) {
    if (!g_pMappedMem) return NULL;
    char* base = (char*)g_pMappedMem;
    return (ScalAZoomEntry*)(base + HEADER_SIZE + (index * ENTRY_SIZE));
}

static int ConnectSharedMem(void) {
    if (g_hMapFile) return 1;
    if (g_clientPID == 0) return 0;

    char name[64];
    sprintf(name, "ScalA_ZoomState_%lu", (unsigned long)g_clientPID);

    g_hMapFile = OpenFileMappingA(FILE_MAP_READ, FALSE, name);
    if (!g_hMapFile) return 0;

    size_t mapSize = HEADER_SIZE + (ENTRY_SIZE * MAX_MAPPED_INSTANCES);
    g_pMappedMem = MapViewOfFile(g_hMapFile, FILE_MAP_READ, 0, 0, mapSize);

    if (!g_pMappedMem) {
        CloseHandle(g_hMapFile);
        g_hMapFile = NULL;
        return 0;
    }

    /* Check version compatibility */
    ScalAZoomHeader* hdr = GetHeader();
    if (hdr->version != STRUCT_VERSION) {
        /* Version mismatch - still usable but may have issues */
    }

    return 1;
}

static void DisconnectSharedMem(void) {
    if (g_pMappedMem) {
        UnmapViewOfFile(g_pMappedMem);
        g_pMappedMem = NULL;
    }
    if (g_hMapFile) {
        CloseHandle(g_hMapFile);
        g_hMapFile = NULL;
    }
}

/*
 * FindActiveEntry
 *
 * Find the ScalA instance whose viewport contains the current cursor position.
 * Returns NULL if cursor is not in any managed viewport.
 */
static ScalAZoomEntry* FindActiveEntry(void) {
    if (!g_pMappedMem) {
        if (!ConnectSharedMem()) return NULL;
    }
    if (!g_pMappedMem) return NULL;

    POINT pt;
    GetCursorPos(&pt);

    ScalAZoomHeader* hdr = GetHeader();

    int maxInstances = hdr->maxInstances;
    if (maxInstances <= 0) maxInstances = DEFAULT_MAX_INSTANCES;
    if (maxInstances > MAX_MAPPED_INSTANCES) maxInstances = MAX_MAPPED_INSTANCES;

    int count = hdr->count;
    if (count <= 0 || count > maxInstances) count = maxInstances;

    for (int i = 0; i < count; i++) {
        ScalAZoomEntry* e = GetEntry(i);
        if (!e) continue;

        if (e->scalaPID == 0 || !e->enabled) continue;
        if (e->viewportW <= 0 || e->viewportH <= 0) continue;
        if (e->clientW <= 0 || e->clientH <= 0) continue;

        /* Check if cursor is inside this viewport */
        if (pt.x >= e->viewportX && pt.x < e->viewportX + e->viewportW &&
            pt.y >= e->viewportY && pt.y < e->viewportY + e->viewportH) {
            return e;
        }
    }

    return NULL;
}

/*
 * =============================================================================
 * API IMPLEMENTATION
 * =============================================================================
 */

SCALA_MOD_API int ScalaMod_Init(uint32_t clientPID) {
    if (g_initialized) return SCALA_MOD_API_VERSION;

    InitializeCriticalSection(&g_lock);
    g_clientPID = clientPID;
    g_initialized = 1;

    /* Try to connect now, but don't fail if ScalA isn't running yet */
    ConnectSharedMem();

    return SCALA_MOD_API_VERSION;
}

SCALA_MOD_API void ScalaMod_Shutdown(void) {
    if (!g_initialized) return;

    EnterCriticalSection(&g_lock);
    DisconnectSharedMem();
    g_clientPID = 0;
    g_initialized = 0;
    LeaveCriticalSection(&g_lock);

    DeleteCriticalSection(&g_lock);
}

SCALA_MOD_API int ScalaMod_TransformMousePos(int* x, int* y) {
    if (!g_initialized || !x || !y) return 0;

    EnterCriticalSection(&g_lock);

    ScalAZoomEntry* entry = FindActiveEntry();
    if (!entry) {
        LeaveCriticalSection(&g_lock);
        return 0;
    }

    POINT pt;
    GetCursorPos(&pt);

    /* Calculate position relative to viewport */
    int relX = pt.x - entry->viewportX;
    int relY = pt.y - entry->viewportY;

    /* Transform to client space */
    /* Using fixed-point arithmetic to reduce jitter from integer division */
    *x = (relX * entry->clientW + entry->viewportW / 2) / entry->viewportW;
    *y = (relY * entry->clientH + entry->viewportH / 2) / entry->viewportH;

    /* Clamp to valid range */
    if (*x < 0) *x = 0;
    if (*y < 0) *y = 0;
    if (*x >= entry->clientW) *x = entry->clientW - 1;
    if (*y >= entry->clientH) *y = entry->clientH - 1;

    LeaveCriticalSection(&g_lock);
    return 1;
}

SCALA_MOD_API int ScalaMod_InverseTransformPos(int* x, int* y) {
    if (!g_initialized || !x || !y) return 0;

    EnterCriticalSection(&g_lock);

    ScalAZoomEntry* entry = FindActiveEntry();
    if (!entry) {
        LeaveCriticalSection(&g_lock);
        return 0;
    }

    /* Transform from client space to viewport space */
    int viewportX = (*x * entry->viewportW + entry->clientW / 2) / entry->clientW;
    int viewportY = (*y * entry->viewportH + entry->clientH / 2) / entry->clientH;

    /* Add viewport offset to get screen coordinates */
    *x = entry->viewportX + viewportX;
    *y = entry->viewportY + viewportY;

    LeaveCriticalSection(&g_lock);
    return 1;
}

SCALA_MOD_API int ScalaMod_IsActive(void) {
    if (!g_initialized) return 0;

    EnterCriticalSection(&g_lock);
    ScalAZoomEntry* entry = FindActiveEntry();
    int active = (entry != NULL);
    LeaveCriticalSection(&g_lock);

    return active;
}

SCALA_MOD_API int ScalaMod_GetScaleFactor(float* scaleX, float* scaleY) {
    if (!g_initialized || !scaleX || !scaleY) return 0;

    EnterCriticalSection(&g_lock);

    ScalAZoomEntry* entry = FindActiveEntry();
    if (!entry) {
        LeaveCriticalSection(&g_lock);
        return 0;
    }

    *scaleX = (float)entry->viewportW / (float)entry->clientW;
    *scaleY = (float)entry->viewportH / (float)entry->clientH;

    LeaveCriticalSection(&g_lock);
    return 1;
}

/*
 * =============================================================================
 * DLL ENTRY POINT
 * =============================================================================
 */

BOOL WINAPI DllMain(HINSTANCE hInst, DWORD reason, LPVOID reserved) {
    (void)hInst; (void)reserved;

    if (reason == DLL_PROCESS_ATTACH) {
        DisableThreadLibraryCalls(hInst);
    }

    return TRUE;
}
