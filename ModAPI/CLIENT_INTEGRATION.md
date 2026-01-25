# ScalA Cursor Mod - Client Integration Guide

## Overview

This mod provides cursor position transformation for ScalA's zoom feature without requiring SDL2 DLL replacement. The Astonia client loads this mod and calls its functions at specific points in the SDL mouse handling code.

## Required Changes in Client

### 1. Load the Mod DLL

On client startup, after SDL initialization:

```c
typedef int (*PFN_ScalaMod_Init)(uint32_t);
typedef void (*PFN_ScalaMod_Shutdown)(void);
typedef int (*PFN_ScalaMod_TransformMousePos)(int*, int*);
typedef int (*PFN_ScalaMod_InverseTransformPos)(int*, int*);

static HMODULE g_hScalaMod = NULL;
static PFN_ScalaMod_Init ScalaMod_Init = NULL;
static PFN_ScalaMod_Shutdown ScalaMod_Shutdown = NULL;
static PFN_ScalaMod_TransformMousePos ScalaMod_TransformMousePos = NULL;
static PFN_ScalaMod_InverseTransformPos ScalaMod_InverseTransformPos = NULL;

void LoadScalaMod(void) {
    g_hScalaMod = LoadLibraryA("scala_cursor_mod.dll");
    if (!g_hScalaMod) return;  /* Mod not installed, that's OK */

    ScalaMod_Init = (PFN_ScalaMod_Init)GetProcAddress(g_hScalaMod, "ScalaMod_Init");
    ScalaMod_Shutdown = (PFN_ScalaMod_Shutdown)GetProcAddress(g_hScalaMod, "ScalaMod_Shutdown");
    ScalaMod_TransformMousePos = (PFN_ScalaMod_TransformMousePos)GetProcAddress(g_hScalaMod, "ScalaMod_TransformMousePos");
    ScalaMod_InverseTransformPos = (PFN_ScalaMod_InverseTransformPos)GetProcAddress(g_hScalaMod, "ScalaMod_InverseTransformPos");

    if (ScalaMod_Init) {
        int version = ScalaMod_Init(GetCurrentProcessId());
        if (version == 0) {
            /* Init failed, unload */
            FreeLibrary(g_hScalaMod);
            g_hScalaMod = NULL;
        }
    }
}

void UnloadScalaMod(void) {
    if (ScalaMod_Shutdown) ScalaMod_Shutdown();
    if (g_hScalaMod) FreeLibrary(g_hScalaMod);
    g_hScalaMod = NULL;
}
```

### 2. Hook SDL_GetMouseState

Wherever the client calls `SDL_GetMouseState`, add the transformation:

```c
/* BEFORE (original code): */
Uint32 buttons = SDL_GetMouseState(&mouseX, &mouseY);

/* AFTER (with mod support): */
Uint32 buttons = SDL_GetMouseState(&mouseX, &mouseY);
if (ScalaMod_TransformMousePos) {
    ScalaMod_TransformMousePos(&mouseX, &mouseY);
}
```

### 3. Hook SDL_GetGlobalMouseState

Same pattern:

```c
Uint32 buttons = SDL_GetGlobalMouseState(&mouseX, &mouseY);
if (ScalaMod_TransformMousePos) {
    ScalaMod_TransformMousePos(&mouseX, &mouseY);
}
```

### 4. Hook SDL_WarpMouseInWindow (if used)

For cursor warping, use inverse transformation:

```c
/* BEFORE: */
SDL_WarpMouseInWindow(window, x, y);

/* AFTER: */
if (ScalaMod_InverseTransformPos) {
    ScalaMod_InverseTransformPos(&x, &y);
}
SDL_WarpMouseInWindow(window, x, y);
```

## Functions to Hook

| SDL Function | Mod Function | When to Call |
|--------------|--------------|--------------|
| `SDL_GetMouseState` | `ScalaMod_TransformMousePos` | AFTER SDL call |
| `SDL_GetGlobalMouseState` | `ScalaMod_TransformMousePos` | AFTER SDL call |
| `SDL_WarpMouseInWindow` | `ScalaMod_InverseTransformPos` | BEFORE SDL call |

## Jitter Fix

The mod uses **rounded integer division** instead of truncating division to reduce cursor jitter at high scale factors:

```c
/* Old (causes jitter): */
*x = (relX * clientW) / viewportW;

/* New (reduced jitter): */
*x = (relX * clientW + viewportW / 2) / viewportW;
```

This adds 0.5 before truncating, effectively rounding to nearest instead of always rounding down.

## Deployment

1. Build `scala_cursor_mod.dll`
2. Place in same directory as the game executable
3. Client loads it automatically on startup (if present)
4. ScalA's existing ZoomStateIPC handles communication

## Compatibility

- The mod gracefully handles ScalA not running (returns 0, no transformation)
- Version checking via `SCALA_MOD_API_VERSION`
- Thread-safe with critical sections
- No changes needed to ScalA itself - uses existing shared memory

## Build

```
cl /LD /O2 scala_cursor_mod.c /Fe:scala_cursor_mod.dll
```

Or with MinGW:
```
gcc -shared -O2 scala_cursor_mod.c -o scala_cursor_mod.dll
```
