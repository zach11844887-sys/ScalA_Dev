/*
 * ScalA Cursor Mod API v1.0
 *
 * Minimal API for Astonia client to support cursor position transformation
 * at higher ScalA scaling levels without requiring SDL2 DLL replacement.
 *
 * Implementation: The client loads this mod DLL and calls the exported functions.
 * ScalA communicates viewport state via shared memory (existing ZoomStateIPC).
 */

#ifndef SCALA_CURSOR_MOD_H
#define SCALA_CURSOR_MOD_H

#include <stdint.h>

#ifdef SCALA_MOD_EXPORTS
#define SCALA_MOD_API __declspec(dllexport)
#else
#define SCALA_MOD_API __declspec(dllimport)
#endif

#ifdef __cplusplus
extern "C" {
#endif

/*
 * =============================================================================
 * API VERSION
 * =============================================================================
 */
#define SCALA_MOD_API_VERSION 1

/*
 * =============================================================================
 * LIFECYCLE FUNCTIONS (called by client)
 * =============================================================================
 */

/*
 * ScalaMod_Init
 *
 * Called once when the client loads the mod.
 * The mod should initialize shared memory connection here.
 *
 * Parameters:
 *   clientPID - The process ID of the client (used for shared memory name)
 *
 * Returns:
 *   API version number on success, 0 on failure
 */
SCALA_MOD_API int ScalaMod_Init(uint32_t clientPID);

/*
 * ScalaMod_Shutdown
 *
 * Called when the client unloads the mod or exits.
 * The mod should clean up resources here.
 */
SCALA_MOD_API void ScalaMod_Shutdown(void);

/*
 * =============================================================================
 * CURSOR TRANSFORMATION FUNCTIONS (called by client)
 * =============================================================================
 */

/*
 * ScalaMod_TransformMousePos
 *
 * Transforms cursor position from viewport (screen) space to client space.
 * Called by client AFTER calling real SDL_GetMouseState/SDL_GetGlobalMouseState.
 *
 * Parameters:
 *   x, y - Pointer to coordinates (modified in place)
 *
 * Returns:
 *   1 if transformation was applied, 0 if passthrough (no active ScalA viewport)
 *
 * Example client integration:
 *   Uint32 buttons = Real_SDL_GetMouseState(&x, &y);
 *   ScalaMod_TransformMousePos(&x, &y);
 *   return buttons;
 */
SCALA_MOD_API int ScalaMod_TransformMousePos(int* x, int* y);

/*
 * ScalaMod_InverseTransformPos
 *
 * Transforms cursor position from client space to viewport (screen) space.
 * Called by client BEFORE calling real SDL_WarpMouseInWindow.
 *
 * Parameters:
 *   x, y - Pointer to coordinates (modified in place)
 *
 * Returns:
 *   1 if transformation was applied, 0 if passthrough
 *
 * Example client integration:
 *   ScalaMod_InverseTransformPos(&x, &y);
 *   Real_SDL_WarpMouseInWindow(window, x, y);
 */
SCALA_MOD_API int ScalaMod_InverseTransformPos(int* x, int* y);

/*
 * =============================================================================
 * QUERY FUNCTIONS (optional, for debugging/UI)
 * =============================================================================
 */

/*
 * ScalaMod_IsActive
 *
 * Check if ScalA is actively managing this client's viewport.
 *
 * Returns:
 *   1 if active, 0 if not
 */
SCALA_MOD_API int ScalaMod_IsActive(void);

/*
 * ScalaMod_GetScaleFactor
 *
 * Get the current scale factor (for UI display purposes).
 *
 * Parameters:
 *   scaleX, scaleY - Pointers to receive scale factors (1.0 = no scaling)
 *
 * Returns:
 *   1 if values were set, 0 if not active
 */
SCALA_MOD_API int ScalaMod_GetScaleFactor(float* scaleX, float* scaleY);

#ifdef __cplusplus
}
#endif

#endif /* SCALA_CURSOR_MOD_H */
