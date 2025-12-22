# Building libxlsxwriter as a 32-bit DLL on Windows

This guide explains how to build libxlsxwriter as a 32-bit DLL with zlib statically linked.

## Prerequisites

- Visual Studio 2026 (or compatible version)
- CMake
- Git

## Step 1: Open Developer Command Prompt

Open a command prompt and run:

```batch
"C:\Program Files\Microsoft Visual Studio\18\Community\VC\Auxiliary\Build\vcvars32.bat"
```

## Step 2: Clone and build zlib as a static library

Clone the zlib repository:

```batch
git clone https://github.com/madler/zlib.git
cd zlib
mkdir build32-static
cd build32-static
```

Configure and build zlib:

```batch
cmake .. -G "Visual Studio 18 2026" -A Win32 -DCMAKE_BUILD_TYPE=Release -DCMAKE_INSTALL_PREFIX="C:/zlib32-static" -DBUILD_SHARED_LIBS=OFF -DCMAKE_C_FLAGS_RELEASE="/MT /O2"
```

```batch
cmake --build . --config Release && cmake --install . --config Release
```

This installs zlib to `C:\zlib32-static` with the static runtime.

## Step 3: Build libxlsxwriter

Navigate to the libxlsxwriter source directory and create a build folder:

```batch
cd C:\path\to\libxlsxwriter
mkdir build32
cd build32
```

Configure and build libxlsxwriter:

```batch
cmake .. -G "Visual Studio 18 2026" -A Win32 -DBUILD_SHARED_LIBS=ON -DZLIB_ROOT="C:/zlib32-static" -DZLIB_LIBRARY="C:/zlib32-static/lib/zlibstatic.lib" -DCMAKE_C_FLAGS_RELEASE="/MT /O2" -DCMAKE_BUILD_TYPE=Release
```

```batch
cmake --build . --config Release
```

## Output

The resulting DLL will be located at:

```
build32\Release\xlsxwriter.dll
```

This DLL has zlib statically linked and does not require zlib.dll to be present.

## Notes

- The `/MT` flag links against the static C runtime, eliminating the need for separate runtime DLLs
- Ensure both zlib and libxlsxwriter are built with the same runtime settings (`/MT` for static, `/MD` for dynamic)
- For 64-bit builds, change `-A Win32` to `-A x64` and update folder names accordingly
