"""Copy Edge's locked cookie database using Windows raw file API."""
import ctypes
import ctypes.wintypes
import os
import sys

GENERIC_READ = 0x80000000
FILE_SHARE_READ = 0x1
FILE_SHARE_WRITE = 0x2
FILE_SHARE_DELETE = 0x4
OPEN_EXISTING = 3
FILE_ATTRIBUTE_NORMAL = 0x80
INVALID_HANDLE_VALUE = ctypes.wintypes.HANDLE(-1).value

kernel32 = ctypes.windll.kernel32

src = os.path.expandvars(
    r"%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\Network\Cookies"
)
dst = os.path.join(os.environ["TEMP"], "edge_cookies_copy.db")

handle = kernel32.CreateFileW(
    src,
    GENERIC_READ,
    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
    None,
    OPEN_EXISTING,
    FILE_ATTRIBUTE_NORMAL,
    None,
)

if handle == INVALID_HANDLE_VALUE:
    err = ctypes.get_last_error()
    print(f"CreateFileW failed with error {err}")
    sys.exit(1)

large_int = ctypes.wintypes.LARGE_INTEGER(0)
kernel32.GetFileSizeEx(handle, ctypes.byref(large_int))
file_size = large_int.value
print(f"Opened {src} ({file_size} bytes)")

CHUNK = 65536
chunks = []
while True:
    buf = ctypes.create_string_buffer(CHUNK)
    bytes_read = ctypes.wintypes.DWORD(0)
    ok = kernel32.ReadFile(handle, buf, CHUNK, ctypes.byref(bytes_read), None)
    if not ok or bytes_read.value == 0:
        break
    chunks.append(buf.raw[:bytes_read.value])
kernel32.CloseHandle(handle)
data = b"".join(chunks)
print(f"Read {len(data)} bytes")

with open(dst, "wb") as f:
    f.write(data)

print(f"Copied {len(data)} bytes to {dst}")
