import os
import subprocess
import time

SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
PORT = 2002
PROFILE = "/tmp/lo-python-profile"

URL = f"uno:socket,host=127.0.0.1,port={PORT};" "urp;StarOffice.ComponentContext"


def _soffice_process():
    return subprocess.Popen(
        [
            SOFFICE,
            "--nologo",
            "--nodefault",
            "--norestore",
            f"-env:UserInstallation=file://{PROFILE}",
            f"--accept=socket,host=127.0.0.1,port={PORT};"
            "urp;StarOffice.ComponentContext",
        ],
        env={},
    )


def connect():
    import uno

    local = uno.getComponentContext()

    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        local,
    )

    # First try the existing server.
    try:
        print(f"Connecting to existing LibreOffice at {URL}")
        return resolver.resolve(URL)
    except Exception:
        print("No existing LibreOffice server found.")

    # Start one.
    print("Starting LibreOffice...")
    _soffice_process()

    # Wait until the socket is actually ready.
    for i in range(50):
        try:
            print(f"Attempt {i + 1}: connecting...")
            return resolver.resolve(URL)
        except Exception:
            time.sleep(1)

    raise RuntimeError("LibreOffice did not start")
