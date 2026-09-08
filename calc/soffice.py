import subprocess
import time

SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
PORT = 2002
PROFILE = "/tmp/lo-python-profile"

URL = f"uno:socket,host=127.0.0.1,port={PORT};" "urp;StarOffice.ComponentContext"


def _soffice_process():
    """Actual soffice subprocess started from python"""
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
    """Connect to existing soffice or create one and try to connect to it"""
    import uno

    local = uno.getComponentContext()

    resolver = local.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver",
        local,
    )

    # First try the existing server.
    try:
        return resolver.resolve(URL)
    except Exception:
        print("No existing LibreOffice server found.")

    # Start one.
    _soffice_process()

    # Wait until the socket is actually ready.
    for i in range(50):
        try:
            return resolver.resolve(URL)
        except Exception:
            time.sleep(1)

    raise RuntimeError("LibreOffice did not start")
