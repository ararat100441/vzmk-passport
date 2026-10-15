"""Build-host workaround: use Python's native Windows-version fallback.

WMI hangs on this host (faulthandler evidence retained). No version is invented,
no registry/service changes, and this hook is absent from the frozen product.
"""
import platform
def unavailable(*args,**kwargs):raise OSError('WMI unavailable on build host; use native fallback')
platform._wmi_query=unavailable
