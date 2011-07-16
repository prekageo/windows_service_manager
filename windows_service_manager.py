"""
This is a GUI utility that can manage Windows services. It can display installed
services and kernel drivers. It can start and stop multiple services at once.
It can, also, alter the startup type of a service.
"""

import ctypes
import win32service as winsvc
import win32api
import winerror
import pywintypes

import wx

class SERVICE_STATUS_PROCESS(ctypes.Structure):
  """
  The SERVICE_STATUS_PROCESS Structure.
  http://msdn.microsoft.com/en-us/library/ms685992%28v=vs.85%29.aspx

  Needed until this revision of pywin32 gets released:
  http://pywin32.hg.sourceforge.net/hgweb/pywin32/pywin32/rev/7f3c50215fa5
  """
  _fields_ = [('dwServiceType', ctypes.c_uint),
              ('dwCurrentState', ctypes.c_uint),
              ('dwControlsAccepted', ctypes.c_uint),
              ('dwWin32ExitCode', ctypes.c_uint),
              ('dwServiceSpecificExitCode', ctypes.c_uint),
              ('dwCheckPoint', ctypes.c_uint),
              ('dwWaitHint', ctypes.c_uint),
              ('dwProcessId', ctypes.c_uint),
              ('dwServiceFlags', ctypes.c_uint)]

class ENUM_SERVICE_STATUS_PROCESS(ctypes.Structure):
  """
  The ENUM_SERVICE_STATUS_PROCESS Structure.
  http://msdn.microsoft.com/en-us/library/ms682648%28v=vs.85%29.aspx

  Needed until this revision of pywin32 gets released:
  http://pywin32.hg.sourceforge.net/hgweb/pywin32/pywin32/rev/7f3c50215fa5
  """
  _fields_ = [('lpServiceName', ctypes.c_wchar_p),
              ('lpDisplayName', ctypes.c_wchar_p),
              ('ServiceStatusProcess', SERVICE_STATUS_PROCESS)]

class Service:
  """
  This class provides class methods for enumerating services, etc. and also
  instance methods for manipulating individual services.
  """

  """ A handle to the Service Control Manager. """
  scm_handle = winsvc.OpenSCManager(None, None, winsvc.SC_MANAGER_ALL_ACCESS)

  """
  A list of protected services that are crucial for normal operation of the OS.
  User cannot manipulate them.
  """
  protected_services = [
    'DcomLaunch',
    'Eventlog',
    'PlugPlay',
    'RpcSs',
    'SamSs',
  ]

  @classmethod
  def get_all(cls):
    """ Get all services and return list of Service objects. """
    SC_ENUM_PROCESS_INFO = 0

    services = cls.EnumServicesStatusExW(cls.scm_handle.handle,
        winsvc.SERVICE_DRIVER | winsvc.SERVICE_WIN32, winsvc.SERVICE_STATE_ALL,
        None, SC_ENUM_PROCESS_INFO)

    ret = []
    for service in services:
      ret.append(Service.create(**service))
    return ret

  @classmethod
  def EnumServicesStatusExW(cls, SCManager, ServiceType, ServiceState,
      GroupName, InfoLevel):
    """
    Needed until this revision of pywin32 gets released:
    http://pywin32.hg.sourceforge.net/hgweb/pywin32/pywin32/rev/7f3c50215fa5
    """
    pcbBytesNeeded = ctypes.c_uint()
    cbBufSize = 0
    lpServicesReturned = ctypes.c_uint()
    lpServices = None
    lpResumeHandle = ctypes.c_uint()
    ret = []

    while True:
      api_ret = ctypes.windll.advapi32.EnumServicesStatusExW(SCManager,
          InfoLevel, ServiceType, ServiceState, lpServices, cbBufSize,
          ctypes.byref(pcbBytesNeeded), ctypes.byref(lpServicesReturned),
          ctypes.byref(lpResumeHandle), GroupName)
      if not api_ret:
        err = win32api.GetLastError()
        if err != winerror.ERROR_MORE_DATA :
          raise err
      lpServices_ = ctypes.cast(lpServices,
        ctypes.POINTER(ENUM_SERVICE_STATUS_PROCESS))
      for i in xrange(lpServicesReturned.value):
        ssp = lpServices_[i].ServiceStatusProcess
        ret.append({
          'ServiceName'             : lpServices_[i].lpServiceName,
          'DisplayName'             : lpServices_[i].lpDisplayName,
          'ServiceType'             : ssp.dwServiceType,
          'CurrentState'            : ssp.dwCurrentState,
          'ControlsAccepted'        : ssp.dwControlsAccepted,
          'Win32ExitCode'           : ssp.dwWin32ExitCode,
          'ServiceSpecificExitCode' : ssp.dwServiceSpecificExitCode,
          'CheckPoint'              : ssp.dwCheckPoint,
          'WaitHint'                : ssp.dwWaitHint,
          'ProcessId'               : ssp.dwProcessId,
          'ServiceFlags'            : ssp.dwServiceFlags,
        })
      if api_ret:
        break
      cbBufSize = pcbBytesNeeded.value
      lpServices = (ctypes.c_ubyte * cbBufSize)()
    return ret

  @classmethod
  def filter(cls, services, **kwargs):
    """
    Filter services and return only those that match the AND bitwise operation
    on the key value pairs given.

    e.g. Service.filter(services, ServiceType = winsvc.SERVICE_DRIVER)
    """
    for service in services:
      for key,value in kwargs.iteritems():
        if not (getattr(service, key) & value):
          break
      else:
        yield service

  @classmethod
  def sort(cls, services, *args):
    """
    Sort services according to the keys given. The sorting algorithm handles
    specially specific keys.

    e.g. Service.sort(drivers, 'CurrentState', 'ControlsAccepted', 'StartType')
    """
    def cmp_func(x, y):
      for key in args:
        if key in ['CurrentState']:
          if getattr(x, key) != getattr(y, key):
            return -1 if getattr(x, key) > getattr(y, key) else 1
        elif key in ['StartType']:
          if getattr(x, key) != getattr(y, key):
            return -1 if getattr(x, key) < getattr(y, key) else 1
        elif key == 'ControlsAccepted':
          a = getattr(x, key) & winsvc.SERVICE_CONTROL_STOP
          b = getattr(y, key) & winsvc.SERVICE_CONTROL_STOP
          if a != b:
            return -1 if a else 1
      return 0
    return sorted(services, cmp = cmp_func)

  @classmethod
  def create(cls, *args, **kwargs):
    """ Factory method to create a new Service object. """
    if kwargs.get('ServiceName') in cls.protected_services:
      return ProtectedService(*args, **kwargs)
    else:
      return Service(*args, **kwargs)

  def __init__(self, *args, **kwargs):
    """ Constructor for Service object. """
    self.ServiceName             = kwargs.get('ServiceName')
    self.DisplayName             = kwargs.get('DisplayName')
    self.ServiceType             = kwargs.get('ServiceType')
    self.CurrentState            = kwargs.get('CurrentState')
    self.ControlsAccepted        = kwargs.get('ControlsAccepted')
    self.Win32ExitCode           = kwargs.get('Win32ExitCode')
    self.ServiceSpecificExitCode = kwargs.get('ServiceSpecificExitCode')
    self.CheckPoint              = kwargs.get('CheckPoint')
    self.WaitHint                = kwargs.get('WaitHint')
    self.ProcessId               = kwargs.get('ProcessId')
    self.ServiceFlags            = kwargs.get('ServiceFlags')
    self.query_service()
    self.last_error = ""

  def query_service(self):
    """
    Query additional information about the service and store them into the
    object.
    """
    hService = winsvc.OpenService(self.scm_handle, self.ServiceName,
      winsvc.SERVICE_QUERY_CONFIG)
    lpServiceConfig = winsvc.QueryServiceConfig(hService)
    winsvc.CloseServiceHandle(hService)
    self.ServiceType      = lpServiceConfig[0]
    self.StartType        = lpServiceConfig[1]
    self.ErrorControl     = lpServiceConfig[2]
    self.BinaryPathName   = lpServiceConfig[3]
    self.LoadOrderGroup   = lpServiceConfig[4]
    self.TagId            = lpServiceConfig[5]
    self.Dependencies     = lpServiceConfig[6]
    self.ServiceStartName = lpServiceConfig[7]
    self.DisplayName      = lpServiceConfig[8]

  def start(self):
    """ Start the service. """
    hService = winsvc.OpenService(self.scm_handle, self.ServiceName,
      winsvc.SERVICE_START)
    try:
      winsvc.StartService(hService, None)
    except pywintypes.error, e:
      self.last_error = e.strerror
    winsvc.CloseServiceHandle(hService)

  def stop(self):
    """ Stop the service. """
    hService = winsvc.OpenService(self.scm_handle, self.ServiceName,
      winsvc.SERVICE_STOP)
    try:
      winsvc.ControlService(hService, winsvc.SERVICE_CONTROL_STOP)
    except pywintypes.error, e:
      if e.winerror == winerror.ERROR_DEPENDENT_SERVICES_RUNNING:
        self.last_error = "%s (%s)" % (e.strerror, ','.join(self.Dependencies))
      else:
        self.last_error = e.strerror
    winsvc.CloseServiceHandle(hService)

  def set_start_type(self, start_type):
    """ Change the startup type of the service. """
    hService = winsvc.OpenService(self.scm_handle, self.ServiceName,
      winsvc.SERVICE_CHANGE_CONFIG)
    try:
      winsvc.ChangeServiceConfig(hService, winsvc.SERVICE_NO_CHANGE, start_type,
        winsvc.SERVICE_NO_CHANGE, None, None, False, None, None, None, None)
    except pywintypes.error, e:
      self.last_error = e.strerror
    winsvc.CloseServiceHandle(hService)

class ProtectedService(Service):
  """ Describes a protected service which cannot be manipulated. """

  """
  The error message displayed when the user attempts to manipulate the service.
  """
  error_msg = 'The service is configured as protected.'

  def start(self):
    """ nop """
    self.last_error = self.error_msg

  def stop(self):
    """ nop """
    self.last_error = self.error_msg

  def set_start_type(self, start_type):
    """ nop """
    self.last_error = self.error_msg

class ServiceFormatter():
  """ Provides formatting for service's fields. """

  def fmt_state(self, state):
    """ Return a string describing the service's state. """
    if state == winsvc.SERVICE_STOPPED:
      return 'Stopped'
    elif state == winsvc.SERVICE_START_PENDING:
      return 'Starting...'
    elif state == winsvc.SERVICE_STOP_PENDING:
      return 'Stopping...'
    elif state == winsvc.SERVICE_RUNNING:
      return 'Running'
    elif state == winsvc.SERVICE_CONTINUE_PENDING:
      return 'Continuing...'
    elif state == winsvc.SERVICE_PAUSE_PENDING:
      return 'Pausing...'
    elif state == winsvc.SERVICE_PAUSED:
      return 'Paused'
    else:
      raise Exception('Unknown service state: %d' % state)

  def fmt_accept(self, accept):
    """ Return a string describing the accepted controls of the service. """
    if accept & winsvc.SERVICE_CONTROL_STOP:
      return 'Stop'
    else:
      return '(None)'

  def fmt_start(self, start):
    """ Return a string describing the startup type of the service. """
    if start == winsvc.SERVICE_BOOT_START:
      return 'Boot'
    elif start == winsvc.SERVICE_SYSTEM_START:
      return 'System'
    elif start == winsvc.SERVICE_AUTO_START:
      return 'Auto'
    elif start == winsvc.SERVICE_DEMAND_START:
      return 'Demand'
    elif start == winsvc.SERVICE_DISABLED:
      return 'Disabled'
    else:
      raise Exception('Unknown service start type: %d' % start)

  def color(self, service):
    """ Return a color for the service, depending on its state. """
    if service.CurrentState == winsvc.SERVICE_RUNNING:
      return '#88FF77'
    elif service.CurrentState == winsvc.SERVICE_STOPPED:
      return '#FF5555'
    else:
      return '#FFDD22'

  def image(self, service):
    """
    Return an image for the service, depending on whether it's protected or not.
    """
    if isinstance(service, ProtectedService):
      return 'lock'
    else:
      return None

class ServiceListCtrl(wx.ListCtrl):
  """ Provides an enhanced ListCtrl for displaying services. """

  def __init__(self, *args, **kwargs):
    """ Construct the list control. """
    kwargs['style'] = wx.LC_VIRTUAL | wx.LC_REPORT
    self.popup_menu = kwargs.pop('popup_menu')
    wx.ListCtrl.__init__(self, *args, **kwargs)

    self.InsertColumn(0, "Service name"     , width = 150)
    self.InsertColumn(1, "Current state"    , width = 100)
    self.InsertColumn(2, "Controls accepted", width = 100)
    self.InsertColumn(3, "Start type"       , width = 100)
    self.InsertColumn(4, "Binary path name" , width = 300)
    self.InsertColumn(5, "Last Error"       , width = 200)

    wx.EVT_LIST_ITEM_RIGHT_CLICK(self, -1, self.on_context_menu)

    self.image_list = wx.ImageList(16,16)
    self.image_list.Add(wx.Bitmap('lock.png', wx.BITMAP_TYPE_PNG))
    self.SetImageList(self.image_list, wx.IMAGE_LIST_SMALL)
    self.image_list_dict = {
      'lock': 0,
    }

  def OnGetItemText(self, item, column):
    """
    Return the text displayed for each row and column of the list control.
    """
    if column == 0:
      return self.services[item].ServiceName
    elif column == 1:
      return ServiceFormatter().fmt_state(self.services[item].CurrentState)
    elif column == 2:
      return ServiceFormatter().fmt_accept(self.services[item].ControlsAccepted)
    elif column == 3:
      return ServiceFormatter().fmt_start(self.services[item].StartType)
    elif column == 4:
      return self.services[item].BinaryPathName
    elif column == 5:
      return self.services[item].last_error

  def OnGetItemImage(self, item):
    """ Return the image displayed for each row of the list control. """
    return self.image_list_dict.get(
        ServiceFormatter().image(self.services[item]), -1)

  def OnGetItemAttr(self, item):
    """ Set the color for each row of the list control. """
    self.attr = wx.ListItemAttr()
    self.attr.SetBackgroundColour(ServiceFormatter().color(self.services[item]))
    return self.attr

  def set_services(self, services):
    """ Set the services displayed in the list control. """
    self.services = services
    self.SetItemCount(len(self.services))
    self.Refresh()

  def get_selected_services(self):
    """ Return the selected services of the list control. """
    item = -1
    while True:
      item = self.GetNextItem(item, wx.LIST_NEXT_ALL, wx.LIST_STATE_SELECTED)
      if item == -1:
        break
      yield self.services[item]

  def start(self):
    """ Start the selected services. """
    for service in self.get_selected_services():
      service.start()
    self.Refresh()

  def stop(self):
    """ Stop the selected services. """
    for service in self.get_selected_services():
      service.stop()
    self.Refresh()

  def set_start_type(self, start_type):
    """ Change the startup type of the selected services. """
    for service in self.get_selected_services():
      service.set_start_type(start_type)
    self.Refresh()

  def on_context_menu(self, event):
    """ Handle right click on the list control and display a popup menu. """
    self.PopupMenu(self.popup_menu)

class MainFrame(wx.Frame):
  """ Describes the main frame of the application. """

  def __init__(self, *args, **kwargs):
    """ Construct the main frame. """
    wx.Frame.__init__(self, *args, **kwargs)

    ID_REFRESH  = wx.NewId()
    ID_ABOUT    = wx.NewId()
    ID_EXIT     = wx.NewId()
    ID_START    = wx.NewId()
    ID_STOP     = wx.NewId()
    ID_BOOT     = wx.NewId()
    ID_SYSTEM   = wx.NewId()
    ID_AUTO     = wx.NewId()
    ID_DEMAND   = wx.NewId()
    ID_DISABLED = wx.NewId()

    self.CreateStatusBar()

    menu_bar = wx.MenuBar()

    menu = wx.Menu()
    menu.Append(ID_REFRESH, "Refresh\tF5", "Refresh services")
    menu.AppendSeparator()
    menu.Append(ID_EXIT, "Exit", "Terminate the program")
    menu_bar.Append(menu, "File");

    service_menu = wx.Menu()
    service_menu.Append(ID_START, "Start", "Start the selected services")
    service_menu.Append(ID_STOP, "Stop", "Stop the selected services")
    service_menu.AppendSeparator()
    service_menu.Append(ID_BOOT, "Boot",
        "Set start type for the selected services")
    service_menu.Append(ID_SYSTEM, "System",
        "Set start type for the selected services")
    service_menu.Append(ID_AUTO, "Auto",
        "Set start type for the selected services")
    service_menu.Append(ID_DEMAND, "Demand",
        "Set start type for the selected services")
    service_menu.Append(ID_DISABLED, "Disabled",
        "Disable the selected services")
    menu_bar.Append(service_menu, "Service");

    menu = wx.Menu()
    menu.Append(ID_ABOUT, "About", "More information about this program")
    menu_bar.Append(menu, "Help");

    self.SetMenuBar(menu_bar)

    self.notebook = wx.Notebook(self)
    self.drivers_panel = wx.Panel(self.notebook)
    self.drivers_listctrl = ServiceListCtrl(self.drivers_panel,
        popup_menu = service_menu)
    sizer = wx.BoxSizer()
    sizer.Add(self.drivers_listctrl, 1, wx.EXPAND)
    self.drivers_panel.SetSizer(sizer)
    self.services_panel = wx.Panel(self.notebook)
    self.services_listctrl = ServiceListCtrl(self.services_panel,
        popup_menu = service_menu)
    sizer = wx.BoxSizer()
    sizer.Add(self.services_listctrl, 1, wx.EXPAND)
    self.services_panel.SetSizer(sizer)
    self.notebook.AddPage(self.services_panel, "User-mode services")
    self.notebook.AddPage(self.drivers_panel, "Kernel drivers")

    wx.EVT_MENU(self, ID_REFRESH , self.on_refresh )
    wx.EVT_MENU(self, ID_ABOUT   , self.on_about   )
    wx.EVT_MENU(self, ID_EXIT    , self.on_exit    )
    wx.EVT_MENU(self, ID_START   , self.on_start   )
    wx.EVT_MENU(self, ID_STOP    , self.on_stop    )
    wx.EVT_MENU(self, ID_BOOT    , self.on_boot    )
    wx.EVT_MENU(self, ID_SYSTEM  , self.on_system  )
    wx.EVT_MENU(self, ID_AUTO    , self.on_auto    )
    wx.EVT_MENU(self, ID_DEMAND  , self.on_demand  )
    wx.EVT_MENU(self, ID_DISABLED, self.on_disabled)

    self.on_refresh(None)

  def get_active_listctrl(self):
    """ Return the active list control depending on the current tab. """
    cur_page = self.notebook.GetCurrentPage()
    if cur_page == self.services_panel:
      return self.services_listctrl
    elif cur_page == self.drivers_panel:
      return self.drivers_listctrl

  def on_refresh(self, event):
    """ Handle refresh events (F5) requested by the user. """
    sort_keys = ['CurrentState', 'ControlsAccepted', 'StartType']
    services = Service.get_all()
    drivers = Service.filter(services, ServiceType = winsvc.SERVICE_DRIVER)
    drivers = Service.sort(drivers, *sort_keys)
    user_services = Service.filter(services, ServiceType = winsvc.SERVICE_WIN32)
    user_services = Service.sort(user_services, *sort_keys)
    self.drivers_listctrl.set_services(drivers)
    self.services_listctrl.set_services(user_services)

  def on_about(self, event):
    """ nop """
    pass

  def on_exit(self, event):
    """ Close the application on exit event. """
    self.Close(True)

  def on_start(self, event):
    """ Start services. """
    self.get_active_listctrl().start()

  def on_stop(self, event):
    """ Stop services. """
    self.get_active_listctrl().stop()

  def on_boot(self, event):
    """ Change the startup type of services. """
    self.get_active_listctrl().set_start_type(winsvc.SERVICE_BOOT_START)

  def on_system(self, event):
    """ Change the startup type of services. """
    self.get_active_listctrl().set_start_type(winsvc.SERVICE_SYSTEM_START)

  def on_auto(self, event):
    """ Change the startup type of services. """
    self.get_active_listctrl().set_start_type(winsvc.SERVICE_AUTO_START)

  def on_demand(self, event):
    """ Change the startup type of services. """
    self.get_active_listctrl().set_start_type(winsvc.SERVICE_DEMAND_START)

  def on_disabled(self, event):
    """ Change the startup type of services. """
    self.get_active_listctrl().set_start_type(winsvc.SERVICE_DISABLED)

class App(wx.App):
  """ The main application. """

  def OnInit(self):
    """ Initialize the application and display the main frame. """
    frame = MainFrame(None, -1, "Windows Service Manager")
    frame.Maximize()
    self.SetTopWindow(frame)
    frame.Show(True)
    return True

def main():
  app = App(0)
  app.MainLoop()

if __name__ == "__main__":
  main()
