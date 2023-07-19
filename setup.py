from cx_Freeze import setup, Executable
executables = [Executable('CRMandAccount.py', base='Win32GUI',
                          target_name='CRMandAccount.exe',
                          icon='ico/analysis_finance_statistics_business_graph_chart_report_icon_254045.ico')]
excludes = ['unittest', 'asyncio', 'sqlite3', 'distutils']
includefiles = ['__PROJECTS__.txt', '__VERS__.txt']


zip_include_packages = ['PySimpleGUI','altgraph', 'asgiref','charset-normalizer',
                        'et-xmlfile','idna','importlib-metadata', 'numpy',
                        'openpyxl','pandas', 'pefile','pip','python-dateutil','pytils', 'pytz',
                        'pywin32','pywin32-ctypes', 'requests', 'setuptools', 'sgtpyutils', 'six','typing-extensions','urllib3',
                        'collections', 'tkinter', 'json', 'dateutil', 'encodings', 'html', 'http', 'importlib', 'multiprocessing',
                        'pywin', 'tcl8', 'tcl8.6', 'tkz8.6', 'urllib', 'win32com', 'classes', 'interfaces', 'ctypes', 'xml',
                        'et_xmlfile', 'email', 'concurrent','xmlrpc', 'test',
                        'lib2to3', 'curses', 'pkg_resources', 'pydoc_data']
options = {
      'build_exe': {
          'include_files': includefiles,
            'excludes': excludes,
            'build_exe': 'build_windows',
            # 'zip_include_packages': zip_include_packages,
            "zip_include_packages": "*",
            "zip_exclude_packages": "",
            'optimize': 1
      }
}


setup(name='CRMandAccount',
      version='1.0.1',
      description='Сверка',
      executables=executables,
      options=options)