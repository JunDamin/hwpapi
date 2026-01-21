# PyInstaller hook for hwpapi
# Automatically bundles FilePathCheckerModuleExample.dll

from PyInstaller.utils.hooks import collect_data_files

# Collect all data files (including .dll) from hwpapi package
datas = collect_data_files('hwpapi', includes=['*.dll'])
