# AUTOGENERATED! DO NOT EDIT! File to edit: ../nbs/02_api/02_functions.ipynb.

# %% auto 0
__all__ = ['dispatch', 'get_absolute_path', 'get_dll_path', 'add_dll_to_registry', 'get_registry_value', 'check_dll', 'get_value',
           'set_charshape_pset', 'set_parashape_pset', 'hex_to_rgb', 'get_rgb_tuple']

# %% ../nbs/02_api/02_functions.ipynb 3
import os
import shutil
from pathlib import Path
import winreg
import importlib.resources

# %% ../nbs/02_api/02_functions.ipynb 4
def dispatch(app_name:str):
    """캐시가 충돌하는 문제를 해결하기 위해 실행합니다. 에러가 발생할 경우 기존 캐시를 삭제하고 다시 불러옵니다."""
    try:
        from win32com import client
        app = client.gencache.EnsureDispatch(app_name)
    except AttributeError:
        # Corner case dependencies.
        import os
        import re
        import sys
        import shutil
        # Remove cache and try again.
        MODULE_LIST = [m.__name__ for m in sys.modules.values()]
        for module in MODULE_LIST:
            if re.match(r'win32com\.gen_py\..+', module):
                del sys.modules[module]
        shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
        from win32com import client
        app = client.gencache.EnsureDispatch(app_name)
    return app

# %% ../nbs/02_api/02_functions.ipynb 5
def get_absolute_path(path):
    """파일 절대 경로를 반환합니다."""
    name = Path(path)
    return name.absolute().as_posix()

# %% ../nbs/02_api/02_functions.ipynb 6
def get_dll_path(package_name, dll_filename):
    """패키지에서 dll 경로를 확보합니다."""
    try:
        with importlib.resources.path(package_name, dll_filename) as dll_path:
            return str(dll_path)
    except FileNotFoundError as e:
        raise FileNotFoundError(f"The DLL file '{dll_filename}' was not found in the package '{package_name}'.") from e


# %% ../nbs/02_api/02_functions.ipynb 7
def add_dll_to_registry(dll_path, key_path):
    """레지스트리에 dll을 등록합니다."""
    try:
        # Connect to the registry and open the specified key
        registry_key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)

        # Set the value for the new registry entry as a string (REG_SZ)
        winreg.SetValueEx(registry_key, "FilePathCheckerModule", 0, winreg.REG_SZ, dll_path)

        # Close the registry key
        winreg.CloseKey(registry_key)
        print("DLL path added to registry as a string value successfully.")
    except WindowsError as e:
        print("Error while adding DLL path to registry: ", e)
        

# %% ../nbs/02_api/02_functions.ipynb 8
def get_registry_value(key_path, value_name):
    """레지스트리에 값이 있는지 확인해 봅니다."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            value, value_type = winreg.QueryValueEx(key, value_name)
            return value
    except FileNotFoundError:
        return None

# %% ../nbs/02_api/02_functions.ipynb 9
def check_dll():
    """dll 모듈을 등록합니다."""
    dll_path = get_dll_path("hwpapi", "FilePathCheckerModuleExample.dll")
    key_path = "SOFTWARE\\HNC\\HwpAutomation\\Modules"
    value_name = "FilePathCheckerModule" 

    value = get_registry_value(key_path, value_name)

    if value is None:
        add_dll_to_registry(dll_path, key_path)
    return True


# %% ../nbs/02_api/02_functions.ipynb 10
def get_value(dict_, key):
    """딕셔너리에서 키를 찾아 값을 반환합니다. 반환할 값이 없으면 키에러와 함께 가능한 키를 알려줍니다."""
    try:
        return dict_[key]
    except KeyError:
        raise KeyError(f"{key}를 해당하는 키 중 찾을 수 없습니다. 키는 {', '.join(dict_.keys())} 중에 있어야 합니다.")

# %% ../nbs/02_api/02_functions.ipynb 11
def set_charshape_pset(
    charshape, 
    face_name:str=None, 
    font_type:int=None, 
    size:int=None,
    ratio:int=None, 
    spacing:int=None, 
    offset:int=None,
    bold:bool=None, 
    italic:bool=None, 
    small_caps:bool=None,
    emboss:bool=None,
    super_script:bool=None,
    sub_script:bool=None,
    underline_type:int=None,
    outline_type:int=None,
    text_color=None,
    shade_color=None,
    underline_shape:int=None,
    underline_color=None,
    shadow_offset_x:int=None,
    shadow_offset_y:int=None,
    shadow_color=None,
    strike_out_type=None,
    diac_sym_mark=None,
    use_font_space=None,
    use_kerning=None,
    height:int=None,
):
    """
    CharShape값을 입력하기 위한 함수입니다.
    `BorderFill`은 별도로 "BorderFill"타입을 사용하기 때문에 제외하였습니다. 
    """
    params = []
    categories = ["Hangul", "Latin", "Hanja", "Japanese", "Other", "Symbol", "User"]
    
    params += [("FaceName"+cat, face_name) for cat in categories] if face_name else []
    params += [("FontType"+cat, font_type) for cat in categories] if face_name else []
    params += [("Size"+cat, size) for cat in categories] if size else []
    params += [("Ratio"+cat, ratio) for cat in categories] if ratio else []
    params += [("Spacing"+cat, spacing) for cat in categories] if spacing else []
    params += [("Offset"+cat, offset) for cat in categories] if offset else []
    
    params += list(
        filter(lambda x: x[1] is not None, 
            [
                ("Bold", bold),
                ("Italic", italic),
                ("SmallCaps", small_caps),
                ("Emboss", emboss),
                ("SuperScript", super_script),
                ("SubScript", sub_script),
                ("UnderlineType", underline_type),
                ("OutlineType", outline_type),
                ("TextColor", text_color),
                ("ShadeColor", shade_color),
                ("UnderlineShape", underline_shape),
                ("UnderlineColor", underline_color),
                ("ShadowOffsetX", shadow_offset_x),
                ("ShadowOffsetY", shadow_offset_y),
                ("ShadowColor", shadow_color),
                ("StrikeOutType", strike_out_type),
                ("DiacSymMark", diac_sym_mark),
                ("UseFontSpace", use_font_space),
                ("UseKerning", use_kerning),
                ("Height", height),
            ]
        )
    ) 
    for key, value in params:
            setattr(charshape, key, value)
    
    return charshape


# %% ../nbs/02_api/02_functions.ipynb 12
def set_parashape_pset(
    parashape,
    left_margin:int=None,
    right_margin:int=None,
    indentation:int=None,
    prev_spacing:int=None,
    next_spacing:int=None,
    line_spacing_type:int=None,
    line_spacing:int=None,
    align_type:int=None,
    break_latin_word:int=None,
    break_non_latin_word:bool=None,
    snap_to_grid:bool=None,
    condense:float=None,
    widow_orphan:bool=None,
    keep_with_next:bool=None,
    page_break_before:bool=None,
    text_alignment:int=None,
    font_line_height:int=None,
    heading_type:int=None,
    level:int=None,
    border_connect:bool=None,
    border_text:bool=None,
    border_offset_left:int=None,
    border_offset_right:int=None,
    border_offset_top:int=None,
    border_offset_bottom:int=None,
    tail_type:bool=None,
    line_wrap:bool=None,
):
    params = list(
        filter(lambda x: x[1] is not None, 
            [
                ("LeftMargin", left_margin),
                ("RightMargin", right_margin),
                ("Indentation", indentation),
                ("PrevSpacing", prev_spacing),
                ("NextSpacing", next_spacing),
                ("LineSpacingType", line_spacing_type),
                ("LineSpacing", line_spacing),
                ("AlignType", align_type),
                ("BreakLatinWord", break_latin_word),
                ("BreakNonLatinWord", break_non_latin_word),
                ("SnapToGrid", snap_to_grid),
                ("Condense", condense),
                ("WidowOrphan", widow_orphan),
                ("KeepWithNext", keep_with_next),
                ("PageBreakBefore", page_break_before),
                ("TextAlignment", text_alignment),
                ("FontLineHeight", font_line_height),
                ("HeadingType", heading_type),
                ("Level", level),
                ("BorderConnect", border_connect),
                ("BorderText", border_text),
                ("BorderOffsetLeft", border_offset_left),
                ("BorderOffsetRight", border_offset_right),
                ("BorderOffsetTop", border_offset_top),
                ("BorderOffsetBottom", border_offset_bottom),
                ("TailType", tail_type),
                ("LineWrap", line_wrap),
            ]
        )
    ) 
    for key, value in params:
        setattr(parashape, key, value)
    
    return parashape

# %% ../nbs/02_api/02_functions.ipynb 13
def hex_to_rgb(hex_string):
    # Remove the "#" symbol if it exists
    if hex_string.startswith("#"):
        hex_string = hex_string[1:]

    # Convert the hex string to decimal integers
    red = int(hex_string[0:2], 16)
    green = int(hex_string[2:4], 16)
    blue = int(hex_string[4:], 16)

    # Return the RGB tuple
    return (red, green, blue)

def get_rgb_tuple(color):
    # check if the input is already a tuple
    if isinstance(color, tuple):
        
        # validate each color 
        if len(color) > 3:
            raise ValueError(f"colors should contains three compoents which represents RGB")
        
        for component in color:
            if component > 255:
                raise ValueError(f"number should be smaller than 255. the value is {color}")
        return color

    # if the input is a string, split it into a list of colors
    elif isinstance(color, str):
        
        colors = {"red": (255, 0, 0), "green": (0, 255, 0), "blue": (0, 0, 255), "black": (0, 0, 0), "white": (255, 255, 255)}
        
        if color in colors.keys():
            return colors.get(color)

        # validate each color
        if not (color.startswith('#') and len(color) == 7 and all(c in '0123456789abcdefABCDEF' for c in color[1:])):
            raise ValueError(f"'{color}' is not a valid hexadecimal color.")

        # convert the list to a tuple and return it
        return hex_to_rgb(color)

    else:
        raise TypeError("Input must be a string or a tuple of colors.")

