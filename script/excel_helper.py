from openpyxl.styles import PatternFill, Font, Border, Side

class Color_style():
    PALE_MINT = PatternFill(fill_type="solid", fgColor="EBF1DE") 
    PALE_PINK = PatternFill(fill_type="solid", fgColor="F2DCDB")
    LIGHT_BLUE = PatternFill(fill_type="solid",fgColor="DCE6F1")
    SKY_BLUE = PatternFill(fill_type="solid",fgColor="C5D9F1")
    PEACH = PatternFill(fill_type="solid", fgColor="FDE9D9") 
    LACENDER = PatternFill(fill_type="solid", fgColor="E4DFEC")
    BEIGE = PatternFill(fill_type="solid",fgColor="DDD9C4")
    LIGHT_CORAL = PatternFill(fill_type="solid",fgColor="FFBDBD")
    YELLOW_GREEN = PatternFill(fill_type="solid",fgColor="92D050")
    SATIN_GREEN = PatternFill(fill_type="solid", fgColor="C6E0B4") 

class Font_style():
    FONT = Font(name='Book Antiqua', size=12)

class Border_style():
    S_M = Side(border_style="medium")
    S_D = Side(border_style="dashed")
    S_T = Side(border_style="thin")

    border_dict = {
        "M" : S_M,
        "D" : S_D,
        "T" : S_T
    }

    @classmethod
    def set(cls,t=None,r=None,l=None,b=None): 
        return Border(
            top = cls.border_dict.get(t),         
            bottom = cls.border_dict.get(b),         
            left = cls.border_dict.get(l),         
            right = cls.border_dict.get(r)         
        )