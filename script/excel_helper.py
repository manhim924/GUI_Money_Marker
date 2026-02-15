from openpyxl.styles import PatternFill, Font, Border, Side
class Color():
    PALE_MINT = PatternFill(fill_type="solid", fgColor="EBF1DE") 
    PALE_PINK = PatternFill(fill_type="solid", fgColor="F2DCDB")
    LIGHT_BLUE = PatternFill(fill_type="solid",fgColor="DCE6F1")
    SKY_BLUE = PatternFill(fill_type="solid",fgColor="C5D9F1")
    PEACH = PatternFill(fill_type="solid", fgColor="FDE9D9") 
    LACENDER = PatternFill(fill_type="solid", fgColor="E4DFEC")
    BEIGE = PatternFill(fill_type="solid",fgColor="DDD9C4")
    LIGHT_CORAL = PatternFill(fill_type="solid",fgColor="FFBDBD")
    YELLOW_GREEN = PatternFill(fill_type="solid",fgColor="92D050")

class Font():
    FONT = Font(name='Book Antiqua', size=12)

class Border():
    STYLE_M = Side(border_style="medium")
    STYLE_D = Side(border_style="dashed")
    STYLE_T = Side(border_style="thin")
    
    F_M = Border(
        top = STYLE_M,
        bottom = STYLE_M, 
        left = STYLE_M,
        right = STYLE_M
    )
    TBL_M_R_D = Border(
        top = STYLE_M,
        bottom = STYLE_M,
        left = STYLE_M,
        right = STYLE_D
    )
    TBR_M_L_D = Border(
        top = STYLE_M,
        bottom = STYLE_M,
        left = STYLE_D,
        right = STYLE_M
    )    
    TB_M_LR_D = Border(
        top = STYLE_M,
        bottom = STYLE_M,
        left = STYLE_D,
        right = STYLE_D
    )
    TLR_M_B_D = Border(
        top = STYLE_M,
        bottom =STYLE_D,
        left = STYLE_M,
        right = STYLE_M
    )
    TB_D_LR_M = Border(
        top = STYLE_D,
        bottom =STYLE_D,
        left = STYLE_M,
        right = STYLE_M
    )
    T_D_LR_M_B_T = Border(
        top = STYLE_D,
        bottom =STYLE_T,
        left = STYLE_M,
        right = STYLE_M
    )
    TB_T_LR_M = Border(
        top = STYLE_T,
        bottom =STYLE_T,
        left = STYLE_M,
        right = STYLE_M
    )
    T_T_LRB_M = Border(
        top = STYLE_T,
        bottom =STYLE_M,
        left = STYLE_M,
        right = STYLE_M
    )
    
    TL_M_BR_D = Border(
        top = STYLE_M,
        bottom =STYLE_D,
        left = STYLE_M,
        right = STYLE_D
    )
    TRB_D_L_M = Border(
        top = STYLE_D,
        bottom =STYLE_D,
        left = STYLE_M,
        right = STYLE_D
    )
    TR_D_LB_M = Border(
        top = STYLE_D,
        bottom =STYLE_M,
        left = STYLE_M,
        right = STYLE_D
    )
    TB_T_L_M_R_D = Border(
        top = STYLE_T,
        bottom =STYLE_T,
        left = STYLE_M,
        right = STYLE_D
    )
    T_T_LB_M_R_D = Border(
        top = STYLE_T,
        bottom =STYLE_M,
        left = STYLE_M,
        right = STYLE_D
    )

    T_M_BLR_D = Border(
        top = STYLE_M,
        bottom =STYLE_D,
        left = STYLE_D,
        right = STYLE_D
    )
    F_D = Border(
        top = STYLE_D,
        bottom =STYLE_D,
        left = STYLE_D,
        right = STYLE_D
    )
    TLR_D_B_M = Border(
        top = STYLE_D,
        bottom =STYLE_M,
        left = STYLE_D,
        right = STYLE_D
    )
    TB_T_LR_D = Border(
        top = STYLE_T,
        bottom =STYLE_T,
        left = STYLE_D,
        right = STYLE_D
    )
    T_T_LR_D_B_M = Border(
        top = STYLE_T,
        bottom =STYLE_M,
        left = STYLE_D,
        right = STYLE_D
    )

    TR_M_BL_D = Border(
        top = STYLE_M,
        bottom =STYLE_D,
        left = STYLE_D,
        right = STYLE_M
    )
    TLB_D_R_M = Border(
        top = STYLE_D,
        bottom =STYLE_D,
        left = STYLE_D,
        right = STYLE_M
    )
    TL_D_RB_M = Border(
        top = STYLE_D,
        bottom =STYLE_M,
        left = STYLE_D,
        right = STYLE_M
    )
    TB_T_L_D_R_M = Border(
        top = STYLE_T,
        bottom =STYLE_T,
        left = STYLE_D,
        right = STYLE_M
    )
    T_T_L_D_BR_M = Border(
        top = STYLE_T,
        bottom =STYLE_M,
        left = STYLE_D,
        right = STYLE_M
    )