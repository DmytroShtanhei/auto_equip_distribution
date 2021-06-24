"""
Styles for cells
"""
from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill

# Header Named Style:
header_style = NamedStyle(name="header_style")
header_style.font = Font(bold=True, size=11)
header_style.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
bd_thin = Side(style='thin', color="000000")
bd_thick = Side(style='thick', color="000000")
header_style.border = Border(left=bd_thin, top=bd_thin, right=bd_thin, bottom=bd_thin)
header_style.fill = PatternFill(fill_type='solid', start_color='00C0C0C0')


# Data Named Style:
data_style = NamedStyle(name="cell_style")
data_style.font = Font(bold=False, size=11)
data_style.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False, indent=0)
bd_thin = Side(style='thin', color="000000")
data_style.border = Border(left=bd_thin, top=bd_thin, right=bd_thin, bottom=bd_thin)
