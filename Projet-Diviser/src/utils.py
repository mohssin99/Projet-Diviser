from openpyxl.styles import Font, Alignment, Border, Side

def configurer_impression(ws):
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.75
    ws.page_margins.bottom = 0.75
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.print_options.horizontalCentered = True

def configurer_styles():
    return {
        'entete_font': Font(bold=True),
        'entete_alignment': Alignment(horizontal='center', vertical='center'),
        'cell_alignment': Alignment(horizontal='left', vertical='center'),
        'border': Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
    }