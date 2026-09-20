"""Styles et helpers partagés par les générateurs de classeurs Apnée Attitude."""
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

FONT = "Arial"
H_FILL = PatternFill("solid", fgColor="1F4E5F")
SUB_FILL = PatternFill("solid", fgColor="D9E7EC")
CALC_FILL = PatternFill("solid", fgColor="EDEDED")   # gris : calculé par formule locale
TODO_FILL = PatternFill("solid", fgColor="FFF2CC")
SEL_FILL = PatternFill("solid", fgColor="E2EFDA")   # vert clair : aide à la saisie
WHITE_FILL = PatternFill("solid", fgColor="FFFFFF")  # blanc : saisie libre
# Bleu pâle : copié automatiquement depuis le Paramétrage à chaque
# synchronisation — avant le 06/09/2026 ce style partageait CALC_FILL,
# indiscernable à l'écran d'une cellule calculée localement (constat sur
# l'onglet Personnes, qui mélange les deux).
LINK_FILL = PatternFill("solid", fgColor="DDEBF7")
H_FONT = Font(name=FONT, size=10, bold=True, color="FFFFFF")
SUB_FONT = Font(name=FONT, size=9, bold=True, color="1F4E5F")
TITLE_FONT = Font(name=FONT, size=14, bold=True, color="1F4E5F")
BIG_FONT = Font(name=FONT, size=16, bold=True, color="1F4E5F")
SEC_FONT = Font(name=FONT, size=12, bold=True, color="1F4E5F")
BODY = Font(name=FONT, size=10)
INPUT_FONT = Font(name=FONT, size=10)
CALC_FONT = Font(name=FONT, size=10, color="000000")
LINK_FONT = Font(name=FONT, size=10, italic=True, color="31859B")
NOTE_FONT = Font(name=FONT, size=9, italic=True, color="808080")
EX_FONT = Font(name=FONT, size=10, bold=True, color="C00000")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ---- domaines de valeurs communs aux deux classeurs
NIVEAUX_COMP = ["non acquis", "en cours", "acquis", "expert"]
STATUT_ACQ = ["non tenté", "non réussi", "réussi"]
ZONE_CONFORT = ["Confort", "Challenge", "Limite"]
MOTIFS_ABS = ["non renseigné", "congé", "professionnel", "blessure", "maladie",
              "autre entraînement", "compétition", "autre"]
MOTIFS_INDISPO = ["congé", "professionnel", "blessure", "maladie", "formation",
                  "compétition", "autre"]
NATURES = ["ordinaire", "technique", "hypercapnique", "mixte", "hypoxique",
           "test maximal", "RIFAA", "PSM", "examen", "formation", "tutorat",
           "multi-lignes"]
STATUTS_SEANCE = ["planifiée", "tenue", "annulée", "fermée"]
AVIS = ["absentéiste", "dernier tiers", "dans la moyenne", "premier tiers",
        "prêt pour le niveau supérieur"]
TYPES_EVT = ["perte de contrôle moteur (PCM)", "samba", "syncope",
             "arrêt de séance", "malaise", "barotraumatisme", "autre"]
GRAVITES = ["mineur", "significatif", "grave"]
JOURS = ["lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"]
ROLES_ENC = ["directeur de bassin", "responsable de ligne", "co-encadrant",
             "tuteur", "stagiaire encadré"]
ROLES_PERS = ["élève", "encadrant", "élève et encadrant"]
SUPPORTS = ["classeur", "application"]
PERIODICITES = ["toutes les semaines", "semaines paires", "semaines impaires",
                "ponctuel"]
GENRES = ["F", "H", "non renseigné"]
STATUTS_PERS = ["actif", "inactif", "invité", "en attente de validation"]
TYPES_OBJ = ["maîtrise", "acquisition"]
CONTRAINTES = ["hypercapnique et/ou hypoxique", "meilleures conditions",
               "aucune contrainte particulière"]

VERSION = "1.1"
DATE_VERSION = "2026-08-20"

# ---- couleurs d'onglet : bleu pour les onglets de saisie/consultation,
# gris pour les onglets techniques (jamais ouverts par un encadrant)
TAB_SAISIE = "1F4E5F"
TAB_TECHNIQUE = "A6A6A6"


def sheet(wb, name, title, intro=None, tab_color=TAB_SAISIE):
    ws = wb.create_sheet(name)
    if tab_color:
        ws.sheet_properties.tabColor = tab_color
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    if intro:
        ws["A2"] = intro
        ws["A2"].font = NOTE_FONT
        ws["A2"].alignment = Alignment(wrap_text=False)
        return ws, 4
    return ws, 3


def section(ws, row, titre, note, cols):
    """Bandeau de titre et en-têtes de colonnes d'une sous-table, dans un
    onglet qui en regroupe plusieurs les unes sous les autres."""
    ws.cell(row, 1, titre).font = SEC_FONT
    row += 1
    if note:
        c = ws.cell(row, 1, note)
        c.font = NOTE_FONT
        c.alignment = Alignment(wrap_text=False)
        row += 1
    return header(ws, row, cols)


def make_table(ws, name, header_row, last_row, ncols):
    """Transforme une plage en Table Excel structurée nommée : lignes à
    filtrer/trier, mise en forme et formules recopiées automatiquement dès
    qu'une ligne est ajoutée juste en dessous."""
    ref = "A%d:%s%d" % (header_row, get_column_letter(ncols), last_row)
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False,
                                       showFirstColumn=False, showLastColumn=False,
                                       showColumnStripes=False)
    ws.add_table(t)
    return t


def header(ws, row, cols, height=32):
    for i, (label, width) in enumerate(cols, start=1):
        c = ws.cell(row, i, label)
        c.fill, c.font, c.border = H_FILL, H_FONT, BOX
        c.alignment = Alignment(vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[row].height = height
    return row + 1


def dv_list(ws, values, cellrange):
    """Liste déroulante à valeurs littérales (courtes)."""
    d = DataValidation(type="list", formula1='"%s"' % ",".join(values),
                       allow_blank=True, showDropDown=False)
    ws.add_data_validation(d)
    d.add(cellrange)


def dv_range(ws, ref, cellrange, msg=None):
    """Liste déroulante alimentée par une plage (listes longues)."""
    d = DataValidation(type="list", formula1=ref, allow_blank=True,
                       showDropDown=False)
    if msg:
        d.promptTitle, d.prompt, d.showInputMessage = "Sélection", msg, True
    ws.add_data_validation(d)
    d.add(cellrange)


def style_cell(c, kind="input", wrap=False, fmt=None):
    if kind == "input":
        c.font, c.fill = INPUT_FONT, WHITE_FILL
    elif kind == "select":
        c.font, c.fill = INPUT_FONT, SEL_FILL
    elif kind == "calc":
        c.font, c.fill = CALC_FONT, CALC_FILL
    elif kind == "link":
        c.font, c.fill = LINK_FONT, LINK_FILL
    elif kind == "body":
        c.font, c.fill = BODY, WHITE_FILL
    c.border = BOX
    c.alignment = Alignment(vertical="top", wrap_text=wrap)
    if fmt:
        c.number_format = fmt
    return c


def id_from_two_labels(col_nom, col_pre, row):
    """Extrait l'identifiant entre crochets, depuis l'une ou l'autre sélection."""
    a, b = "$%s%d" % (col_nom, row), "$%s%d" % (col_pre, row)
    piece = ('IFERROR(MID({c},FIND("[",{c})+1,FIND("]",{c})-FIND("[",{c})-1),"")')
    return "=IF({a}<>\"\",{pa},IF({b}<>\"\",{pb},\"\"))".format(
        a=a, b=b, pa=piece.format(c=a), pb=piece.format(c=b))


def lookup(target_range, key_cell, key_range):
    # « &"" » évite qu'une cellule source vide ressorte en 0
    return '=IFERROR(INDEX(%s,MATCH(%s,%s,0))&"","")' % (
        target_range, key_cell, key_range)


def write_readme(ws, titre, blocs):
    ws["A1"] = titre
    ws["A1"].font = BIG_FONT
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 118
    r = 3
    for kind, txt in blocs:
        if kind == "titre":
            ws.cell(r, 2, txt).font = SEC_FONT
        elif kind:
            c = ws.cell(r, 2, txt)
            c.font = BODY
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 14 * (1 + len(txt) // 105)
        r += 1
    return r


# ------------------------------------------------------------------ saisie
def resolve_roster(inp, code_rng, id_rng, nom_rng, pre_rng):
    """Identifiant depuis un code numérique court OU un début de nom/prénom."""
    return ('=IF({i}="","",IFERROR(IF(ISNUMBER({i}),'
            'INDEX({id},MATCH({i},{code},0)),'
            'IFERROR(INDEX({id},MATCH({i}&"*",{nom},0)),'
            'INDEX({id},MATCH({i}&"*",{pre},0)))),""))'
            ).format(i=inp, code=code_rng, id=id_rng, nom=nom_rng, pre=pre_rng)


def resolve_club(inp, id_rng, nom_rng, pre_rng):
    """Identifiant depuis un début de nom ou de prénom, sur tout le club."""
    return ('=IF({i}="","",IFERROR(IFERROR('
            'INDEX({id},MATCH({i}&"*",{nom},0)),'
            'INDEX({id},MATCH({i}&"*",{pre},0))),""))'
            ).format(i=inp, id=id_rng, nom=nom_rng, pre=pre_rng)


def count_roster(inp, code_rng, nom_rng, pre_rng):
    return ('IF(ISNUMBER({i}),COUNTIF({code},{i}),'
            'COUNTIF({nom},{i}&"*")+COUNTIF({pre},{i}&"*"))'
            ).format(i=inp, code=code_rng, nom=nom_rng, pre=pre_rng)


def count_club(inp, nom_rng, pre_rng):
    return ('COUNTIF({nom},{i}&"*")+COUNTIF({pre},{i}&"*")'
            ).format(i=inp, nom=nom_rng, pre=pre_rng)
