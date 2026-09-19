from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph
from reportlab.pdfgen import canvas


OUT = "guides/Guide-AppSuiviAA.pdf"
W, H = A4
M = 20 * mm
NAVY = colors.HexColor("#08242d")
TEAL = colors.HexColor("#0f7d78")
INK = colors.HexColor("#16202b")
PALE = colors.HexColor("#f3f7f6")
LINE = colors.HexColor("#d8e5e3")
AMBER_BG = colors.HexColor("#fff8ea")


def draw_paragraph(c, text, x, top, width, style):
    paragraph = Paragraph(text, style)
    _, height = paragraph.wrap(width, H)
    paragraph.drawOn(c, x, top - height)
    return height


def header(c, title, subtitle, page):
    c.setFillColor(NAVY)
    c.rect(0, H - 59 * mm, W, 59 * mm, fill=1, stroke=0)
    c.setFillColor(colors.HexColor("#b9c8ca"))
    c.setFont("Helvetica-Bold", 9.5)
    c.drawString(M, H - 18 * mm, "APNÉE ATTITUDE  |  SUIVI DES PRÉSENCES")
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 25)
    c.drawString(M, H - 32 * mm, title)
    subtitle_style = ParagraphStyle("subtitle", fontName="Helvetica", fontSize=11.5,
                                    leading=15, textColor=colors.HexColor("#d6e0e1"))
    draw_paragraph(c, subtitle, M, H - 38 * mm, W - 2 * M, subtitle_style)
    c.setFillColor(colors.HexColor("#90a2a7"))
    c.setFont("Helvetica", 8.5)
    c.drawRightString(W - M, 11 * mm, f"AppSuiviAA  |  saison 2026-2027  |  {page}/2")


def numbered_step(c, top, number, title, detail):
    x, width, height = M, W - 2 * M, 28 * mm
    c.setFillColor(PALE)
    c.setStrokeColor(LINE)
    c.roundRect(x, top - height, width, height, 5 * mm, fill=1, stroke=1)
    c.setFillColor(TEAL)
    c.circle(x + 12 * mm, top - 14 * mm, 6.5 * mm, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(x + 12 * mm, top - 16.5 * mm, str(number))
    title_style = ParagraphStyle("step_title", fontName="Helvetica-Bold", fontSize=13,
                                 leading=15, textColor=INK)
    detail_style = ParagraphStyle("step_detail", fontName="Helvetica", fontSize=10.5,
                                  leading=13.5, textColor=INK)
    draw_paragraph(c, title, x + 25 * mm, top - 6.5 * mm, width - 31 * mm, title_style)
    draw_paragraph(c, detail, x + 25 * mm, top - 14 * mm, width - 31 * mm, detail_style)
    return top - height - 4 * mm


def callout(c, top, title, text):
    height = 29 * mm
    c.setFillColor(AMBER_BG)
    c.setStrokeColor(colors.HexColor("#efd39a"))
    c.roundRect(M, top - height, W - 2 * M, height, 5 * mm, fill=1, stroke=1)
    style = ParagraphStyle("callout", fontName="Helvetica", fontSize=10.5,
                           leading=14, textColor=INK)
    draw_paragraph(c, f"<b><font color='#925d0c'>{title}</font></b><br/>{text}",
                   M + 7 * mm, top - 6 * mm, W - 2 * M - 14 * mm, style)


def section_card(c, x, top, width, title, text):
    height = 43 * mm
    c.setFillColor(PALE)
    c.setStrokeColor(LINE)
    c.roundRect(x, top - height, width, height, 5 * mm, fill=1, stroke=1)
    title_style = ParagraphStyle("card_title", fontName="Helvetica-Bold", fontSize=12,
                                 leading=15, textColor=TEAL)
    body_style = ParagraphStyle("card_body", fontName="Helvetica", fontSize=10,
                                leading=13.5, textColor=INK)
    draw_paragraph(c, title, x + 6 * mm, top - 7 * mm, width - 12 * mm, title_style)
    draw_paragraph(c, text, x + 6 * mm, top - 15 * mm, width - 12 * mm, body_style)


def main():
    c = canvas.Canvas(OUT, pagesize=A4)
    c.setTitle("Guide rapide AppSuiviAA")

    header(c, "Renseigner les présences", "Le guide rapide pour une saisie simple et fiable, depuis un téléphone ou un ordinateur.", 1)
    y = H - 69 * mm
    y = numbered_step(c, y, 1, "Ouvrez AppSuiviAA", "Utilisez le lien unique du club : apneeattitude.github.io/AppSuiviAA/. Le menu permet de choisir votre ligne.")
    y = numbered_step(c, y, 2, "Connectez-vous avec votre compte Google du club", "En cas de refus de connexion, contactez Fred.")
    y = numbered_step(c, y, 3, "Choisissez la séance", "Utilisez la liste déroulante en haut de page. La séance du jour est souvent déjà sélectionnée.")
    y = numbered_step(c, y, 4, "Ajoutez les participants", "Tapez les premières lettres du prénom ou du nom, puis touchez Ajouter.")
    y = numbered_step(c, y, 5, "Enregistrez", "Touchez Enregistrer la séance. Le message final confirme le nombre de présences enregistrées.")
    callout(c, y - 2 * mm, "À savoir", "Vous pouvez revenir sur une séance pour la modifier ou la compléter.")
    c.showPage()

    header(c, "Compléter les informations", "Les options utiles pour compléter ou corriger la saisie.", 2)
    y, gap = H - 69 * mm, 5 * mm
    card_width = (W - 2 * M - gap) / 2
    section_card(c, M, y, card_width, "Vérifier l'encadrant", "Le nom de l'encadrant habituel est affiché sous la séance sélectionnée.")
    section_card(c, M + card_width + gap, y, card_width, "Remplacer un encadrant", "Touchez Remplacer, recherchez un encadrant du club, puis choisissez-le. Le changement vaut pour cette séance seulement.")
    y -= 49 * mm
    section_card(c, M, y, card_width, "Zone de confort", "Confort : facile, fluide et maîtrisé.<br/>Challenge : exigeant mais maîtrisé.<br/>Limite : trop difficile ; à adapter ou arrêter.")
    section_card(c, M + card_width + gap, y, card_width, "Observation", "Ajoutez un commentaire facultatif. La zone s'agrandit jusqu'à trois lignes puis peut défiler.")
    y -= 49 * mm
    section_card(c, M, y, card_width, "Corriger une erreur", "Touchez la croix à côté d'un participant pour le retirer avant l'enregistrement.")
    section_card(c, M + card_width + gap, y, card_width, "Participant invité", "Le badge invité indique une personne connue du club, mais non membre officiel de la ligne.")

    y -= 50 * mm
    c.setFillColor(colors.HexColor("#0b4446"))
    c.setFont("Helvetica-Bold", 14)
    c.drawString(M, y, "Un seul lien pour toutes les lignes")
    link_style = ParagraphStyle("unique_link", fontName="Helvetica", fontSize=11,
                                leading=16, textColor=INK)
    c.setFillColor(colors.HexColor("#e8f5f6"))
    c.setStrokeColor(LINE)
    c.roundRect(M, y - 27 * mm, W - 2 * M, 20 * mm, 5 * mm, fill=1, stroke=1)
    draw_paragraph(c, "<b>https://apneeattitude.github.io/AppSuiviAA/</b><br/>Ouvrez le menu en haut à droite pour choisir L1, L2, L3, L4, LC, DNF1, DNF2, STA1 ou STA2.",
                   M + 7 * mm, y - 11 * mm, W - 2 * M - 14 * mm, link_style)
    c.save()


if __name__ == "__main__":
    main()
