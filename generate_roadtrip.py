#!/usr/bin/env python3
"""
generate_roadtrip.py
--------------------
Generates a visually rich Excel file for a 5-day France road trip.
Run:  python generate_roadtrip.py
Output: Road_Trip_France_5_Jours.xlsx
"""

import io
import os
import tempfile

import requests
from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import (
    Alignment, Border, Font, GradientFill, PatternFill, Side
)
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Palette
# ---------------------------------------------------------------------------
C_NIGHT   = "1A1A2E"   # bleu nuit
C_RED     = "E74C3C"   # rouge vif
C_SLATE   = "2C3E50"   # ardoise
C_LGRAY   = "F0F3F4"   # gris clair
C_WHITE   = "FDFEFE"   # blanc cassé
C_BORDER  = "CCCCCC"   # gris bordure

DAY_COLORS = {
    "JOUR 1": "4472C4",
    "JOUR 2": "ED7D31",
    "JOUR 3": "70AD47",
    "JOUR 4": "9E480E",
    "JOUR 5": "264478",
}

# ---------------------------------------------------------------------------
# Itinerary data
# ---------------------------------------------------------------------------
ITINERARY = [
    {
        "day": "JOUR 1",
        "place": "Le Mont-Saint-Michel",
        "region": "Normandie",
        "category": "🏰 Monument",
        "distance": "350 km de Paris",
        "duration": "4h30 de route",
        "maps_url": "https://www.google.com/maps/place/Le+Mont-Saint-Michel",
        "tips": "Arrivée idéale à marée montante. Visiter l'abbaye (9h-19h). Balade sur les remparts.",
        "img_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/5/56/Mont_Saint-Michel_at_night.jpg/800px-Mont_Saint-Michel_at_night.jpg",
    },
    {
        "day": "JOUR 2",
        "place": "Puy du Fou",
        "region": "Pays de la Loire",
        "category": "🎭 Parc d'attractions",
        "distance": "280 km du Mont-Saint-Michel",
        "duration": "3h de route",
        "maps_url": "https://www.google.com/maps/place/Puy+du+Fou",
        "tips": "Réserver les billets à l'avance. Journée complète recommandée. Cinéscénie le soir.",
        "img_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/4/45/Cinesc%C3%A9nie_2019.jpg/800px-Cinesc%C3%A9nie_2019.jpg",
    },
    {
        "day": "JOUR 2",
        "place": "Dune du Pilat",
        "region": "Nouvelle-Aquitaine",
        "category": "🏖️ Nature",
        "distance": "200 km du Puy du Fou",
        "duration": "2h de route",
        "maps_url": "https://www.google.com/maps/place/Dune+du+Pilat",
        "tips": "Plus haute dune d'Europe (110 m). Coucher de soleil depuis le sommet. Parking payant.",
        "img_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/b/b0/Dune_du_Pilat_grande.jpg/800px-Dune_du_Pilat_grande.jpg",
    },
    {
        "day": "JOUR 3",
        "place": "Rocamadour",
        "region": "Occitanie",
        "category": "🏘️ Village perché",
        "distance": "300 km de la Dune",
        "duration": "3h30 de route",
        "maps_url": "https://www.google.com/maps/place/Rocamadour",
        "tips": "Ascension des grands escaliers. Sanctuaire de la Vierge Noire. Village médiéval.",
        "img_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/5/5e/Rocamadour_-_panoramio_%281%29.jpg/800px-Rocamadour_-_panoramio_%281%29.jpg",
    },
    {
        "day": "JOUR 3",
        "place": "Gouffre de Padirac",
        "region": "Occitanie",
        "category": "🕳️ Grotte",
        "distance": "20 km de Rocamadour",
        "duration": "30 min de route",
        "maps_url": "https://www.google.com/maps/place/Gouffre+de+Padirac",
        "tips": "Visite en barque sous terre. Réserver en ligne. Ouvert avril-novembre.",
        "img_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/9/9c/Gouffre_de_Padirac.jpg/800px-Gouffre_de_Padirac.jpg",
    },
    {
        "day": "JOUR 4",
        "place": "Gorges du Verdon",
        "region": "Provence-Alpes-Côte d'Azur",
        "category": "🏞️ Canyon",
        "distance": "480 km de Padirac",
        "duration": "5h de route",
        "maps_url": "https://www.google.com/maps/place/Verdon+Gorge",
        "tips": "Route des Crêtes. Kayak ou pédalo sur le lac de Sainte-Croix. Vue depuis la Palud.",
        "img_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/f/f4/Gorges_du_Verdon_01.jpg/800px-Gorges_du_Verdon_01.jpg",
    },
    {
        "day": "JOUR 4",
        "place": "Calanques de Cassis",
        "region": "Provence-Alpes-Côte d'Azur",
        "category": "🌊 Mer & Nature",
        "distance": "120 km du Verdon",
        "duration": "1h30 de route",
        "maps_url": "https://www.google.com/maps/place/Les+Calanques+De+Cassis",
        "tips": "Bateau depuis le port de Cassis. Randonnée calanque d'En-Vau. Baignade possible.",
        "img_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/4/4a/En-Vau_Calanque_Cassis.jpg/800px-En-Vau_Calanque_Cassis.jpg",
    },
    {
        "day": "JOUR 5",
        "place": "Cirque du Fer à Cheval",
        "region": "Alpes",
        "category": "⛰️ Montagne",
        "distance": "310 km de Cassis",
        "duration": "3h30 de route",
        "maps_url": "https://www.google.com/maps/place/Cirque+du+Fer+%C3%A0+Cheval",
        "tips": "Cascades en forme de fer à cheval. Randonnée jusqu'au Bout du Monde. Entrée gratuite.",
        "img_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Cirque_du_Fer_%C3%A0_Cheval.jpg/800px-Cirque_du_Fer_%C3%A0_Cheval.jpg",
    },
    {
        "day": "JOUR 5",
        "place": "QC Terme Chamonix",
        "region": "Alpes",
        "category": "💆 Spa & Bien-être",
        "distance": "50 km du Cirque",
        "duration": "1h de route",
        "maps_url": "https://www.google.com/maps/place/QC+Terme+Chamonix",
        "tips": "Spa de luxe avec vue sur le Mont-Blanc. Réserver absolument. Fin parfaite du road trip.",
        "img_url": "https://upload.wikimedia.org/wikipedia/commons/thumb/9/9c/Chamonix_vue_panoramique.jpg/800px-Chamonix_vue_panoramique.jpg",
    },
]

# ---------------------------------------------------------------------------
# Planning data (tab 2)
# ---------------------------------------------------------------------------
PLANNING = [
    {
        "day": "JOUR 1",
        "label": "Le Mont-Saint-Michel · Normandie",
        "matin": "Départ de Paris (6h00). Route vers la Normandie (A13/N175). Arrivée ~10h30.",
        "aprem": "Traversée à pied jusqu'au Mont. Visite de l'abbaye bénédictine. Balade sur les remparts.",
        "soir": "Dîner dans un restaurant du Mont (agneau pré-salé). Coucher de soleil sur la baie.",
        "hotel": "🏨 Hôtel Normandie Avranches ★★★",
    },
    {
        "day": "JOUR 2",
        "label": "Puy du Fou + Dune du Pilat · Pays de la Loire / Gironde",
        "matin": "Route vers le Puy du Fou (3h). Arrivée à l'ouverture (9h00). Spectacles historiques.",
        "aprem": "Suite du Puy du Fou (Les Vikings, le Signe du Triomphe). Route vers la Dune du Pilat.",
        "soir": "Montée de la dune au coucher du soleil. Dîner à Arcachon (fruits de mer).",
        "hotel": "🏨 Hôtel du Parc Puy du Fou / Hôtel Le Dauphin Arcachon ★★★★",
    },
    {
        "day": "JOUR 3",
        "label": "Rocamadour + Gouffre de Padirac · Occitanie",
        "matin": "Route vers Rocamadour (3h30). Visite du sanctuaire et de la chapelle Notre-Dame.",
        "aprem": "Déjeuner à Rocamadour. Route vers le Gouffre de Padirac (30 min). Visite souterraine.",
        "soir": "Dîner à Gramat ou Figeac. Promenade dans les ruelles de Rocamadour illuminées.",
        "hotel": "🏨 Hôtel Les Vieilles Tours Rocamadour ★★★",
    },
    {
        "day": "JOUR 4",
        "label": "Gorges du Verdon + Calanques de Cassis · Provence",
        "matin": "Route vers le Verdon (5h). Route des Crêtes. Kayak ou pédalo sur le lac de Sainte-Croix.",
        "aprem": "Route vers Cassis (1h30). Bateau pour les calanques d'En-Vau et Port-Pin. Baignade.",
        "soir": "Dîner au port de Cassis. Rosé provençal au coucher du soleil.",
        "hotel": "🏨 Hôtel Les Roches Blanches Cassis ★★★★★",
    },
    {
        "day": "JOUR 5",
        "label": "Cirque du Fer à Cheval + QC Terme Chamonix · Alpes",
        "matin": "Route vers Sixt-Fer-à-Cheval (3h30). Randonnée dans le cirque. Cascades spectaculaires.",
        "aprem": "Route vers Chamonix (1h). Entrée au QC Terme. Piscines avec vue sur le Mont-Blanc.",
        "soir": "Soirée spa (sauna, bains extérieurs). Dîner à Chamonix. Nuit alpine bien méritée.",
        "hotel": "🏨 QC Terme Chamonix ou Hôtel du Mont-Blanc ★★★★★",
    },
]

# ---------------------------------------------------------------------------
# Route data (tab 3)
# ---------------------------------------------------------------------------
ROUTE = [
    ("1", "Paris",                      "Mont-Saint-Michel",        "350 km",  "4h30"),
    ("2", "Mont-Saint-Michel",          "Puy du Fou",               "280 km",  "3h00"),
    ("3", "Puy du Fou",                 "Dune du Pilat",            "200 km",  "2h00"),
    ("4", "Dune du Pilat",              "Rocamadour",               "300 km",  "3h30"),
    ("5", "Rocamadour",                 "Gouffre de Padirac",       "20 km",   "0h30"),
    ("6", "Gouffre de Padirac",         "Gorges du Verdon",         "480 km",  "5h00"),
    ("7", "Gorges du Verdon",           "Calanques de Cassis",      "120 km",  "1h30"),
    ("8", "Calanques de Cassis",        "Cirque du Fer à Cheval",   "310 km",  "3h30"),
    ("9", "Cirque du Fer à Cheval",     "Chamonix (QC Terme)",      "50 km",   "1h00"),
    ("10", "Chamonix",                  "Paris",                    "600 km",  "6h00"),
]

# ---------------------------------------------------------------------------
# Checklist data (tab 4)
# ---------------------------------------------------------------------------
CHECKLIST = {
    "📄 Documents indispensables": [
        "Carte d'identité / Passeport",
        "Permis de conduire",
        "Carte grise du véhicule",
        "Assurance voiture (carte verte)",
        "Réservations hôtels imprimées",
        "Billets Puy du Fou / Gouffre de Padirac",
        "Carte bancaire + espèces",
        "Assurance voyage / santé",
    ],
    "🎒 Bagages essentiels": [
        "Vêtements chauds pour les Alpes",
        "Maillot de bain (Calanques, QC Terme)",
        "Chaussures de randonnée",
        "Crème solaire haute protection",
        "Lunettes de soleil",
        "Trousse de premiers secours",
        "Bouteille d'eau réutilisable",
        "Appareil photo / chargeur",
    ],
    "🚗 Préparatifs voiture": [
        "Révision / niveaux vérifiés",
        "Pression des pneus",
        "Kit de sécurité (gilet, triangle)",
        "GPS / supports téléphone",
        "Chargeur voiture USB",
        "Musique / podcasts téléchargés",
        "Glacière / snacks pour la route",
    ],
    "📱 Apps utiles": [
        "Google Maps (cartes hors-ligne téléchargées)",
        "Waze (trafic temps réel)",
        "Booking.com / Airbnb",
        "Météo France",
        "Komoot (randonnées)",
        "TripAdvisor",
        "Visite+ (audioguides monuments)",
    ],
    "💡 Conseils pratiques": [
        "Réserver les hébergements à l'avance",
        "Arriver tôt aux sites touristiques populaires",
        "Prévoir des espèces pour les parkings",
        "Vérifier les horaires d'ouverture avant chaque visite",
        "Télécharger les cartes Google Maps hors-ligne",
        "Garder 30% de batterie de téléphone en permanence",
        "Photographier les plaques de parking pour se souvenir du lieu",
    ],
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def font(name="Calibri", size=11, bold=False, italic=False, color="000000"):
    return Font(name=name, size=size, bold=bold, italic=italic, color=color)


def border(color=C_BORDER):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)


def align(h="left", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def set_cell(ws, row, col, value,
             fnt=None, fll=None, brd=None, aln=None, hyperlink=None):
    c = ws.cell(row=row, column=col, value=value)
    if fnt:
        c.font = fnt
    if fll:
        c.fill = fll
    if brd:
        c.border = brd
    if aln:
        c.alignment = aln
    if hyperlink:
        c.hyperlink = hyperlink
    return c


def download_image(url, size=(240, 160)):
    """Download image from URL, resize, return NamedTemporaryFile path or None."""
    headers = {"User-Agent": "Mozilla/5.0 (RoadTripBot/1.0)"}
    try:
        r = requests.get(url, headers=headers, timeout=15)
        r.raise_for_status()
        img = Image.open(io.BytesIO(r.content)).convert("RGB")
        img.thumbnail(size, Image.LANCZOS)
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        img.save(tmp.name, "PNG")
        tmp.close()
        return tmp.name
    except Exception as exc:
        print(f"  ⚠️  Could not download {url}: {exc}")
        return None


def hide_gridlines(ws):
    ws.sheet_view.showGridLines = False


# ---------------------------------------------------------------------------
# Tab 1 — Itinéraire
# ---------------------------------------------------------------------------

def build_itineraire(wb, img_paths):
    ws = wb.create_sheet("🗺️ Itinéraire")
    hide_gridlines(ws)

    # Column widths
    col_widths = [12, 28, 24, 22, 22, 16, 20, 42, 32]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ---- Row 1: Grand titre ----
    ws.row_dimensions[1].height = 55
    ws.merge_cells("A1:I1")
    c = ws["A1"]
    c.value = "🇫🇷  ROAD TRIP FRANCE — 5 JOURS INOUBLIABLES"
    c.font = Font(name="Calibri", size=26, bold=True, color="FDFEFE")
    c.fill = fill(C_NIGHT)
    c.alignment = align("center", "center")

    # ---- Row 2: Sous-titre ----
    ws.row_dimensions[2].height = 30
    ws.merge_cells("A2:I2")
    c = ws["A2"]
    c.value = (
        "Mont-Saint-Michel  ·  Puy du Fou  ·  Dune du Pilat  ·  Rocamadour  ·  "
        "Gorges du Verdon  ·  Calanques de Cassis  ·  Chamonix"
    )
    c.font = Font(name="Calibri", size=12, italic=True, color="BDC3C7")
    c.fill = fill(C_NIGHT)
    c.alignment = align("center", "center")

    # ---- Row 3: Barre rouge ----
    ws.row_dimensions[3].height = 6
    ws.merge_cells("A3:I3")
    ws["A3"].fill = fill(C_RED)

    # ---- Row 4: En-têtes colonnes ----
    ws.row_dimensions[4].height = 30
    headers = ["JOUR", "LIEU", "RÉGION", "CATÉGORIE", "DISTANCE", "DURÉE", "LIEN MAPS",
               "CONSEILS PRATIQUES", "📸 PHOTO"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=4, column=col, value=h)
        c.font = Font(name="Calibri", size=10, bold=True, color="FDFEFE")
        c.fill = fill(C_SLATE)
        c.border = border()
        c.alignment = align("center", "center")

    # ---- Data rows ----
    ROW_H_PX = 130   # pixels → ~97 pt  (1 pt ≈ 0.75 px, Excel uses pt internally ~= px/0.75)
    ROW_H_PT = 98    # approximate

    for idx, stop in enumerate(ITINERARY):
        r = 5 + idx
        ws.row_dimensions[r].height = ROW_H_PT

        row_fill = fill(C_WHITE) if idx % 2 == 0 else fill(C_LGRAY)
        day_hex = DAY_COLORS.get(stop["day"], C_NIGHT)

        values = [
            stop["day"],
            stop["place"],
            stop["region"],
            stop["category"],
            stop["distance"],
            stop["duration"],
            "📍 Google Maps",
            stop["tips"],
            "",  # image column
        ]

        for col, val in enumerate(values, 1):
            c = ws.cell(row=r, column=col, value=val)
            if col == 1:
                c.fill = fill(day_hex)
                c.font = Font(name="Calibri", size=10, bold=True, color="FDFEFE")
            else:
                c.fill = row_fill
                c.font = Font(name="Calibri", size=10, color="2C3E50")
            c.border = border()
            c.alignment = align("center", "center", wrap=True)
            if col == 7:  # LIEN MAPS
                c.hyperlink = stop["maps_url"]
                c.font = Font(name="Calibri", size=10, color="1565C0", underline="single")

        # Insert image if available
        img_path = img_paths[idx]
        if img_path and os.path.exists(img_path):
            try:
                xl_img = XLImage(img_path)
                xl_img.width = 220
                xl_img.height = 120
                cell_ref = f"{get_column_letter(9)}{r}"
                ws.add_image(xl_img, cell_ref)
            except Exception as exc:
                print(f"  ⚠️  Could not insert image for {stop['place']}: {exc}")
                ws.cell(row=r, column=9, value=f"🖼️ {stop['place']}")
        else:
            ws.cell(row=r, column=9).value = f"🖼️ {stop['place']}"

    # ---- Footer ----
    footer_row = 5 + len(ITINERARY)
    ws.row_dimensions[footer_row].height = 24
    ws.merge_cells(f"A{footer_row}:I{footer_row}")
    c = ws[f"A{footer_row}"]
    c.value = "✈️  Bon voyage !  |  Total trajet estimé : ~2 800 km  |  ~30h de route  |  © Road Trip France 2025"
    c.font = Font(name="Calibri", size=10, italic=True, color="BDC3C7")
    c.fill = fill(C_NIGHT)
    c.alignment = align("center", "center")


# ---------------------------------------------------------------------------
# Tab 2 — Planning détaillé
# ---------------------------------------------------------------------------

def build_planning(wb):
    ws = wb.create_sheet("📅 Planning détaillé")
    hide_gridlines(ws)

    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 60
    ws.column_dimensions["D"].width = 30

    current_row = 1

    for plan in PLANNING:
        day_hex = DAY_COLORS.get(plan["day"], C_SLATE)

        # Day title
        ws.row_dimensions[current_row].height = 36
        ws.merge_cells(f"A{current_row}:D{current_row}")
        c = ws[f"A{current_row}"]
        c.value = f"  {plan['day']}  —  {plan['label']}"
        c.font = Font(name="Calibri", size=14, bold=True, color="FDFEFE")
        c.fill = fill(day_hex)
        c.alignment = align("left", "center")
        current_row += 1

        # Sub-header
        ws.row_dimensions[current_row].height = 22
        for col, header in enumerate(["MOMENT", "ACTIVITÉ", "DÉTAILS", "HÉBERGEMENT"], 1):
            c = ws.cell(row=current_row, column=col, value=header)
            c.font = Font(name="Calibri", size=9, bold=True, color="FDFEFE")
            c.fill = fill(C_SLATE)
            c.border = border()
            c.alignment = align("center", "center")
        current_row += 1

        # Morning / Afternoon / Evening rows
        moments = [
            ("🌅 Matin",      plan["matin"], ""),
            ("☀️ Après-midi", plan["aprem"], ""),
            ("🌙 Soir",       plan["soir"],  plan["hotel"]),
        ]
        for i, (moment, detail, hotel) in enumerate(moments):
            ws.row_dimensions[current_row].height = 50
            row_fill = fill(C_WHITE) if i % 2 == 0 else fill(C_LGRAY)

            c = ws.cell(row=current_row, column=1, value=moment)
            c.font = Font(name="Calibri", size=10, bold=True, color=day_hex)
            c.fill = row_fill
            c.border = border()
            c.alignment = align("center", "center", wrap=True)

            c = ws.cell(row=current_row, column=2, value="")
            c.fill = row_fill
            c.border = border()

            c = ws.cell(row=current_row, column=3, value=detail)
            c.font = Font(name="Calibri", size=10, color=C_SLATE)
            c.fill = row_fill
            c.border = border()
            c.alignment = align("left", "center", wrap=True)

            c = ws.cell(row=current_row, column=4, value=hotel)
            c.font = Font(name="Calibri", size=10, bold=bool(hotel), color="1565C0" if hotel else C_SLATE)
            c.fill = row_fill
            c.border = border()
            c.alignment = align("left", "center", wrap=True)

            current_row += 1

        # Spacing row
        ws.row_dimensions[current_row].height = 10
        current_row += 1


# ---------------------------------------------------------------------------
# Tab 3 — Carte du trajet
# ---------------------------------------------------------------------------

def build_carte(wb):
    ws = wb.create_sheet("🚗 Carte du trajet")
    hide_gridlines(ws)

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 16
    ws.column_dimensions["E"].width = 16

    # Title
    ws.row_dimensions[1].height = 50
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = "🚗  CARTE DU TRAJET — ROAD TRIP FRANCE 5 JOURS"
    c.font = Font(name="Calibri", size=20, bold=True, color="FDFEFE")
    c.fill = fill(C_NIGHT)
    c.alignment = align("center", "center")

    # Sub-info
    ws.row_dimensions[2].height = 24
    ws.merge_cells("A2:E2")
    c = ws["A2"]
    c.value = "Boucle au départ de Paris  ·  Total estimé : ~2 800 km  ·  ~30h de route"
    c.font = Font(name="Calibri", size=11, italic=True, color="BDC3C7")
    c.fill = fill(C_SLATE)
    c.alignment = align("center", "center")

    # Red separator
    ws.row_dimensions[3].height = 5
    ws.merge_cells("A3:E3")
    ws["A3"].fill = fill(C_RED)

    # Headers
    ws.row_dimensions[4].height = 28
    for col, h in enumerate(["#", "DÉPART", "ARRIVÉE", "DISTANCE", "DURÉE ROUTE"], 1):
        c = ws.cell(row=4, column=col, value=h)
        c.font = Font(name="Calibri", size=10, bold=True, color="FDFEFE")
        c.fill = fill(C_SLATE)
        c.border = border()
        c.alignment = align("center", "center")

    # Data
    for idx, (num, dep, arr, dist, dur) in enumerate(ROUTE):
        r = 5 + idx
        ws.row_dimensions[r].height = 26
        row_fill = fill(C_WHITE) if idx % 2 == 0 else fill(C_LGRAY)
        for col, val in enumerate([num, dep, arr, dist, dur], 1):
            c = ws.cell(row=r, column=col, value=val)
            c.font = Font(name="Calibri", size=10, color=C_SLATE)
            c.fill = row_fill
            c.border = border()
            c.alignment = align("center", "center")

    # Total row
    total_row = 5 + len(ROUTE)
    ws.row_dimensions[total_row].height = 30
    for col, val in enumerate(["", "TOTAL", "", "~2 800 km", "~30h00"], 1):
        c = ws.cell(row=total_row, column=col, value=val)
        c.font = Font(name="Calibri", size=11, bold=True, color="FDFEFE")
        c.fill = fill(C_RED)
        c.border = border()
        c.alignment = align("center", "center")


# ---------------------------------------------------------------------------
# Tab 4 — Checklist voyage
# ---------------------------------------------------------------------------

def build_checklist(wb):
    ws = wb.create_sheet("✅ Checklist voyage")
    hide_gridlines(ws)

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 55

    # Title
    ws.row_dimensions[1].height = 50
    ws.merge_cells("A1:B1")
    c = ws["A1"]
    c.value = "✅  CHECKLIST VOYAGE — ROAD TRIP FRANCE"
    c.font = Font(name="Calibri", size=20, bold=True, color="FDFEFE")
    c.fill = fill(C_NIGHT)
    c.alignment = align("center", "center")

    current_row = 2

    for section_title, items in CHECKLIST.items():
        # Section header
        ws.row_dimensions[current_row].height = 30
        ws.merge_cells(f"A{current_row}:B{current_row}")
        c = ws[f"A{current_row}"]
        c.value = f"  {section_title}"
        c.font = Font(name="Calibri", size=12, bold=True, color="FDFEFE")
        c.fill = fill(C_SLATE)
        c.alignment = align("left", "center")
        current_row += 1

        for i, item in enumerate(items):
            ws.row_dimensions[current_row].height = 22
            row_fill = fill(C_WHITE) if i % 2 == 0 else fill(C_LGRAY)

            c = ws.cell(row=current_row, column=1, value="☐")
            c.font = Font(name="Calibri", size=14, color=C_RED)
            c.fill = row_fill
            c.border = border()
            c.alignment = align("center", "center")

            c = ws.cell(row=current_row, column=2, value=f"  {item}")
            c.font = Font(name="Calibri", size=10, color=C_SLATE)
            c.fill = row_fill
            c.border = border()
            c.alignment = align("left", "center")

            current_row += 1

        # Spacing row
        ws.row_dimensions[current_row].height = 8
        current_row += 1

    # Footer
    ws.row_dimensions[current_row].height = 24
    ws.merge_cells(f"A{current_row}:B{current_row}")
    c = ws[f"A{current_row}"]
    c.value = "🚗  Bon road trip !  Profite de chaque kilomètre.  © Road Trip France 2025"
    c.font = Font(name="Calibri", size=10, italic=True, color="BDC3C7")
    c.fill = fill(C_NIGHT)
    c.alignment = align("center", "center")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    output_file = "Road_Trip_France_5_Jours.xlsx"

    print("🇫🇷  Génération du Road Trip France — 5 Jours")
    print("=" * 55)

    # 1. Download images
    print("\n📸 Téléchargement des images…")
    img_paths = []
    tmp_files = []
    for stop in ITINERARY:
        print(f"  → {stop['place']}")
        path = download_image(stop["img_url"])
        img_paths.append(path)
        if path:
            tmp_files.append(path)

    # 2. Build workbook
    print("\n📊 Création du fichier Excel…")
    wb = Workbook()
    wb.remove(wb.active)  # remove default sheet

    print("  → Onglet 1 : 🗺️ Itinéraire")
    build_itineraire(wb, img_paths)

    print("  → Onglet 2 : 📅 Planning détaillé")
    build_planning(wb)

    print("  → Onglet 3 : 🚗 Carte du trajet")
    build_carte(wb)

    print("  → Onglet 4 : ✅ Checklist voyage")
    build_checklist(wb)

    # 3. Save
    wb.save(output_file)
    print(f"\n✅  Fichier généré : {output_file}")

    # 4. Cleanup temp files
    for path in tmp_files:
        try:
            os.unlink(path)
        except OSError:
            pass

    print("🎉  Terminé ! Ouvre le fichier Excel pour découvrir ton road trip.")


if __name__ == "__main__":
    main()
