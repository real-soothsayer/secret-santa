#!/usr/bin/env python3
"""
Génère un fichier Excel premium pour un road trip 5 jours en France.
Uniquement des attractions notées 4 étoiles et plus sur Google Maps.
"""

import io
import os
import requests
from openpyxl import Workbook
from openpyxl.styles import (
    Font, Alignment, PatternFill, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

# ---------------------------------------------------------------------------
# DONNÉES DU VOYAGE
# ---------------------------------------------------------------------------

TRIP_DATA = [
    {"jour":"JOUR 1","note":"4.7","lieu":"Etretat - Falaises d'Aval","region":"Normandie","categorie":"Falaises et sentier cotier","distance_km":"~200 km depuis Paris","duree":"2-3 heures","maps_url":"https://www.google.com/maps/place/Etretat","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/c/ce/Etretat_-_Falaise_d%27Aval_-_2012.jpg/800px-Etretat_-_Falaise_d%27Aval_-_2012.jpg","conseil":"Prendre le sentier des falaises. Arche naturelle et Aiguille Creuse. Lever de soleil depuis la falaise d'Amont.","matin":"Depart Paris A13. Randonnee sur les falaises d'Aval et d'Amont.","aprem":"Route vers Mont-Saint-Michel (3h).","soir":"","hebergement":"Hotel Mercure Mont-Saint-Michel","couleur_jour":"4472C4"},
    {"jour":"JOUR 1","note":"4.7","lieu":"Le Mont-Saint-Michel","region":"Normandie","categorie":"Monument UNESCO","distance_km":"~200 km depuis Etretat","duree":"4-5 heures","maps_url":"https://www.google.com/maps/place/Le+Mont-Saint-Michel","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/5/56/Mont_Saint-Michel_at_night.jpg/800px-Mont_Saint-Michel_at_night.jpg","conseil":"Arriver en fin d'apres-midi pour la lumiere doree. Rester pour voir le Mont illumine la nuit. Maree haute = spectaculaire!","matin":"","aprem":"Visite de l'abbaye et des remparts.","soir":"Diner fruits de mer sur place. Mont illumine la nuit.","hebergement":"Hotel Mercure Mont-Saint-Michel","couleur_jour":"4472C4"},
    {"jour":"JOUR 2","note":"4.7","lieu":"Chateau de Chenonceau","region":"Val de Loire","categorie":"Chateau Renaissance","distance_km":"~250 km depuis Mont-Saint-Michel","duree":"2 heures","maps_url":"https://www.google.com/maps/place/Chateau+de+Chenonceau","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/d/d1/Chenonceau_Castle_-_2013.jpg/800px-Chenonceau_Castle_-_2013.jpg","conseil":"Surnomme le chateau des Dames. Enjambe le Cher. Jardins sublimes.","matin":"Route vers le Val de Loire (2h30). Visite de Chenonceau des l'ouverture (9h).","aprem":"","soir":"","hebergement":"Hotel du Parc Puy du Fou","couleur_jour":"ED7D31"},
    {"jour":"JOUR 2","note":"4.6","lieu":"Chateau de Chambord","region":"Val de Loire","categorie":"Chateau Renaissance","distance_km":"~40 km depuis Chenonceau","duree":"2 heures","maps_url":"https://www.google.com/maps/place/Chateau+de+Chambord","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/7/72/Ch%C3%A2teau_de_Chambord_2008.jpg/800px-Ch%C3%A2teau_de_Chambord_2008.jpg","conseil":"Le plus grand chateau de la Loire. Escalier a double revolution de Leonard de Vinci. Parc de 5440 ha avec cerfs.","matin":"","aprem":"Visite de Chambord. Escalier double helice. Terrasses avec vue sur 365 cheminees.","soir":"","hebergement":"Hotel du Parc Puy du Fou","couleur_jour":"ED7D31"},
    {"jour":"JOUR 2","note":"4.8","lieu":"Puy du Fou","region":"Pays de la Loire","categorie":"Meilleur parc du monde","distance_km":"~150 km depuis Chambord","duree":"Soiree complete","maps_url":"https://www.google.com/maps/place/Puy+du+Fou","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/4/45/Cinesc%C3%A9nie_2019.jpg/800px-Cinesc%C3%A9nie_2019.jpg","conseil":"Elu meilleur parc du monde. Cinesecenie nocturne avec 2500 acteurs. RESERVER A L'AVANCE!","matin":"","aprem":"Arrivee et premiers spectacles.","soir":"Cinesecenie nocturne - show grandiose. Nuit sur place.","hebergement":"Hotel du Parc Puy du Fou","couleur_jour":"ED7D31"},
    {"jour":"JOUR 3","note":"4.6","lieu":"La Roque-Gageac","region":"Perigord Noir - Dordogne","categorie":"Plus Beau Village de France","distance_km":"~280 km depuis Puy du Fou","duree":"1h30-2 heures","maps_url":"https://www.google.com/maps/place/La+Roque-Gageac","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/2/26/La_Roque-Gageac.jpg/800px-La_Roque-Gageac.jpg","conseil":"Croisiere en gabarre sur la Dordogne (1h). Vegetation tropicale accrochee a la falaise.","matin":"Route vers le Perigord (3h). Croisiere en gabarre sur la Dordogne.","aprem":"","soir":"","hebergement":"Hotel Les Vieilles Tours Rocamadour","couleur_jour":"70AD47"},
    {"jour":"JOUR 3","note":"4.7","lieu":"Beynac-et-Cazenac","region":"Perigord Noir - Dordogne","categorie":"Chateau medieval panoramique","distance_km":"~10 km depuis La Roque-Gageac","duree":"1-2 heures","maps_url":"https://www.google.com/maps/place/Beynac-et-Cazenac","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/3/30/Ch%C3%A2teau_de_Beynac.jpg/800px-Ch%C3%A2teau_de_Beynac.jpg","conseil":"Vue panoramique sublime sur la vallee de la Dordogne depuis le chateau. Parmi les plus beaux villages de France.","matin":"","aprem":"Montee au chateau. Vue a 360 sur la vallee.","soir":"","hebergement":"Hotel Les Vieilles Tours Rocamadour","couleur_jour":"70AD47"},
    {"jour":"JOUR 3","note":"4.6","lieu":"Rocamadour","region":"Occitanie - Quercy","categorie":"Village perche et Sanctuaire","distance_km":"~55 km depuis Beynac","duree":"2-3 heures","maps_url":"https://www.google.com/maps/place/Rocamadour","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/5/5e/Rocamadour_-_panoramio_%281%29.jpg/800px-Rocamadour_-_panoramio_%281%29.jpg","conseil":"Gravir le chemin de croix a pied (216 marches). Vue imprenable depuis le chateau. Lumiere doree en fin d'apres-midi.","matin":"","aprem":"Chemin de croix, sanctuaire de la Vierge Noire, chateau.","soir":"Diner agneau du Quercy.","hebergement":"Hotel Les Vieilles Tours Rocamadour","couleur_jour":"70AD47"},
    {"jour":"JOUR 3","note":"4.6","lieu":"Gouffre de Padirac","region":"Occitanie - Quercy","categorie":"Grotte souterraine","distance_km":"~30 km depuis Rocamadour","duree":"2 heures","maps_url":"https://www.google.com/maps/place/Gouffre+de+Padirac","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/9/9c/Gouffre_de_Padirac.jpg/800px-Gouffre_de_Padirac.jpg","conseil":"Reservation en ligne OBLIGATOIRE en ete. Veste indispensable (12 C). Balade en barque souterraine feerique.","matin":"","aprem":"","soir":"Descente dans le gouffre. Riviere souterraine en barque.","hebergement":"Hotel Les Vieilles Tours Rocamadour","couleur_jour":"70AD47"},
    {"jour":"JOUR 4","note":"4.7","lieu":"Sentier des Ocres de Roussillon","region":"Provence - Luberon","categorie":"Randonnee paysage martien","distance_km":"~350 km depuis Rocamadour","duree":"1h30-2 heures","maps_url":"https://www.google.com/maps/place/Roussillon","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/4/4c/Roussillon_-_ocres.jpg/800px-Roussillon_-_ocres.jpg","conseil":"Paysage rouge-orange unique en Europe. Vetements que vous ne craignez pas (ocres tachent!). Circuit court = 35 min.","matin":"Grande route vers la Provence (3h30). Sentier des Ocres.","aprem":"","soir":"","hebergement":"Hotel Les Roches Blanches Cassis","couleur_jour":"9E480E"},
    {"jour":"JOUR 4","note":"4.6","lieu":"Les Baux-de-Provence","region":"Provence - Alpilles","categorie":"Village perche sur rocher","distance_km":"~80 km depuis Roussillon","duree":"1h30-2 heures","maps_url":"https://www.google.com/maps/place/Les+Baux-de-Provence","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/e/e6/Les_Baux-de-Provence_village.jpg/800px-Les_Baux-de-Provence_village.jpg","conseil":"Village perche sur rocher calcaire. Chateau en ruines avec vue panoramique sur les Alpilles.","matin":"","aprem":"Visite du village medieval et du chateau des Baux.","soir":"","hebergement":"Hotel Les Roches Blanches Cassis","couleur_jour":"9E480E"},
    {"jour":"JOUR 4","note":"4.7","lieu":"Carrieres de Lumieres","region":"Provence - Alpilles","categorie":"Art immersif monumental","distance_km":"~500 m depuis Les Baux","duree":"1h30","maps_url":"https://www.google.com/maps/place/Carrieres+de+Lumieres","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/f/f9/Carrieres_lumieres_les_baux.jpg/800px-Carrieres_lumieres_les_baux.jpg","conseil":"Spectacle immersif dans d'anciennes carrieres de calcaire. 70 projecteurs. Monet, Klimt, Van Gogh selon les annees.","matin":"","aprem":"Spectacle immersif dans les carrieres.","soir":"Route vers Cassis (1h).","hebergement":"Hotel Les Roches Blanches Cassis","couleur_jour":"9E480E"},
    {"jour":"JOUR 4","note":"4.7","lieu":"Calanques de Cassis","region":"Provence - Cote d'Azur","categorie":"Mer turquoise et Nature","distance_km":"~80 km depuis Les Baux","duree":"3 heures","maps_url":"https://www.google.com/maps/place/Les+Calanques+De+Cassis","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/4/4a/En-Vau_Calanque_Cassis.jpg/800px-En-Vau_Calanque_Cassis.jpg","conseil":"Bateau-promenade depuis le port = vue ideale. Baignade en calanque d'En-Vau. Eau cristalline turquoise.","matin":"","aprem":"","soir":"Bateau-promenade. Diner bouillabaisse a Cassis.","hebergement":"Hotel Les Roches Blanches Cassis","couleur_jour":"9E480E"},
    {"jour":"JOUR 5","note":"4.8","lieu":"Lac d'Annecy","region":"Alpes - Haute-Savoie","categorie":"Lac le plus pur d'Europe","distance_km":"~300 km depuis Cassis","duree":"2-3 heures","maps_url":"https://www.google.com/maps/place/Lac+Annecy","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/b/b8/Lac_Annecy_vue_Talloires.jpg/800px-Lac_Annecy_vue_Talloires.jpg","conseil":"Balade sur les Jardins de l'Europe. Vieille ville d'Annecy (canaux, chateau). Baignade a la plage d'Albigny. Eau a 22 C en ete.","matin":"Route vers Annecy (3h). Balade sur les rives du lac et vieille ville.","aprem":"","soir":"","hebergement":"QC Terme Chamonix","couleur_jour":"264478"},
    {"jour":"JOUR 5","note":"4.8","lieu":"Cirque du Fer a Cheval","region":"Alpes - Haute-Savoie","categorie":"Montagne et 30 cascades","distance_km":"~50 km depuis Annecy","duree":"2-3 heures","maps_url":"https://www.google.com/maps/place/Cirque+du+Fer+a+Cheval","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/2/2f/Cirque_du_Fer_%C3%A0_Cheval.jpg/800px-Cirque_du_Fer_%C3%A0_Cheval.jpg","conseil":"30 cascades en ete dont la Grande Cascade (126m). Randonnee facile. Ambiance montagne sauvage et preservee.","matin":"","aprem":"Randonnee dans le cirque. Grandes cascades.","soir":"","hebergement":"QC Terme Chamonix","couleur_jour":"264478"},
    {"jour":"JOUR 5","note":"4.8","lieu":"Aiguille du Midi - Chamonix","region":"Alpes - Massif du Mont-Blanc","categorie":"Telepherique 3842m et Vue 360","distance_km":"~45 km depuis Cirque du Fer a Cheval","duree":"3-4 heures","maps_url":"https://www.google.com/maps/place/Aiguille+du+Midi","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/5/5b/Aiguille_du_Midi_2013.jpg/800px-Aiguille_du_Midi_2013.jpg","conseil":"Telepherique jusqu'a 3842m. Vue 360 sur le Mont-Blanc (4808m), les Alpes francaises, italiennes et suisses. RESERVER a l'avance.","matin":"","aprem":"Telepherique et panorama sur le toit de l'Europe.","soir":"","hebergement":"QC Terme Chamonix","couleur_jour":"264478"},
    {"jour":"JOUR 5","note":"4.5","lieu":"QC Terme Chamonix","region":"Alpes - Chamonix-Mont-Blanc","categorie":"Spa et Jacuzzi vue Mont-Blanc","distance_km":"~5 min depuis Chamonix centre","duree":"3-4 heures","maps_url":"https://www.google.com/maps/place/QC+Terme+Chamonix","image_url":"https://upload.wikimedia.org/wikipedia/commons/thumb/9/9c/Chamonix_vue_panoramique.jpg/800px-Chamonix_vue_panoramique.jpg","conseil":"Jacuzzi exterieur avec vue directe sur le Mont-Blanc et les glaciers. Fin du road trip en beaute. Reserver OBLIGATOIRE.","matin":"","aprem":"","soir":"Spa, piscines exterieures, sauna. Diner fondue savoyarde.","hebergement":"QC Terme Chamonix","couleur_jour":"264478"},
]

# ---------------------------------------------------------------------------
# STYLES
# ---------------------------------------------------------------------------

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def _border():
    thin = Side(style="thin", color="CCCCCC")
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _font(bold=False, size=11, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic,
                name="Calibri")


def _align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


# ---------------------------------------------------------------------------
# IMAGE HELPER
# ---------------------------------------------------------------------------

def _download_image(url, max_width=120, max_height=90):
    """Download image from URL and return an openpyxl Image object, or None."""
    try:
        headers = {"User-Agent": "Mozilla/5.0 (road-trip-generator/1.0)"}
        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status()
        buf = io.BytesIO(resp.content)
        img = XLImage(buf)
        # Scale to fit cell
        ratio = min(max_width / img.width, max_height / img.height)
        img.width = int(img.width * ratio)
        img.height = int(img.height * ratio)
        return img
    except Exception as exc:
        print(f"  [WARN] Image non disponible ({url[:60]}…): {exc}")
        return None


# ---------------------------------------------------------------------------
# ONGLET 1 — ITINÉRAIRE
# ---------------------------------------------------------------------------

def build_itineraire(wb):
    ws = wb.create_sheet("🗺️ Itinéraire", 0)

    # ── titre principal ──────────────────────────────────────────────────────
    ws.merge_cells("A1:J1")
    title_cell = ws["A1"]
    title_cell.value = "🇫🇷  ROAD TRIP 5 JOURS EN FRANCE  —  Attractions 4 étoiles et plus  🌟"
    title_cell.font = _font(bold=True, size=16, color="FFFFFF")
    title_cell.fill = _fill("1F3864")
    title_cell.alignment = _align(h="center")
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:J2")
    sub_cell = ws["A2"]
    sub_cell.value = (
        "Paris → Étretat → Mont-Saint-Michel → Val de Loire → "
        "Puy du Fou → Périgord → Rocamadour → Provence → Alpes → Paris"
    )
    sub_cell.font = _font(italic=True, size=11, color="FFFFFF")
    sub_cell.fill = _fill("2E4D8A")
    sub_cell.alignment = _align(h="center")
    ws.row_dimensions[2].height = 22

    # ── en-têtes colonnes ───────────────────────────────────────────────────
    headers = [
        "JOUR", "NOTE ★", "LIEU", "RÉGION", "CATÉGORIE",
        "DISTANCE", "DURÉE", "LIEN MAPS", "CONSEILS", "📸 PHOTO",
    ]
    col_widths = [12, 10, 28, 22, 20, 28, 14, 30, 34, 35]

    for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.font = _font(bold=True, size=11, color="FFFFFF")
        cell.fill = _fill("1F3864")
        cell.alignment = _align(h="center")
        cell.border = _border()
    ws.row_dimensions[3].height = 22

    # ── lignes de données ────────────────────────────────────────────────────
    current_jour = None
    row = 4
    for item in TRIP_DATA:
        jour = item["jour"]
        bg = item.get("couleur_jour", "FFFFFF")
        text_color = "FFFFFF" if bg not in ("FFFFFF", "F2F2F2") else "000000"

        # note formatée
        note_val = item.get("note", "")
        note_display = f"⭐ {note_val}" if note_val else ""

        values = [
            jour if jour != current_jour else "",
            note_display,
            item["lieu"],
            item["region"],
            item["categorie"],
            item["distance_km"],
            item["duree"],
            item["maps_url"],
            item["conseil"],
            "",  # photo placeholder (column J = index 10)
        ]
        if jour != current_jour:
            current_jour = jour

        ws.row_dimensions[row].height = 90

        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            cell.fill = _fill(bg)
            cell.font = _font(size=10, color=text_color)
            cell.border = _border()
            if col_idx == 8:  # LIEN MAPS — hyperlien
                cell.hyperlink = val
                cell.font = _font(size=10, color="0563C1")
                cell.value = "📍 Voir sur Maps"
            elif col_idx == 1:  # JOUR
                cell.font = _font(bold=True, size=11, color=text_color)
                cell.alignment = _align(h="center")
            elif col_idx == 2:  # NOTE
                cell.font = _font(bold=True, size=11, color=text_color)
                cell.alignment = _align(h="center")
            else:
                cell.alignment = _align(wrap=True)

        # ── image dans colonne J ─────────────────────────────────────────────
        img_url = item.get("image_url", "")
        if img_url:
            print(f"  Téléchargement image : {item['lieu']} …")
            img_obj = _download_image(img_url, max_width=120, max_height=88)
            if img_obj:
                cell_ref = f"J{row}"
                img_obj.anchor = cell_ref
                ws.add_image(img_obj)

        row += 1

    # ── freeze panes ────────────────────────────────────────────────────────
    ws.freeze_panes = "A4"

    return ws


# ---------------------------------------------------------------------------
# ONGLET 2 — TRAJET
# ---------------------------------------------------------------------------

def build_trajet(wb):
    ws = wb.create_sheet("🚗 Trajet")

    # titre
    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = "🚗  DÉTAIL DES TRAJETS — Road Trip 5 Jours en France"
    t.font = _font(bold=True, size=14, color="FFFFFF")
    t.fill = _fill("1F3864")
    t.alignment = _align(h="center")
    ws.row_dimensions[1].height = 32

    # en-têtes
    col_widths_t = [6, 28, 28, 14, 16, 20]
    headers_t = ["#", "DÉPART", "ARRIVÉE", "DISTANCE", "DURÉE APPROX.", "NOTES"]
    for col_idx, (h, w) in enumerate(zip(headers_t, col_widths_t), start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w
        cell = ws.cell(row=2, column=col_idx, value=h)
        cell.font = _font(bold=True, size=11, color="FFFFFF")
        cell.fill = _fill("2E4D8A")
        cell.alignment = _align(h="center")
        cell.border = _border()
    ws.row_dimensions[2].height = 20

    etapes = [
        (1,  "Paris",                     "Etretat",                    "~200 km",  "~2h15"),
        (2,  "Etretat",                   "Le Mont-Saint-Michel",        "~200 km",  "~2h30"),
        (3,  "Le Mont-Saint-Michel",      "Chateau de Chenonceau",       "~250 km",  "~2h45"),
        (4,  "Chateau de Chenonceau",     "Chateau de Chambord",         "~40 km",   "~35 min"),
        (5,  "Chateau de Chambord",       "Puy du Fou",                  "~150 km",  "~1h45"),
        (6,  "Puy du Fou",                "La Roque-Gageac",             "~280 km",  "~3h00"),
        (7,  "La Roque-Gageac",           "Beynac-et-Cazenac",           "~10 km",   "~15 min"),
        (8,  "Beynac-et-Cazenac",         "Rocamadour",                  "~55 km",   "~1h00"),
        (9,  "Rocamadour",                "Gouffre de Padirac",          "~30 km",   "~30 min"),
        (10, "Gouffre de Padirac",        "Ocres de Roussillon",         "~350 km",  "~3h30"),
        (11, "Ocres de Roussillon",       "Les Baux-de-Provence",        "~80 km",   "~1h15"),
        (12, "Les Baux-de-Provence",      "Calanques de Cassis",         "~80 km",   "~1h00"),
        (13, "Calanques de Cassis",       "Lac d'Annecy",                "~300 km",  "~3h00"),
        (14, "Lac d'Annecy",              "Cirque du Fer a Cheval",      "~50 km",   "~50 min"),
        (15, "Cirque du Fer a Cheval",    "Aiguille du Midi - Chamonix", "~45 km",   "~50 min"),
        (16, "QC Terme Chamonix",         "Paris (retour)",              "~580 km",  "~5h30"),
    ]

    # alternance de couleurs par jour
    jour_colors = {
        (1, 2): "DCE6F1",
        (3, 4, 5): "FCE4D6",
        (6, 7, 8, 9): "E2EFDA",
        (10, 11, 12): "FFF2CC",
        (13, 14, 15, 16): "D9E1F2",
    }

    def _etape_color(num):
        for keys, color in jour_colors.items():
            if num in keys:
                return color
        return "F2F2F2"

    for row_idx, (num, depart, arrivee, dist, duree) in enumerate(etapes, start=3):
        bg = _etape_color(num)
        ws.row_dimensions[row_idx].height = 20
        for col_idx, val in enumerate([num, depart, arrivee, dist, duree, ""], start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.fill = _fill(bg)
            cell.font = _font(size=10)
            cell.border = _border()
            cell.alignment = _align(h="center" if col_idx in (1, 4, 5) else "left",
                                    wrap=False)

    # ligne total
    total_row = len(etapes) + 3
    ws.row_dimensions[total_row].height = 22
    ws.merge_cells(f"A{total_row}:C{total_row}")
    t_cell = ws.cell(row=total_row, column=1, value="🏁  TOTAL DU VOYAGE")
    t_cell.font = _font(bold=True, size=11, color="FFFFFF")
    t_cell.fill = _fill("1F3864")
    t_cell.alignment = _align(h="center")
    t_cell.border = _border()

    for col_idx, val in enumerate(["~2 700 km", "~32h de route", ""], start=4):
        cell = ws.cell(row=total_row, column=col_idx, value=val)
        cell.font = _font(bold=True, size=11, color="FFFFFF")
        cell.fill = _fill("1F3864")
        cell.alignment = _align(h="center")
        cell.border = _border()

    ws.freeze_panes = "A3"

    return ws


# ---------------------------------------------------------------------------
# ONGLET 3 — PLANNING JOURNALIER
# ---------------------------------------------------------------------------

def build_planning(wb):
    ws = wb.create_sheet("📅 Planning")

    ws.merge_cells("A1:D1")
    t = ws["A1"]
    t.value = "📅  PLANNING JOURNALIER DÉTAILLÉ"
    t.font = _font(bold=True, size=14, color="FFFFFF")
    t.fill = _fill("1F3864")
    t.alignment = _align(h="center")
    ws.row_dimensions[1].height = 30

    col_widths_p = [20, 35, 35, 35]
    for col_idx, w in enumerate(col_widths_p, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    headers_p = ["LIEU", "🌅 MATIN", "☀️ APRÈS-MIDI", "🌙 SOIR / HÉBERGEMENT"]
    for col_idx, h in enumerate(headers_p, start=1):
        cell = ws.cell(row=2, column=col_idx, value=h)
        cell.font = _font(bold=True, size=11, color="FFFFFF")
        cell.fill = _fill("2E4D8A")
        cell.alignment = _align(h="center")
        cell.border = _border()
    ws.row_dimensions[2].height = 20

    current_jour = None
    row = 3
    for item in TRIP_DATA:
        jour = item["jour"]
        bg = item.get("couleur_jour", "FFFFFF")
        text_color = "FFFFFF" if bg not in ("FFFFFF", "F2F2F2") else "000000"

        if jour != current_jour:
            # sous-titre du jour
            ws.merge_cells(f"A{row}:D{row}")
            jour_cell = ws.cell(row=row, column=1, value=f"── {jour} ──")
            jour_cell.font = _font(bold=True, size=12, color="FFFFFF")
            jour_cell.fill = _fill(bg)
            jour_cell.alignment = _align(h="center")
            ws.row_dimensions[row].height = 20
            row += 1
            current_jour = jour

        hebergement_val = item.get("hebergement", "")
        soir_val = item.get("soir", "")
        soir_display = soir_val
        if hebergement_val:
            soir_display = f"{soir_val}\n🏨 {hebergement_val}" if soir_val else f"🏨 {hebergement_val}"

        ws.row_dimensions[row].height = 60
        for col_idx, val in enumerate([
            item["lieu"],
            item.get("matin", ""),
            item.get("aprem", ""),
            soir_display,
        ], start=1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            cell.fill = _fill(bg)
            cell.font = _font(size=10, color=text_color)
            cell.alignment = _align(wrap=True)
            cell.border = _border()
            if col_idx == 1:
                cell.font = _font(bold=True, size=10, color=text_color)

        row += 1

    ws.freeze_panes = "A3"
    return ws


# ---------------------------------------------------------------------------
# ONGLET 4 — CHECKLIST
# ---------------------------------------------------------------------------

def build_checklist(wb):
    ws = wb.create_sheet("✅ Checklist")

    ws.merge_cells("A1:C1")
    t = ws["A1"]
    t.value = "✅  CHECKLIST ROAD TRIP FRANCE"
    t.font = _font(bold=True, size=14, color="FFFFFF")
    t.fill = _fill("1F3864")
    t.alignment = _align(h="center")
    ws.row_dimensions[1].height = 30

    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 20

    categories = [
        ("🧳 BAGAGES", [
            "Valise / sac de voyage",
            "Vêtements adaptés (chaud + léger)",
            "Chaussures de randonnée",
            "Maillot de bain (calanques / Annecy)",
            "Veste imperméable",
            "Crème solaire SPF 50+",
        ]),
        ("📋 RÉSERVATIONS", [
            "Puy du Fou — billets Cinéscénie (OBLIGATOIRE)",
            "Aiguille du Midi — téléphérique (OBLIGATOIRE)",
            "QC Terme Chamonix — spa (OBLIGATOIRE)",
            "Gouffre de Padirac — en ligne en été",
            "Hôtel Mont-Saint-Michel",
            "Hôtel Puy du Fou",
            "Hôtel Rocamadour",
            "Hôtel Cassis",
            "QC Terme Chamonix (hébergement)",
        ]),
        ("🚗 VOITURE", [
            "Plein d'essence avant départ",
            "GPS / Google Maps téléchargé hors-ligne",
            "Vignette Crit'Air si zone ZFE",
            "Carte routière papier (backup)",
        ]),
        ("📱 APPS UTILES", [
            "Google Maps — itinéraire",
            "Viamichelin — trafic",
            "Météo France / Météo-Consult",
            "Réservations : Puy du Fou app",
        ]),
    ]

    row = 2
    for cat_title, items in categories:
        ws.merge_cells(f"A{row}:C{row}")
        cat_cell = ws.cell(row=row, column=1, value=cat_title)
        cat_cell.font = _font(bold=True, size=12, color="FFFFFF")
        cat_cell.fill = _fill("2E4D8A")
        cat_cell.alignment = _align(h="left")
        ws.row_dimensions[row].height = 20
        row += 1

        for item_text in items:
            ws.row_dimensions[row].height = 18
            check_cell = ws.cell(row=row, column=1, value="☐")
            check_cell.font = _font(size=12)
            check_cell.alignment = _align(h="center")
            check_cell.border = _border()
            check_cell.fill = _fill("F9F9F9")

            text_cell = ws.cell(row=row, column=2, value=item_text)
            text_cell.font = _font(size=11)
            text_cell.alignment = _align(wrap=False)
            text_cell.border = _border()
            text_cell.fill = _fill("F9F9F9")

            note_cell = ws.cell(row=row, column=3, value="")
            note_cell.border = _border()
            note_cell.fill = _fill("F9F9F9")
            row += 1

        row += 1  # blank line between sections

    return ws


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    output_file = "Road_Trip_5J_France.xlsx"
    print("🚗  Génération du fichier Excel Road Trip France …")

    wb = Workbook()
    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    print("  ► Onglet Itinéraire …")
    build_itineraire(wb)

    print("  ► Onglet Trajet …")
    build_trajet(wb)

    print("  ► Onglet Planning …")
    build_planning(wb)

    print("  ► Onglet Checklist …")
    build_checklist(wb)

    wb.save(output_file)
    print(f"\n✅  Fichier généré : {output_file}")
    print(f"   ({os.path.getsize(output_file) // 1024} Ko)")


if __name__ == "__main__":
    main()
