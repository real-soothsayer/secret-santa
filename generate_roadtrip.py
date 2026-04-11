#!/usr/bin/env python3
"""
generate_roadtrip.py
====================
Génère un fichier Excel premium «Road_Trip_Vendeur.xlsx» pour un road trip
en France sur un itinéraire précis de 17 étapes.

Feuilles générées :
  1. Itinéraire complet — images, prix, logement, horaires, restos, routes, ambiance
  2. Planning jour par jour — heures, activités, arrêts, liaisons
  3. Carte du trajet — distances, temps, points de vue
  4. Checklist voyage
  5. Panorama — mosaïque d'images des étapes

Dépendances : openpyxl, requests, Pillow  (voir requirements.txt)
"""

import io
import os
import sys
import time
import urllib.parse
from pathlib import Path

import requests
from PIL import Image
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import (
    Alignment, Border, Font, GradientFill, PatternFill, Side
)
from openpyxl.styles.numbers import FORMAT_TEXT
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# PALETTE / THEME
# ---------------------------------------------------------------------------
C_HEADER_BG   = "0D1B2A"   # bleu nuit profond
C_HEADER_FG   = "F5E6CA"   # beige/or
C_ACCENT_GOLD = "C9A84C"   # or
C_ACCENT_CORAL= "E07B54"   # corail
C_ROW_ODD     = "EBF2FA"   # bleu très clair
C_ROW_EVEN    = "FFFFFF"   # blanc
C_SECTION_BG  = "1B2B3A"   # bleu nuit section
C_TEXT_DARK   = "0D1B2A"
C_TEXT_MID    = "3A4A5A"
C_BORDER      = "8FA8C0"

IMG_W, IMG_H = 220, 120      # pixels — taille cible images

CACHE_DIR = Path("/tmp/roadtrip_img_cache")
CACHE_DIR.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# DONNÉES DES ÉTAPES
# ---------------------------------------------------------------------------
ETAPES = [
    {
        "id": 1,
        "nom": "Château de Chambord",
        "emoji": "🏰",
        "region": "Centre-Val de Loire",
        "ambiance": "🌲 Forêt & Renaissance",
        "prix_entree": "14,50 €",
        "horaires": "9h–18h (juil-août jusqu'à 21h)",
        "logement": "Hôtel du Grand Saint-Michel ★★★ — ~120 €/nuit",
        "logement_url": "https://www.booking.com/hotel/fr/du-grand-saint-michel-chambord.fr.html",
        "maps_url": "https://maps.google.com/?q=Château+de+Chambord",
        "resto": "Le Chambord (terrasse château) — formule 28 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Chambord",
        "routes": "D112 & D33 depuis Blois — bords de Loire panoramiques",
        "conseils_photo": "Lever de soleil sur le donjon, reflets dans les douves",
        "points_vue": "Terrasse du donjon (vue 360°), Allée Royale",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c5/Chambord_Castle_Northwest_facade.jpg/320px-Chambord_Castle_Northwest_facade.jpg",
    },
    {
        "id": 2,
        "nom": "Château de Chenonceau",
        "emoji": "🌹",
        "region": "Centre-Val de Loire",
        "ambiance": "🌊 Rivière & Jardins",
        "prix_entree": "17 €",
        "horaires": "9h–18h30",
        "logement": "La Roseraie (chambres d'hôtes) — ~100 €/nuit",
        "logement_url": "https://www.booking.com/searchresults.fr.html?ss=Chenonceau",
        "maps_url": "https://maps.google.com/?q=Château+de+Chenonceau",
        "resto": "L'Orangerie du Château — menu 35 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Chenonceau",
        "routes": "D40 longeant le Cher — route vallonnée et boisée",
        "conseils_photo": "Vue depuis le pont sur le Cher au coucher du soleil",
        "points_vue": "Jardin Diana de Poitiers, pont sur le Cher",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/a/af/Château_de_Chenonceau_2013.jpg/320px-Château_de_Chenonceau_2013.jpg",
    },
    {
        "id": 3,
        "nom": "ZooParc de Beauval",
        "emoji": "🐼",
        "region": "Centre-Val de Loire",
        "ambiance": "🌿 Jungle & Savane",
        "prix_entree": "35 €",
        "horaires": "9h–18h30 (19h30 été)",
        "logement": "Val de Loire Village (hôtel du parc) — ~150 €/nuit",
        "logement_url": "https://www.booking.com/hotel/fr/val-de-loire-village.fr.html",
        "maps_url": "https://maps.google.com/?q=ZooParc+de+Beauval",
        "resto": "Le Comptoir des Savanes (dans le zoo) — ~20 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Saint-Aignan-sur-Cher",
        "routes": "A85 puis D675 — plaines agricoles ouvertes",
        "conseils_photo": "Pandas géants au lever, gorilles à l'heure du repas",
        "points_vue": "Passerelle tropicale, grande volière",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/0/0b/Beauval_%281%29.JPG/320px-Beauval_%281%29.JPG",
    },
    {
        "id": 4,
        "nom": "Futuroscope",
        "emoji": "🚀",
        "region": "Nouvelle-Aquitaine (Vienne)",
        "ambiance": "🌆 Urbain & Futuriste",
        "prix_entree": "50 € (journée)",
        "horaires": "10h–22h (spectacle nocturne inclus)",
        "logement": "Hôtel des Lumières ★★★ (sur site) — ~130 €/nuit",
        "logement_url": "https://www.futuroscope.com/fr/hotels",
        "maps_url": "https://maps.google.com/?q=Futuroscope+Poitiers",
        "resto": "L'Espace Découverte (buffet dans le parc) — ~18 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Futuroscope+Poitiers",
        "routes": "A10 depuis Tours — autoroute rapide, 1h45",
        "conseils_photo": "Dômes illuminés la nuit, fontaines en soirée",
        "points_vue": "Gyrotour (panorama sur la région), spectacle nocturne",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e3/Futuroscope_-_Pavillon_de_la_Vienne_%282012%29.jpg/320px-Futuroscope_-_Pavillon_de_la_Vienne_%282012%29.jpg",
    },
    {
        "id": 5,
        "nom": "Lagon du Bois de Saint-Pierre",
        "emoji": "🏊",
        "region": "Nouvelle-Aquitaine (Charente)",
        "ambiance": "🌲 Forêt & Lagon turquoise",
        "prix_entree": "Gratuit / Parking 3 €",
        "horaires": "Accès libre (baignade juin–sept)",
        "logement": "Camping La Garenne ★★★ — ~35 €/nuit",
        "logement_url": "https://www.booking.com/searchresults.fr.html?ss=Ansac-sur-Vienne",
        "maps_url": "https://maps.google.com/?q=Lagon+Bois+Saint-Pierre+Charente",
        "resto": "Les Pieds dans l'Eau (snack sur place été) — ~12 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Ansac-sur-Vienne",
        "routes": "D951 à travers bocage charentais — route calme et verte",
        "conseils_photo": "Eau turquoise en lumière rasante tôt le matin",
        "points_vue": "Belvédère du lagon, sentier forestier autour",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/b/b0/Charente_Riviere.jpg/320px-Charente_Riviere.jpg",
    },
    {
        "id": 6,
        "nom": "Puy du Fou",
        "emoji": "⚔️",
        "region": "Pays de la Loire (Vendée)",
        "ambiance": "🏰 Médiéval & Épique",
        "prix_entree": "42 € (journée)",
        "horaires": "10h–23h30 (Grand Parc + soirées)",
        "logement": "Le Logis de Lescure (hôtel du parc) — ~180 €/nuit",
        "logement_url": "https://www.puydufou.com/fr/hebergements",
        "maps_url": "https://maps.google.com/?q=Puy+du+Fou+Vendée",
        "resto": "Auberge du Moulin Fondu (reconstitution médiévale) — ~22 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Puy+du+Fou",
        "routes": "D160 bocage vendéen — petites routes rurales charmantes",
        "conseils_photo": "Cinéscénie de nuit, chevaliers en plein air",
        "points_vue": "Lac du Grand Parc, village gaulois",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/3/33/PuyDuFou2.jpg/320px-PuyDuFou2.jpg",
    },
    {
        "id": 7,
        "nom": "Dune du Pilat",
        "emoji": "🏖️",
        "region": "Nouvelle-Aquitaine (Gironde)",
        "ambiance": "🌊 Océan & Sable doré",
        "prix_entree": "Gratuit (parking 8 €)",
        "horaires": "Accès libre toute l'année",
        "logement": "Hôtel La Co(o)rniche ★★★★ — ~220 €/nuit",
        "logement_url": "https://www.booking.com/hotel/fr/la-co-o-rniche.fr.html",
        "maps_url": "https://maps.google.com/?q=Dune+du+Pilat",
        "resto": "Restaurant La Guinguette (vue mer) — ~30 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Dune+du+Pilat",
        "routes": "D218 longeant le Bassin d'Arcachon — panoramique",
        "conseils_photo": "Sommet dune au coucher : forêt, dune et océan simultanément",
        "points_vue": "Sommet de la dune (vue 360° Arcachon/forêt/Atlantique)",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/3/3d/DunePilat.jpg/320px-DunePilat.jpg",
    },
    {
        "id": 8,
        "nom": "La Roque-Gageac",
        "emoji": "🛶",
        "region": "Nouvelle-Aquitaine (Dordogne)",
        "ambiance": "🌊 Rivière & Falaises ocres",
        "prix_entree": "Gratuit (village) / Canöé 15 €",
        "horaires": "Village accessible toute l'année",
        "logement": "Hôtel Belle Etoile ★★★ — ~110 €/nuit",
        "logement_url": "https://www.booking.com/hotel/fr/belle-etoile-la-roque-gageac.fr.html",
        "maps_url": "https://maps.google.com/?q=La+Roque-Gageac",
        "resto": "La Belle Etoile (terrasse Dordogne) — formule 28 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+La+Roque-Gageac",
        "routes": "D703 longer la Dordogne — une des plus belles routes de France",
        "conseils_photo": "Village en reflet dans la Dordogne à l'aube",
        "points_vue": "Fort troglodyte, promenade en gabarre sur la Dordogne",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/e/e7/La_Roque-Gageac_6.jpg/320px-La_Roque-Gageac_6.jpg",
    },
    {
        "id": 9,
        "nom": "Beynac-et-Cazenac",
        "emoji": "🏯",
        "region": "Nouvelle-Aquitaine (Dordogne)",
        "ambiance": "🌿 Vallée verdoyante & Château perché",
        "prix_entree": "8 € (château)",
        "horaires": "10h–18h30",
        "logement": "Maison d'hôtes Domaine de Tayac — ~95 €/nuit",
        "logement_url": "https://www.booking.com/searchresults.fr.html?ss=Beynac-et-Cazenac",
        "maps_url": "https://maps.google.com/?q=Beynac-et-Cazenac",
        "resto": "Café de la Rivière — omelette périgourdine ~15 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Beynac+Cazenac",
        "routes": "D703 — 4 km de La Roque-Gageac, longer la rivière",
        "conseils_photo": "Vue du château sur la Dordogne et les méandres",
        "points_vue": "Chemin de ronde du château (panorama Dordogne)",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/6/64/Beynac-et-Cazenac.jpg/320px-Beynac-et-Cazenac.jpg",
    },
    {
        "id": 10,
        "nom": "Rocamadour",
        "emoji": "⛪",
        "region": "Occitanie (Lot)",
        "ambiance": "🏔️ Falaises & Spiritualité",
        "prix_entree": "9 € (cité religieuse + château)",
        "horaires": "9h–19h",
        "logement": "Hôtel Beau Site ★★★ — ~105 €/nuit",
        "logement_url": "https://www.booking.com/hotel/fr/beau-site-rocamadour.fr.html",
        "maps_url": "https://maps.google.com/?q=Rocamadour",
        "resto": "Le Roc du Berger (spécialités du Quercy) — ~25 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Rocamadour",
        "routes": "D673 depuis Périgord — gorges de l'Alzou sur la fin",
        "conseils_photo": "Ensemble depuis L'Hospitalet en fin de journée",
        "points_vue": "Belvédère de L'Hospitalet, escalier des pèlerins",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/1/11/Rocamadour-panorama.jpg/320px-Rocamadour-panorama.jpg",
    },
    {
        "id": 11,
        "nom": "Gouffre de Padirac",
        "emoji": "🕳️",
        "region": "Occitanie (Lot)",
        "ambiance": "🌊 Souterrain & Mystère",
        "prix_entree": "16 € (adulte)",
        "horaires": "9h30–17h (juil-août 9h–19h)",
        "logement": "Camping Les Chênes (près de Rocamadour) — ~30 €/nuit",
        "logement_url": "https://www.booking.com/searchresults.fr.html?ss=Padirac",
        "maps_url": "https://maps.google.com/?q=Gouffre+de+Padirac",
        "resto": "Le Buffet du Gouffre (sur site) — ~20 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Padirac",
        "routes": "D90 — causses du Lot, paysage de plateaux calcaires",
        "conseils_photo": "Lac souterrain au fond du gouffre (appareil étanche conseillé)",
        "points_vue": "Bord du gouffre (75 m de profondeur), chapelle souterraine",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/5/5e/Padirac_04.jpg/320px-Padirac_04.jpg",
    },
    {
        "id": 12,
        "nom": "Parc naturel régional des Volcans d'Auvergne",
        "emoji": "🌋",
        "region": "Auvergne-Rhône-Alpes",
        "ambiance": "🌿 Volcans & Prairies d'altitude",
        "prix_entree": "Gratuit / Télécabine Puy de Dôme 19 €",
        "horaires": "Télécabine : 7h–21h (été)",
        "logement": "La Bergerie de Sarpoil ★★ (Issoire) — ~90 €/nuit",
        "logement_url": "https://www.booking.com/searchresults.fr.html?ss=Clermont-Ferrand+volcans",
        "maps_url": "https://maps.google.com/?q=Puy+de+Dôme+Auvergne",
        "resto": "Crêperie La Bourbonnaise (Clermont-Fd) — galette 12 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Clermont-Ferrand",
        "routes": "D941A & D941 — Route des Puys, panoramas volcaniques",
        "conseils_photo": "Sommet Puy de Dôme en début de matinée (moins de brume)",
        "points_vue": "Puy de Dôme (1465 m), Cheire d'Aydat, Lac Chambon",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/b/ba/Puy_de_d%C3%B4me_depuis_la_route_de_Laschamps.jpg/320px-Puy_de_d%C3%B4me_depuis_la_route_de_Laschamps.jpg",
    },
    {
        "id": 13,
        "nom": "Gorges de la Sioule",
        "emoji": "🏞️",
        "region": "Auvergne-Rhône-Alpes (Allier)",
        "ambiance": "🌊 Gorges sauvages & Nature",
        "prix_entree": "Gratuit",
        "horaires": "Accès libre toute l'année",
        "logement": "Gîte Le Moulin de la Sioule — ~75 €/nuit",
        "logement_url": "https://www.booking.com/searchresults.fr.html?ss=Gorges+Sioule",
        "maps_url": "https://maps.google.com/?q=Gorges+de+la+Sioule",
        "resto": "Café du Pont de Menat — assiette campagnarde ~14 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Menat+Sioule",
        "routes": "D916 Menat–Châteauneuf-les-Bains — route sauvage en corniche",
        "conseils_photo": "Belvédère de la Roche Ronde au lever de soleil",
        "points_vue": "Viaduc de Pontgibaud, Pont de Menat, sentier des Fées",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/7/7b/Gorges_de_la_Sioule.JPG/320px-Gorges_de_la_Sioule.JPG",
    },
    {
        "id": 14,
        "nom": "Montluçon",
        "emoji": "🎶",
        "region": "Auvergne-Rhône-Alpes (Allier)",
        "ambiance": "🏙️ Vieille ville & Musique",
        "prix_entree": "Gratuit (vieux Montluçon)",
        "horaires": "Vieille ville accessible toute la journée",
        "logement": "Hôtel des Bourbons ★★★ — ~85 €/nuit",
        "logement_url": "https://www.booking.com/searchresults.fr.html?ss=Montluçon",
        "maps_url": "https://maps.google.com/?q=Montluçon",
        "resto": "Le Grand Café Montluçon — plat du jour 13 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Montluçon",
        "routes": "D2144 — Allier & bocage bourbonnais",
        "conseils_photo": "Remparts médiévaux éclairés la nuit",
        "points_vue": "Château des Ducs de Bourbon, vieux pont sur le Cher",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/b/b5/Montlu%C3%A7on_-_Ch%C3%A2teau_des_Bourbons_%28n%C3%B4t%C3%A9%29.jpg/320px-Montlu%C3%A7on_-_Ch%C3%A2teau_des_Bourbons_%28n%C3%B4t%C3%A9%29.jpg",
    },
    {
        "id": 15,
        "nom": "Bourges — Cathédrale Saint-Étienne",
        "emoji": "🕍",
        "region": "Centre-Val de Loire (Cher)",
        "ambiance": "🏙️ Médiéval & Gothique",
        "prix_entree": "Gratuit (cathédrale) / Tour 7 €",
        "horaires": "Cathédrale 8h30–18h15",
        "logement": "Hôtel de Bourbon ★★★★ — ~120 €/nuit",
        "logement_url": "https://www.booking.com/hotel/fr/de-bourbon-bourges.fr.html",
        "maps_url": "https://maps.google.com/?q=Cathédrale+Saint-Étienne+Bourges",
        "resto": "L'Alcazar (cuisine du Berry) — menu 28 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Bourges",
        "routes": "A71 — rapide ou D940 par campagne berrichonne",
        "conseils_photo": "Portails sculptés, vitraux en lumière du matin",
        "points_vue": "Tour Nord (panorama ville), jardins de l'Archevêché",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/5/51/Bourges_-_Cath%C3%A9drale_Saint-%C3%89tienne_%28880%29.jpg/320px-Bourges_-_Cath%C3%A9drale_Saint-%C3%89tienne_%28880%29.jpg",
    },
    {
        "id": 16,
        "nom": "Orléans",
        "emoji": "⚜️",
        "region": "Centre-Val de Loire (Loiret)",
        "ambiance": "🌊 Loire & Histoire",
        "prix_entree": "Gratuit (cathédrale)",
        "horaires": "Cathédrale 8h30–18h",
        "logement": "Hôtel Mercure Orléans Centre ★★★★ — ~110 €/nuit",
        "logement_url": "https://www.booking.com/hotel/fr/mercure-orleans-centre.fr.html",
        "maps_url": "https://maps.google.com/?q=Orléans+cathédrale",
        "resto": "La Chancellerie (place du Martroi) — menu 30 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+Orléans+centre",
        "routes": "A71 puis N152 quais de Loire — promenade bords de Loire",
        "conseils_photo": "Cathédrale et quais de Loire au crépuscule",
        "points_vue": "Pont George-V, quais de Loire, Maison de Jeanne d'Arc",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/2/28/Orleanscathedral.jpg/320px-Orleanscathedral.jpg",
    },
    {
        "id": 17,
        "nom": "France Miniature",
        "emoji": "🗺️",
        "region": "Île-de-France (Yvelines)",
        "ambiance": "🌟 Ludique & Culturel",
        "prix_entree": "22 € (adulte)",
        "horaires": "10h–18h (mars–novembre)",
        "logement": "Ibis Styles Élancourt ★★★ — ~95 €/nuit",
        "logement_url": "https://www.booking.com/searchresults.fr.html?ss=Élancourt",
        "maps_url": "https://maps.google.com/?q=France+Miniature+Élancourt",
        "resto": "Le Bistrot Miniature (sur site) — formule 16 €",
        "resto_url": "https://www.google.com/maps/search/restaurant+France+Miniature+Elancourt",
        "routes": "A12/A13 depuis Orléans — 1h30 d'autoroute, fin de boucle vers Paris",
        "conseils_photo": "Vue plongeante sur les monuments en maquette",
        "points_vue": "Tour Eiffel miniature, Mont-Saint-Michel en modèle",
        "wikimedia_img": "https://upload.wikimedia.org/wikipedia/commons/thumb/f/fb/FranceMiniature.jpg/320px-FranceMiniature.jpg",
    },
]

# ---------------------------------------------------------------------------
# DONNÉES PLANNING JOUR PAR JOUR
# ---------------------------------------------------------------------------
PLANNING = [
    {"jour": "Jour 1", "date": "J+0",
     "matin": "Départ tôt vers Chambord (2h depuis Paris)",
     "visite1": "9h30 — Château de Chambord (3h) — terrasse et parc",
     "dejeuner": "12h30 — Le Chambord (château) ou pique-nique dans le parc",
     "apresmidi": "14h30 — Route D40 vers Chenonceau (45 min)",
     "visite2": "15h30 — Château de Chenonceau (2h)",
     "diner": "19h — L'Orangerie du Château ou restaurant à Amboise",
     "nuit": "🛏️ Amboise ou Chenonceau (~100 €)",
     "km": "~70 km", "temps_route": "~1h30"},
    {"jour": "Jour 2", "date": "J+1",
     "matin": "9h — ZooParc de Beauval (journée entière recommandée)",
     "visite1": "9h–12h30 — Pandas, gorilles, éléphants",
     "dejeuner": "12h30 — Comptoir des Savanes (dans le zoo)",
     "apresmidi": "14h–17h — Serre tropicale et spectacles",
     "diner": "19h — Retour à l'hôtel, dîner à Saint-Aignan",
     "nuit": "🛏️ Val de Loire Village ou Gîte (~130 €)",
     "km": "~30 km (de Chenonceau)", "temps_route": "~40 min"},
    {"jour": "Jour 3", "date": "J+2",
     "matin": "9h — Départ vers Futuroscope (1h30)",
     "visite1": "10h30 — Futuroscope (journée + spectacle nocturne)",
     "dejeuner": "12h30 — Espace Découverte (buffet dans le parc)",
     "apresmidi": "14h–21h — Attractions, simulateurs, 4D",
     "diner": "Dîner dans le parc ou hôtel du site",
     "nuit": "🛏️ Hôtel des Lumières (sur site) — ~130 €",
     "km": "~120 km (depuis Beauval)", "temps_route": "~1h30"},
    {"jour": "Jour 4", "date": "J+3",
     "matin": "9h — Lagon du Bois de Saint-Pierre (1h de Futuroscope)",
     "visite1": "10h–12h — Baignade et balade forestière",
     "dejeuner": "12h30 — Pique-nique au bord du lagon",
     "apresmidi": "14h — Route vers Puy du Fou (2h)",
     "visite2": "16h–23h30 — Puy du Fou (après-midi + Cinéscénie nocturne)",
     "diner": "Auberge du Moulin Fondu dans le parc",
     "nuit": "🛏️ Logis de Lescure (Puy du Fou) — ~180 €",
     "km": "~160 km", "temps_route": "~2h"},
    {"jour": "Jour 5", "date": "J+4",
     "matin": "9h — Départ vers Dune du Pilat (2h)",
     "visite1": "11h — Ascension Dune du Pilat (1h30) + baignade",
     "dejeuner": "13h — La Guinguette (vue mer)",
     "apresmidi": "15h — Route D703 vers Dordogne (2h30)",
     "visite2": "18h — La Roque-Gageac au coucher du soleil",
     "diner": "19h30 — La Belle Etoile (terrasse Dordogne)",
     "nuit": "🛏️ Hôtel Belle Etoile ou gîte — ~110 €",
     "km": "~270 km", "temps_route": "~3h"},
    {"jour": "Jour 6", "date": "J+5",
     "matin": "9h — Beynac-et-Cazenac (15 min de La Roque-Gageac)",
     "visite1": "9h30–11h30 — Château de Beynac (vue panoramique)",
     "dejeuner": "12h — Café de la Rivière (omelette périgourdine)",
     "apresmidi": "14h — Route vers Rocamadour (1h30)",
     "visite2": "15h30–18h — Rocamadour (cité religieuse + château)",
     "diner": "19h30 — Le Roc du Berger (spécialités Quercy)",
     "nuit": "🛏️ Hôtel Beau Site Rocamadour — ~105 €",
     "km": "~90 km", "temps_route": "~1h30"},
    {"jour": "Jour 7", "date": "J+6",
     "matin": "9h — Gouffre de Padirac (30 min de Rocamadour)",
     "visite1": "9h30–12h — Visite souterraine en barque",
     "dejeuner": "12h30 — Buffet du Gouffre (sur site)",
     "apresmidi": "14h — Route vers Volcans d'Auvergne (2h30)",
     "visite2": "16h30 — Montée Puy de Dôme en télécabine",
     "diner": "Crêperie à Clermont-Ferrand",
     "nuit": "🛏️ Bergerie de Sarpoil ou Clermont — ~90 €",
     "km": "~200 km", "temps_route": "~2h30"},
    {"jour": "Jour 8", "date": "J+7",
     "matin": "9h — Gorges de la Sioule (1h de Clermont)",
     "visite1": "10h–12h30 — Sentier des gorges, belvédère Roche Ronde",
     "dejeuner": "12h30 — Café du Pont de Menat",
     "apresmidi": "14h30 — Route vers Montluçon (45 min)",
     "visite2": "15h30–17h — Vieux Montluçon, château des Bourbons",
     "diner": "19h — Grand Café Montluçon",
     "nuit": "🛏️ Hôtel des Bourbons — ~85 €",
     "km": "~80 km", "temps_route": "~1h30"},
    {"jour": "Jour 9", "date": "J+8",
     "matin": "9h30 — Départ vers Bourges (1h)",
     "visite1": "11h–13h — Cathédrale Saint-Étienne (UNESCO) + Tour Nord",
     "dejeuner": "13h — L'Alcazar (cuisine du Berry)",
     "apresmidi": "15h — Palais Jacques-Cœur + jardins de l'Archevêché",
     "visite2": "17h — Départ vers Orléans (1h)",
     "diner": "19h30 — La Chancellerie (place du Martroi, Orléans)",
     "nuit": "🛏️ Mercure Orléans Centre — ~110 €",
     "km": "~140 km", "temps_route": "~2h"},
    {"jour": "Jour 10", "date": "J+9",
     "matin": "9h — Promenade bords de Loire, Cathédrale d'Orléans",
     "visite1": "10h–12h — Maison de Jeanne d'Arc + cathédrale",
     "dejeuner": "12h30 — Brasserie sur les quais de Loire",
     "apresmidi": "14h — Route vers France Miniature (1h30)",
     "visite2": "15h30–18h — France Miniature (fin de boucle !)",
     "diner": "19h — Retour vers Paris (30 min)",
     "nuit": "🏠 Retour à la maison",
     "km": "~140 km", "temps_route": "~1h30"},
]

# ---------------------------------------------------------------------------
# DONNÉES CARTE / DISTANCES
# ---------------------------------------------------------------------------
CARTE = [
    ("Départ Paris", "Château de Chambord", "180 km", "1h50", "A10 → Blois", "Bords de Loire (A10)", "Vue sur la Loire à Blois"),
    ("Château de Chambord", "Château de Chenonceau", "45 km", "0h50", "D112 → D40", "Route forestière & vallée du Cher", "Jardins de Chaumont-sur-Loire"),
    ("Château de Chenonceau", "ZooParc de Beauval", "25 km", "0h30", "D176", "Vallée du Cher", "Réserve naturelle de Beauval"),
    ("ZooParc de Beauval", "Futuroscope", "130 km", "1h30", "A85 → A10", "Autoroute rapide", "Viaduc de Mirepoix (panorama)"),
    ("Futuroscope", "Lagon Saint-Pierre", "55 km", "0h55", "D951", "Bocage charentais", "Forêt et rivière Vienne"),
    ("Lagon Saint-Pierre", "Puy du Fou", "115 km", "1h30", "D160 via Parthenay", "Bocage vendéen", "Abbaye de Parthenay-le-Vieux"),
    ("Puy du Fou", "Dune du Pilat", "185 km", "2h10", "D938 → A89 → A63", "Bandes de vignes Médoc", "Bassin d'Arcachon en approche"),
    ("Dune du Pilat", "La Roque-Gageac", "215 km", "2h30", "A63 → D703", "Forêts des Landes puis vallée Dordogne", "Méandre de Trémolat (D29E4)"),
    ("La Roque-Gageac", "Beynac-et-Cazenac", "4 km", "0h08", "D703", "Bord de Dordogne", "Village de Domme (D46)"),
    ("Beynac-et-Cazenac", "Rocamadour", "80 km", "1h15", "D673", "Causses du Quercy", "Gouffre de Presque (en chemin)"),
    ("Rocamadour", "Gouffre de Padirac", "17 km", "0h20", "D90", "Plateau calcaire (causses)", "Village de Gramat"),
    ("Gouffre de Padirac", "Volcans d'Auvergne", "195 km", "2h30", "D840 → A75", "Plateaux du Massif Central", "Viaduc de Garabit (A75)"),
    ("Volcans d'Auvergne", "Gorges de la Sioule", "55 km", "0h55", "D916", "Route des Puys", "Châteauneuf-les-Bains"),
    ("Gorges de la Sioule", "Montluçon", "55 km", "0h55", "N144", "Bocage bourbonnais", "Forêt de Tronçais (proche)"),
    ("Montluçon", "Bourges", "95 km", "1h05", "D2144 → D940", "Campagne berrichonne", "Abbaye de Noirlac (D30)"),
    ("Bourges", "Orléans", "100 km", "1h05", "A71", "Sologne boisée", "Château de La Ferté (D17)"),
    ("Orléans", "France Miniature", "125 km", "1h25", "A10 → A12", "Val de Loire puis Île-de-France", "Chartres (détour 30 min)"),
    ("France Miniature", "Retour Paris", "40 km", "0h40", "A13", "Île-de-France", "Versailles (5 min de détour)"),
]

# ---------------------------------------------------------------------------
# CHECKLIST VOYAGE
# ---------------------------------------------------------------------------
CHECKLIST = {
    "📋 Documents & Administratif": [
        "Carte d'identité ou passeport valide",
        "Permis de conduire",
        "Carte Vitale / assurance maladie",
        "Cartes bancaires (2 minimum)",
        "Réservations imprimées (hôtels, parcs, trains)",
        "Carnet d'adresses d'urgence",
        "Attestation d'assurance véhicule",
        "Carte européenne d'assurance maladie",
    ],
    "👕 Vêtements & Bagages": [
        "Vêtements pour 10 jours (rotatif)",
        "Tenue de randonnée (gorges/volcans)",
        "Maillot de bain (lagon, dune, calanques)",
        "Imperméable / coupe-vent",
        "Chaussures de marche",
        "Sandales ou tongs",
        "Casquette / chapeau soleil",
        "Lunettes de soleil",
        "Crème solaire SPF 50",
    ],
    "📷 Équipement Photo & Tech": [
        "Appareil photo + chargeur",
        "Drone (vérifier autorisations zones)",
        "Trépied léger",
        "Batteries de rechange / power bank",
        "Cartes mémoire (32 Go+)",
        "Adaptateur voiture (chargeur USB)",
        "GPS ou téléphone avec offline maps",
        "Écouteurs / AirPods",
    ],
    "💊 Santé & Pharmacie": [
        "Trousse de premiers secours",
        "Médicaments personnels",
        "Antidouleurs (ibuprofène, paracétamol)",
        "Anti-nausées (utile pour routes sinueuses)",
        "Spray anti-moustiques",
        "Pansements et crème antiseptique",
        "Crème anti-ampoules (pieds)",
        "Collyre hydratant",
    ],
    "🚗 Voiture & Route": [
        "Vignette Crit'Air à jour",
        "Triangle de signalisation + gilet",
        "Éthylotest homologué",
        "Sac poubelle voiture",
        "Thermos / gourdes pour la route",
        "Snacks de voyage",
        "Playlist road trip 🎵",
        "Chargeur voiture multi-USB",
        "Couverture de survie",
    ],
    "🏕️ Confort & Loisirs": [
        "Oreiller de voyage",
        "Masque sommeil et bouchons d'oreilles",
        "Livre ou Kindle",
        "Jeux de cartes / dés",
        "Carnets de voyage / journal",
        "Pique-nique kit (nappe, couverts)",
        "Sac à dos journée",
        "Sac imperméable (kayak/lagon)",
    ],
}

# ---------------------------------------------------------------------------
# FONCTIONS UTILITAIRES
# ---------------------------------------------------------------------------

def hex_to_argb(hex_color: str) -> str:
    """Convertit un hex 6-char en ARGB 8-char (opacité FF)."""
    return "FF" + hex_color.upper()


def make_font(name="Calibri", size=11, bold=False, italic=False, color=None):
    kw = dict(name=name, size=size, bold=bold, italic=italic)
    if color:
        kw["color"] = hex_to_argb(color)
    return Font(**kw)


def make_fill(hex_color: str):
    return PatternFill("solid", fgColor=hex_to_argb(hex_color))


def make_border(color=C_BORDER, style="thin"):
    s = Side(style=style, color=hex_to_argb(color))
    return Border(left=s, right=s, top=s, bottom=s)


def make_align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def add_header_row(ws, row_idx: int, values: list, col_start=1):
    """Remplit une ligne d'en-tête avec le style dark."""
    for col, val in enumerate(values, start=col_start):
        cell = ws.cell(row=row_idx, column=col, value=val)
        cell.font = make_font(size=10, bold=True, color=C_HEADER_FG)
        cell.fill = make_fill(C_HEADER_BG)
        cell.alignment = make_align("center")
        cell.border = make_border(C_ACCENT_GOLD, "medium")


def set_row_style(ws, row_idx: int, ncols: int, odd: bool, col_start=1):
    """Style alterné pour les lignes de données."""
    bg = C_ROW_ODD if odd else C_ROW_EVEN
    for col in range(col_start, col_start + ncols):
        cell = ws.cell(row=row_idx, column=col)
        if not cell.value:
            pass
        cell.fill = make_fill(bg)
        cell.border = make_border()
        cell.alignment = make_align()
        if cell.font is None or not cell.font.bold:
            cell.font = make_font(color=C_TEXT_DARK)


# ---------------------------------------------------------------------------
# TÉLÉCHARGEMENT DES IMAGES
# ---------------------------------------------------------------------------

def download_image(url: str, name: str) -> Path | None:
    """Télécharge et redimensionne une image. Retourne le chemin local ou None."""
    cache_path = CACHE_DIR / f"{name}.jpg"
    if cache_path.exists():
        return cache_path
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (compatible; RoadTripBot/1.0; "
            "+https://github.com/real-soothsayer/secret-santa)"
        )
    }
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        img = Image.open(io.BytesIO(resp.content)).convert("RGB")
        img = img.resize((IMG_W, IMG_H), Image.LANCZOS)
        img.save(cache_path, "JPEG", quality=85)
        print(f"  ✅ Image téléchargée : {name}")
        return cache_path
    except Exception as exc:
        print(f"  ⚠️  Image non disponible pour {name} : {exc}")
        return None


def create_placeholder_image(label: str) -> Path:
    """Crée une image placeholder avec texte si téléchargement échoue."""
    cache_path = CACHE_DIR / f"placeholder_{label[:20]}.jpg"
    if cache_path.exists():
        return cache_path
    try:
        from PIL import ImageDraw, ImageFont
        img = Image.new("RGB", (IMG_W, IMG_H), color=(13, 27, 42))
        draw = ImageDraw.Draw(img)
        # Texte centré
        text = f"📍 {label[:25]}"
        bbox = draw.textbbox((0, 0), text, font=None)
        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        x = (IMG_W - w) // 2
        y = (IMG_H - h) // 2
        draw.text((x, y), text, fill=(245, 230, 202))
        img.save(cache_path, "JPEG", quality=85)
    except Exception:
        img = Image.new("RGB", (IMG_W, IMG_H), color=(13, 27, 42))
        img.save(cache_path, "JPEG")
    return cache_path


def get_image_for_step(etape: dict) -> Path:
    """Retourne le chemin de l'image (téléchargée ou placeholder)."""
    path = download_image(
        etape["wikimedia_img"],
        f"step_{etape['id']:02d}_{etape['nom'][:20].replace(' ', '_')}"
    )
    if path is None:
        path = create_placeholder_image(etape["nom"])
    return path


# ---------------------------------------------------------------------------
# FEUILLE 1 : ITINÉRAIRE COMPLET
# ---------------------------------------------------------------------------

def build_sheet1(wb: Workbook):
    ws = wb.active
    ws.title = "🗺️ Itinéraire"

    # Dimensions colonnes
    # A: N°, B: Image, C: Étape, D: Région/Ambiance, E: Prix entrée,
    # F: Horaires, G: Logement, H: Resto, I: Route conseillée, J: Points de vue
    col_widths = {
        "A": 5, "B": 33, "C": 30, "D": 22, "E": 16,
        "F": 22, "G": 32, "H": 28, "I": 35, "J": 32,
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # --- TITRE ---
    ws.merge_cells("A1:J1")
    title_cell = ws["A1"]
    title_cell.value = "🚗  ROAD TRIP FRANCE PREMIUM  🇫🇷  — Itinéraire Complet"
    title_cell.font = make_font("Calibri", 18, bold=True, color=C_ACCENT_GOLD)
    title_cell.fill = make_fill(C_HEADER_BG)
    title_cell.alignment = make_align("center")
    ws.row_dimensions[1].height = 40

    # --- SOUS-TITRE ---
    ws.merge_cells("A2:J2")
    sub_cell = ws["A2"]
    sub_cell.value = (
        "Chambord → Chenonceau → Beauval → Futuroscope → Lagon Saint-Pierre → "
        "Puy du Fou → Dune du Pilat → La Roque-Gageac → Beynac → Rocamadour → "
        "Padirac → Volcans Auvergne → Gorges Sioule → Montluçon → Bourges → Orléans → France Miniature"
    )
    sub_cell.font = make_font("Calibri", 10, italic=True, color=C_ACCENT_CORAL)
    sub_cell.fill = make_fill(C_SECTION_BG)
    sub_cell.alignment = make_align("center")
    ws.row_dimensions[2].height = 30

    # --- EN-TÊTES ---
    headers = [
        "N°", "📸 Photo", "📍 Étape", "🌿 Ambiance / Région",
        "💶 Prix entrée", "⏰ Horaires",
        "🛏️ Logement suggéré", "🍽️ Resto suggéré",
        "🛣️ Route conseillée", "🔭 Points de vue",
    ]
    add_header_row(ws, 3, headers)
    ws.row_dimensions[3].height = 22

    # --- LIGNES D'ÉTAPES ---
    ROW_H = 95  # hauteur pour les images
    for i, etape in enumerate(ETAPES):
        row = 4 + i
        ws.row_dimensions[row].height = ROW_H
        odd = (i % 2 == 0)
        bg = C_ROW_ODD if odd else C_ROW_EVEN

        # Numéro
        c_num = ws.cell(row=row, column=1, value=etape["id"])
        c_num.font = make_font("Calibri", 14, bold=True, color=C_ACCENT_GOLD)
        c_num.fill = make_fill(C_SECTION_BG if not odd else C_HEADER_BG)
        c_num.alignment = make_align("center")
        c_num.border = make_border(C_ACCENT_GOLD, "medium")

        # Image (colonne B = col 2)
        img_path = get_image_for_step(etape)
        try:
            xl_img = XLImage(str(img_path))
            xl_img.width = IMG_W
            xl_img.height = IMG_H
            ws.add_image(xl_img, f"B{row}")
        except Exception as exc:
            ws.cell(row=row, column=2, value=f"📍 {etape['nom']}")

        # Étape (avec hyperlien Maps)
        c_nom = ws.cell(row=row, column=3,
                        value=f"{etape['emoji']} {etape['nom']}")
        c_nom.hyperlink = etape["maps_url"]
        c_nom.font = make_font("Calibri", 12, bold=True, color="1155CC")
        c_nom.fill = make_fill(bg)
        c_nom.alignment = make_align()
        c_nom.border = make_border()

        # Ambiance / Région
        ws.cell(row=row, column=4,
                value=f"{etape['ambiance']}\n{etape['region']}"
                ).fill = make_fill(bg)
        ws.cell(row=row, column=4).font = make_font(size=10, color=C_TEXT_MID)
        ws.cell(row=row, column=4).alignment = make_align()
        ws.cell(row=row, column=4).border = make_border()

        # Prix entrée
        ws.cell(row=row, column=5,
                value=f"🎟️ {etape['prix_entree']}"
                ).fill = make_fill(bg)
        ws.cell(row=row, column=5).font = make_font(size=10, bold=True, color=C_ACCENT_CORAL)
        ws.cell(row=row, column=5).alignment = make_align("center")
        ws.cell(row=row, column=5).border = make_border()

        # Horaires
        ws.cell(row=row, column=6, value=f"⏰ {etape['horaires']}").fill = make_fill(bg)
        ws.cell(row=row, column=6).font = make_font(size=10, color=C_TEXT_DARK)
        ws.cell(row=row, column=6).alignment = make_align()
        ws.cell(row=row, column=6).border = make_border()

        # Logement (avec hyperlien)
        c_log = ws.cell(row=row, column=7, value=f"🛏️ {etape['logement']}")
        c_log.hyperlink = etape["logement_url"]
        c_log.font = make_font(size=10, color="1155CC")
        c_log.fill = make_fill(bg)
        c_log.alignment = make_align()
        c_log.border = make_border()

        # Resto (avec hyperlien)
        c_resto = ws.cell(row=row, column=8, value=f"🍽️ {etape['resto']}")
        c_resto.hyperlink = etape["resto_url"]
        c_resto.font = make_font(size=10, color="1155CC")
        c_resto.fill = make_fill(bg)
        c_resto.alignment = make_align()
        c_resto.border = make_border()

        # Route conseillée
        ws.cell(row=row, column=9, value=f"🛣️ {etape['routes']}").fill = make_fill(bg)
        ws.cell(row=row, column=9).font = make_font(size=10, color=C_TEXT_DARK)
        ws.cell(row=row, column=9).alignment = make_align()
        ws.cell(row=row, column=9).border = make_border()

        # Points de vue
        ws.cell(row=row, column=10,
                value=f"🔭 {etape['points_vue']}\n📷 {etape['conseils_photo']}"
                ).fill = make_fill(bg)
        ws.cell(row=row, column=10).font = make_font(size=9, italic=True, color=C_TEXT_MID)
        ws.cell(row=row, column=10).alignment = make_align()
        ws.cell(row=row, column=10).border = make_border()

    # Ligne de total / pied
    last_row = 4 + len(ETAPES)
    ws.merge_cells(f"A{last_row}:J{last_row}")
    footer = ws.cell(
        row=last_row, column=1,
        value=(
            "🏁  Boucle complète : ~2 200 km  |  10 jours recommandés  "
            "|  Budget hébergement ~1 100 €  |  Budget entrées ~230 €"
        ),
    )
    footer.font = make_font("Calibri", 11, bold=True, color=C_ACCENT_GOLD)
    footer.fill = make_fill(C_HEADER_BG)
    footer.alignment = make_align("center")
    footer.border = make_border(C_ACCENT_GOLD, "medium")
    ws.row_dimensions[last_row].height = 24

    # Figer la ligne d'en-têtes
    ws.freeze_panes = "A4"


# ---------------------------------------------------------------------------
# FEUILLE 2 : PLANNING JOUR PAR JOUR
# ---------------------------------------------------------------------------

def build_sheet2(wb: Workbook):
    ws = wb.create_sheet("📅 Planning")

    col_widths = {
        "A": 12, "B": 8, "C": 38, "D": 38, "E": 38,
        "F": 38, "G": 30, "H": 22, "I": 12, "J": 12,
    }
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Titre
    ws.merge_cells("A1:J1")
    t = ws["A1"]
    t.value = "📅  PLANNING DÉTAILLÉ — Jour par Jour"
    t.font = make_font("Calibri", 16, bold=True, color=C_ACCENT_GOLD)
    t.fill = make_fill(C_HEADER_BG)
    t.alignment = make_align("center")
    ws.row_dimensions[1].height = 36

    headers = [
        "🗓️ Jour", "📌 J+",
        "🌅 Matin (départ / visite 1)", "🌞 Déjeuner",
        "☀️ Après-midi (visite 2)", "🌆 Dîner",
        "🛏️ Nuit / Hébergement", "📍 Étapes clés",
        "📏 km", "🕐 Route",
    ]
    add_header_row(ws, 2, headers)
    ws.row_dimensions[2].height = 22

    for i, jour in enumerate(PLANNING):
        row = 3 + i
        odd = (i % 2 == 0)
        bg = C_ROW_ODD if odd else C_ROW_EVEN
        ws.row_dimensions[row].height = 60

        values = [
            jour["jour"],
            jour["date"],
            f"{jour['matin']}\n{jour['visite1']}",
            jour["dejeuner"],
            f"{jour['apresmidi']}\n{jour.get('visite2', '')}",
            jour["diner"],
            jour["nuit"],
            "",  # étapes clés (vide, formaté séparément)
            jour["km"],
            jour["temps_route"],
        ]
        for col_idx, val in enumerate(values, start=1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            cell.fill = make_fill(bg)
            cell.alignment = make_align()
            cell.border = make_border()
            if col_idx == 1:
                cell.font = make_font("Calibri", 11, bold=True, color=C_ACCENT_GOLD)
                cell.fill = make_fill(C_SECTION_BG)
            elif col_idx in (9, 10):
                cell.font = make_font("Calibri", 10, bold=True, color=C_ACCENT_CORAL)
                cell.alignment = make_align("center")
            else:
                cell.font = make_font(size=10, color=C_TEXT_DARK)

    # Pied de page
    last = 3 + len(PLANNING)
    ws.merge_cells(f"A{last}:J{last}")
    f = ws.cell(row=last, column=1,
                value="💡 Conseil : adapter le planning selon météo et temps de visite souhaité. Prévoir +1 jour pour Futuroscope/Puy du Fou.")
    f.font = make_font(size=10, italic=True, color=C_ACCENT_GOLD)
    f.fill = make_fill(C_SECTION_BG)
    f.alignment = make_align("center")
    ws.row_dimensions[last].height = 24
    ws.freeze_panes = "A3"


# ---------------------------------------------------------------------------
# FEUILLE 3 : CARTE DU TRAJET
# ---------------------------------------------------------------------------

def build_sheet3(wb: Workbook):
    ws = wb.create_sheet("🗺️ Carte Trajet")

    col_widths = {"A": 30, "B": 30, "C": 12, "D": 12, "E": 25, "F": 30, "G": 30}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = "🗺️  CARTE DU TRAJET — Distances, Temps & Points de Vue"
    t.font = make_font("Calibri", 16, bold=True, color=C_ACCENT_GOLD)
    t.fill = make_fill(C_HEADER_BG)
    t.alignment = make_align("center")
    ws.row_dimensions[1].height = 36

    headers = ["🅰️ Départ", "🅱️ Arrivée", "📏 Distance", "⏱️ Temps", "🛣️ Route", "🌄 Paysage", "🔭 Point de vue"]
    add_header_row(ws, 2, headers)
    ws.row_dimensions[2].height = 22

    total_km = 0
    total_min = 0
    for i, seg in enumerate(CARTE):
        row = 3 + i
        odd = (i % 2 == 0)
        bg = C_ROW_ODD if odd else C_ROW_EVEN
        ws.row_dimensions[row].height = 30

        dep, arr, dist, temps, route, paysage, vue = seg

        # Parse km
        try:
            km_val = int(dist.replace(" km", "").replace("~", "").strip())
            total_km += km_val
        except Exception:
            pass
        # Parse minutes
        try:
            parts = temps.replace("h", ":").split(":")
            total_min += int(parts[0]) * 60 + int(parts[1])
        except Exception:
            pass

        data = [dep, arr, dist, temps, route, paysage, vue]
        for col_idx, val in enumerate(data, start=1):
            cell = ws.cell(row=row, column=col_idx, value=val)
            cell.fill = make_fill(bg)
            cell.alignment = make_align()
            cell.border = make_border()
            cell.font = make_font(size=10, color=C_TEXT_DARK)
            if col_idx in (3, 4):
                cell.font = make_font(size=11, bold=True, color=C_ACCENT_CORAL)
                cell.alignment = make_align("center")

    # Total
    total_row = 3 + len(CARTE)
    ws.merge_cells(f"A{total_row}:B{total_row}")
    ws.cell(row=total_row, column=1, value="🏁  TOTAL BOUCLE").font = make_font(
        "Calibri", 12, bold=True, color=C_ACCENT_GOLD)
    ws.cell(row=total_row, column=1).fill = make_fill(C_SECTION_BG)
    ws.cell(row=total_row, column=1).alignment = make_align("center")
    ws.cell(row=total_row, column=1).border = make_border(C_ACCENT_GOLD, "medium")

    ws.cell(row=total_row, column=3,
            value=f"~{total_km} km").font = make_font("Calibri", 12, bold=True, color=C_ACCENT_GOLD)
    ws.cell(row=total_row, column=3).fill = make_fill(C_SECTION_BG)
    ws.cell(row=total_row, column=3).alignment = make_align("center")
    ws.cell(row=total_row, column=3).border = make_border(C_ACCENT_GOLD, "medium")

    total_h = total_min // 60
    total_m = total_min % 60
    ws.cell(row=total_row, column=4,
            value=f"~{total_h}h{total_m:02d}").font = make_font("Calibri", 12, bold=True, color=C_ACCENT_GOLD)
    ws.cell(row=total_row, column=4).fill = make_fill(C_SECTION_BG)
    ws.cell(row=total_row, column=4).alignment = make_align("center")
    ws.cell(row=total_row, column=4).border = make_border(C_ACCENT_GOLD, "medium")

    for col in (5, 6, 7):
        ws.cell(row=total_row, column=col).fill = make_fill(C_SECTION_BG)
        ws.cell(row=total_row, column=col).border = make_border(C_ACCENT_GOLD, "medium")

    ws.row_dimensions[total_row].height = 28
    ws.freeze_panes = "A3"


# ---------------------------------------------------------------------------
# FEUILLE 4 : CHECKLIST VOYAGE
# ---------------------------------------------------------------------------

def build_sheet4(wb: Workbook):
    ws = wb.create_sheet("✅ Checklist")

    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 4
    ws.column_dimensions["D"].width = 38
    ws.column_dimensions["E"].width = 4
    ws.column_dimensions["F"].width = 38

    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = "✅  CHECKLIST VOYAGE — Ne rien oublier !"
    t.font = make_font("Calibri", 16, bold=True, color=C_ACCENT_GOLD)
    t.fill = make_fill(C_HEADER_BG)
    t.alignment = make_align("center")
    ws.row_dimensions[1].height = 36

    categories = list(CHECKLIST.keys())
    # Disposition en 3 colonnes : cat 0/1 col A-B, cat 2/3 col C-D, cat 4/5 col E-F
    col_map = [(1, 2), (3, 4), (5, 6)]

    # Calculer le nombre max de lignes nécessaires par colonne paire
    max_rows_per_col = [0, 0, 0]
    for ci, cat in enumerate(categories):
        col_pair = ci // 2
        rows_needed = 2 + len(CHECKLIST[cat])  # titre + items
        if ci % 2 == 0:
            max_rows_per_col[col_pair] = rows_needed
        else:
            max_rows_per_col[col_pair] = max(max_rows_per_col[col_pair], rows_needed)

    current_rows = [2, 2, 2]  # starting row for each column pair

    for ci, cat in enumerate(categories):
        col_pair = ci // 2
        check_col, label_col = col_map[col_pair]
        start_row = current_rows[col_pair]

        # En-tête de catégorie
        ws.merge_cells(
            start_row=start_row, start_column=check_col,
            end_row=start_row, end_column=label_col
        )
        cat_cell = ws.cell(row=start_row, column=check_col, value=cat)
        cat_cell.font = make_font("Calibri", 11, bold=True, color=C_HEADER_FG)
        cat_cell.fill = make_fill(C_HEADER_BG)
        cat_cell.alignment = make_align("center")
        cat_cell.border = make_border(C_ACCENT_GOLD, "medium")
        ws.row_dimensions[start_row].height = 22

        # Items
        for j, item in enumerate(CHECKLIST[cat]):
            r = start_row + 1 + j
            odd = (j % 2 == 0)
            bg = C_ROW_ODD if odd else C_ROW_EVEN
            ws.row_dimensions[r].height = 20

            # Case à cocher (□)
            chk = ws.cell(row=r, column=check_col, value="□")
            chk.font = make_font("Calibri", 14, color=C_ACCENT_CORAL)
            chk.fill = make_fill(bg)
            chk.alignment = make_align("center")
            chk.border = make_border()

            # Label
            lbl = ws.cell(row=r, column=label_col, value=item)
            lbl.font = make_font(size=10, color=C_TEXT_DARK)
            lbl.fill = make_fill(bg)
            lbl.alignment = make_align()
            lbl.border = make_border()

        # Mettre à jour la ligne courante pour cette paire
        items_count = len(CHECKLIST[cat])
        # Pour les catégories impaires (seconde de la paire), on avance la ligne
        if ci % 2 == 1:
            current_rows[col_pair] += max(
                len(CHECKLIST[categories[ci - 1]]),
                items_count
            ) + 2
        # Pour la première catégorie d'une paire, on commence par la même ligne
        # et la seconde catégorie sera dans l'autre colonne de la même paire

    # Note de bas
    max_end = max(current_rows)
    ws.merge_cells(f"A{max_end}:F{max_end}")
    note = ws.cell(
        row=max_end, column=1,
        value="💡 Imprimez cette checklist et cochez chaque item avant le départ !  Bon voyage ! 🚗✨"
    )
    note.font = make_font(size=10, italic=True, color=C_ACCENT_GOLD)
    note.fill = make_fill(C_SECTION_BG)
    note.alignment = make_align("center")
    note.border = make_border(C_ACCENT_GOLD, "medium")


# ---------------------------------------------------------------------------
# FEUILLE 5 : PANORAMA — Mosaïque d'images
# ---------------------------------------------------------------------------

def build_sheet5(wb: Workbook, images: list):
    ws = wb.create_sheet("🖼️ Panorama")

    COLS_PER_ROW = 4
    CELL_W = 35   # largeur colonne (caractères)
    CELL_H = 90   # hauteur ligne (points)
    TITLE_H = 20  # hauteur extra pour titre sous image

    # Titre principal
    total_cols = COLS_PER_ROW * 2  # image + titre ?
    ws.merge_cells(start_row=1, start_column=1,
                   end_row=1, end_column=COLS_PER_ROW)
    t = ws["A1"]
    t.value = "🖼️  PANORAMA DU VOYAGE — Les Plus Belles Images"
    t.font = make_font("Calibri", 16, bold=True, color=C_ACCENT_GOLD)
    t.fill = make_fill(C_HEADER_BG)
    t.alignment = make_align("center")
    ws.row_dimensions[1].height = 36

    # Configurer les colonnes
    for col in range(1, COLS_PER_ROW + 1):
        ws.column_dimensions[get_column_letter(col)].width = CELL_W

    current_row = 2
    for i, (etape, img_path) in enumerate(zip(ETAPES, images)):
        col = (i % COLS_PER_ROW) + 1  # 1..4
        if i % COLS_PER_ROW == 0:
            # Nouvelle rangée — définir hauteur image + titre
            img_row = current_row
            title_row = current_row + 1
            ws.row_dimensions[img_row].height = CELL_H
            ws.row_dimensions[title_row].height = TITLE_H
            current_row += 2

        img_row = current_row - 2
        title_row = current_row - 1

        # Image
        try:
            xl_img = XLImage(str(img_path))
            xl_img.width = IMG_W
            xl_img.height = IMG_H
            cell_addr = f"{get_column_letter(col)}{img_row}"
            ws.add_image(xl_img, cell_addr)
        except Exception:
            ws.cell(row=img_row, column=col, value=f"{etape['emoji']} {etape['nom']}")

        # Titre sous image
        tc = ws.cell(row=title_row, column=col,
                     value=f"{etape['emoji']} {etape['nom']}")
        tc.font = make_font("Calibri", 10, bold=True, color=C_ACCENT_GOLD)
        tc.fill = make_fill(C_SECTION_BG)
        tc.alignment = make_align("center")
        tc.border = make_border(C_ACCENT_GOLD)

        # Fond de la cellule image
        ic = ws.cell(row=img_row, column=col)
        ic.fill = make_fill(C_HEADER_BG)
        ic.border = make_border(C_ACCENT_GOLD, "medium")

    # Pied de page
    last_row = current_row
    ws.merge_cells(start_row=last_row, start_column=1,
                   end_row=last_row, end_column=COLS_PER_ROW)
    f = ws.cell(row=last_row, column=1,
                value="📸 Photos sources : Wikimedia Commons (CC) — Libre de droits")
    f.font = make_font(size=9, italic=True, color=C_TEXT_MID)
    f.fill = make_fill(C_ROW_ODD)
    f.alignment = make_align("center")


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

def main():
    output_file = "Road_Trip_Vendeur.xlsx"
    print("🚗  Génération du fichier Excel Road Trip Premium…\n")

    # 1. Télécharger les images
    print("📥  Téléchargement des images…")
    images = []
    for etape in ETAPES:
        img_path = get_image_for_step(etape)
        images.append(img_path)
        time.sleep(0.3)  # politesse API Wikimedia

    # 2. Créer le workbook
    wb = Workbook()

    print("\n📊  Construction des feuilles Excel…")

    print("  → Feuille 1 : Itinéraire complet")
    build_sheet1(wb)

    print("  → Feuille 2 : Planning jour par jour")
    build_sheet2(wb)

    print("  → Feuille 3 : Carte du trajet")
    build_sheet3(wb)

    print("  → Feuille 4 : Checklist voyage")
    build_sheet4(wb)

    print("  → Feuille 5 : Panorama")
    build_sheet5(wb, images)

    # 3. Propriétés du workbook
    wb.properties.title = "Road Trip France Premium"
    wb.properties.subject = "Itinéraire road trip France 17 étapes"
    wb.properties.description = (
        "Fichier Excel premium généré automatiquement — "
        "Chambord · Chenonceau · Beauval · Futuroscope · Puy du Fou · "
        "Dune du Pilat · Rocamadour · Volcans d'Auvergne · Bourges · Orléans"
    )

    # 4. Sauvegarder
    wb.save(output_file)
    print(f"\n✅  Fichier généré : {output_file}")
    print("🎉  Bon voyage !")


if __name__ == "__main__":
    main()
