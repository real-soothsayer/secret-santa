# 🇫🇷 Road Trip France — 5 Jours Inoubliables

Génère un fichier Excel **visuellement attrayant et vendeur** pour un road trip de 5 jours en France, avec images des attractions, design premium et liens Google Maps cliquables.

## Itinéraire (9 étapes)

| Jour | Lieu | Région |
|------|------|--------|
| Jour 1 | Le Mont-Saint-Michel | Normandie |
| Jour 2 | Puy du Fou | Pays de la Loire |
| Jour 2 | Dune du Pilat | Nouvelle-Aquitaine |
| Jour 3 | Rocamadour | Occitanie |
| Jour 3 | Gouffre de Padirac | Occitanie |
| Jour 4 | Gorges du Verdon | Provence-Alpes-Côte d'Azur |
| Jour 4 | Calanques de Cassis | Provence-Alpes-Côte d'Azur |
| Jour 5 | Cirque du Fer à Cheval | Alpes |
| Jour 5 | QC Terme Chamonix | Alpes |

## Installation

```bash
pip install -r requirements.txt
```

## Utilisation

```bash
python generate_roadtrip.py
```

Le fichier **`Road_Trip_France_5_Jours.xlsx`** est généré dans le répertoire courant.

## Contenu du fichier Excel (4 onglets)

| Onglet | Description |
|--------|-------------|
| 🗺️ Itinéraire | Tableau principal avec photos, liens Maps et conseils |
| 📅 Planning détaillé | Programme heure par heure avec hébergements suggérés |
| 🚗 Carte du trajet | Toutes les étapes avec distances et durées de route |
| ✅ Checklist voyage | Liste de vérification complète avant le départ |

## Dépendances

- `openpyxl>=3.1.0` — création et mise en forme Excel
- `requests>=2.28.0` — téléchargement des images Wikimedia
- `Pillow>=9.0.0` — redimensionnement des images

## Notes

- Les images sont téléchargées depuis Wikimedia Commons et redimensionnées en 240×160 px.
- Si une image ne peut pas être téléchargée, un texte placeholder est inséré dans la cellule.
- Les liens Google Maps sont des hyperliens cliquables dans Excel.
- Total du trajet estimé : ~2 800 km / ~30h de route (boucle au départ de Paris).
