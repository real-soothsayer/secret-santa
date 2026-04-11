# 🚗 Road Trip France Premium — Générateur Excel

> Génère un fichier **Excel haut de gamme** pour un road trip en France sur 17 étapes iconiques, avec images, hyperliens, prix, horaires, restos et design brochure de voyage.

---

## ✨ Aperçu du fichier généré

Le script produit `Road_Trip_Vendeur.xlsx`, un fichier Excel **5 feuilles** entièrement illustré et mis en forme :

| Feuille | Contenu |
|---------|---------|
| 🗺️ **Itinéraire** | 17 étapes avec photo, prix, horaires, logement, resto, route, ambiance |
| 📅 **Planning** | Jour par jour détaillé : matin, déjeuner, après-midi, dîner, nuit |
| 🗺️ **Carte Trajet** | Distances, temps de route, routes conseillées, points de vue |
| ✅ **Checklist** | Liste complète de préparation en 6 catégories |
| 🖼️ **Panorama** | Mosaïque 4 colonnes des plus belles photos de chaque étape |

---

## 🗺️ Itinéraire (ordre exact)

```
1  🏰  Château de Chambord
2  🌹  Château de Chenonceau
3  🐼  ZooParc de Beauval
4  🚀  Futuroscope
5  🏊  Lagon du Bois de Saint-Pierre
6  ⚔️  Puy du Fou
7  🏖️  Dune du Pilat
8  🛶  La Roque-Gageac
9  🏯  Beynac-et-Cazenac
10 ⛪  Rocamadour
11 🕳️  Gouffre de Padirac
12 🌋  Parc naturel régional des Volcans d'Auvergne
13 🏞️  Gorges de la Sioule
14 🎶  Montluçon
15 🕍  Bourges — Cathédrale Saint-Étienne
16 ⚜️  Orléans
17 🗺️  France Miniature
```

**Distance totale :** ~2 200 km  
**Durée recommandée :** 10 jours  

---

## 🚀 Installation & Utilisation

### Prérequis
- Python 3.9+
- pip

### Installation

```bash
git clone https://github.com/real-soothsayer/secret-santa.git
cd secret-santa
pip install -r requirements.txt
```

### Génération du fichier

```bash
python generate_roadtrip.py
```

Le fichier `Road_Trip_Vendeur.xlsx` est créé dans le répertoire courant.

> **Note :** Le script télécharge automatiquement les images depuis Wikimedia Commons (libre de droits). Les images sont mises en cache dans `/tmp/roadtrip_img_cache/` pour éviter de les re-télécharger. Si une image n'est pas disponible, un placeholder coloré est généré automatiquement.

---

## 🎨 Design Premium

Le fichier Excel utilise un thème **voyage haut de gamme** :

- 🌙 **Palette** : Bleu nuit (`#0D1B2A`), or (`#C9A84C`), corail (`#E07B54`), beige
- 🔤 **Typographie** : Calibri, tailles hiérarchisées (18pt titres → 9pt légendes)
- 📸 **Images** : 220×120 px, auto-téléchargées depuis Wikimedia Commons
- 🔗 **Hyperliens** : Google Maps, Booking.com, TripAdvisor directement cliquables
- 🎨 **Lignes alternées** : bleu clair / blanc pour faciliter la lecture
- 🏷️ **En-têtes foncés** avec texte or/beige, bordures accent or

---

## 🛠️ Personnalisation

### Modifier les prix / hébergements

Ouvrez `generate_roadtrip.py` et éditez le dictionnaire `ETAPES` :

```python
{
    "nom": "Château de Chambord",
    "prix_entree": "14,50 €",          # ← modifier le prix
    "logement": "Mon hôtel — ~90 €",   # ← changer l'hébergement
    "logement_url": "https://...",     # ← URL cliquable
    "resto": "Mon resto favori",       # ← votre resto
    ...
}
```

### Changer les images

Remplacez la clé `wikimedia_img` par l'URL de votre image (JPG/PNG) :

```python
"wikimedia_img": "https://monsite.com/ma-photo.jpg",
```

### Adapter le planning

Modifiez la liste `PLANNING` pour ajuster les horaires, durées et hébergements.

### Ajouter/supprimer des étapes

Ajoutez un dictionnaire dans `ETAPES` en suivant la structure existante. La numérotation et la mise en page s'adaptent automatiquement.

---

## 📦 Dépendances

| Package | Version | Usage |
|---------|---------|-------|
| `openpyxl` | ≥ 3.1.2 | Création et mise en forme Excel |
| `requests` | ≥ 2.31.0 | Téléchargement des images |
| `Pillow` | ≥ 10.0.0 | Redimensionnement des images |

---

## 📋 Structure des données

### Chaque étape contient

| Champ | Description |
|-------|-------------|
| `nom` | Nom de l'étape |
| `emoji` | Icône représentative |
| `region` | Région administrative |
| `ambiance` | Description ambiance/paysage |
| `prix_entree` | Prix adulte (entrée) |
| `horaires` | Horaires d'ouverture |
| `logement` | Hébergement suggéré + prix indicatif |
| `logement_url` | Lien Booking/AirBnB (cliquable) |
| `maps_url` | Lien Google Maps (cliquable) |
| `resto` | Restaurant suggéré + prix |
| `resto_url` | Lien Google Maps resto |
| `routes` | Routes conseillées / panoramiques |
| `conseils_photo` | Meilleurs spots photo |
| `points_vue` | Points de vue à ne pas manquer |
| `wikimedia_img` | URL image Wikimedia Commons |

---

## 💡 Astuces

- **Hors-saison** : les prix d'entrée affichés sont indicatifs — vérifiez les sites officiels avant de partir
- **Réservations** : Futuroscope, Puy du Fou et ZooParc de Beauval se réservent à l'avance en haute saison
- **Route des Châteaux** : combiner Chambord et Chenonceau en une journée est possible mais intense
- **Gorges de la Sioule** : moins connue mais magnifique — idéale pour la randonnée
- **Dune du Pilat** : allez-y tôt le matin pour éviter les foules en juillet/août

---

## 📄 Licence

Images : **Wikimedia Commons** (Creative Commons)  
Code : **MIT**

---

*Fait avec ❤️ pour les amoureux de la France 🇫🇷*
