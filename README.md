# L'Embryogenèse Somatique chez les Plantes

Présentation académique PowerPoint (7 diapositives, 16:9) — Module **Biotechnologie Appliquée**, 3ème Année Ingénieur, Sciences Agronomiques.

## Contenu du dossier

| Fichier | Description |
|---|---|
| `Embryogenese_Somatique_Presentation.pptx` | Présentation PowerPoint éditable (7 diapositives, transitions et animations natives, notes de l'orateur en français) |
| `Embryogenese_Somatique_Presentation.pdf` | Aperçu PDF (rendu statique) |
| `preview_slide1..7.png` | Aperçus PNG des diapositives |
| `assets/` | Photographies scientifiques authentiques (Wikimedia Commons) |
| `CREDITS.md` | Crédits photographiques et licences |
| `build_deck.py`, `deckkit.py` | Script de génération (python-pptx + XML natif de transitions/animations) |

## Animations natives intégrées

- **Transitions** : Fondu (1), Morphose (2, repli : Fondu), Poussée (3, 5), Balayage (4), Fondu au noir (7).
- **Animations d'objets** : Fondu, Apparition, Balayage (flèches/processus), Zoom (termes clés, photos), enchaînées **Au clic** / **Après la précédente** / **Avec la précédente**, durées 0,3–0,8 s.

## Notes de l'orateur

Les explications orales complètes en français sont intégrées dans le volet Notes de chaque diapositive.

## Régénération

```bash
pip install python-pptx pillow
python3 build_deck.py
```
