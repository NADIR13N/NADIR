# -*- coding: utf-8 -*-
"""Builds the 7-slide premium deck with native PPTX transitions + animations."""

import os
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from deckkit import (In, C, rect, oval, shape, shape_text, tb, bullets, pic,
                     header, footer, notes, add_animations, add_transition,
                     GREEN_D, GREEN_E, SAGE, IVORY, WHITE, TEAL, GOLD,
                     CHARCOAL, TERRA, BROWN, GREY_TX, F_TITLE, F_BODY)

A = os.environ.get("ASSET_DIR", "assets")
os.chdir(os.path.dirname(os.path.abspath(__file__)))

prs = Presentation()
prs.slide_width = In(13.333)
prs.slide_height = In(7.5)
BLANK = prs.slide_layouts[6]

prs.core_properties.title = "L'embryogenèse somatique chez les plantes"
prs.core_properties.subject = "Biotechnologie Appliquée — Sciences Agronomiques"
prs.core_properties.author = "3ème Année Ingénieur"
prs.core_properties.language = "fr-FR"


def new_slide():
    return prs.slides.add_slide(BLANK)


def S(shp, effect, trigger, delay=0, dur=500):
    return {'spid': shp.shape_id, 'effect': effect, 'trigger': trigger,
            'delay': delay, 'dur': dur}


# ================================================================ SLIDE 1
s1 = new_slide()
bg = rect(s1, 0, 0, 13.334, 7.5, fill=GREEN_D)
photo = pic(s1, f"{A}/cover_lab.jpg", 7.75, 0, 5.584, 7.5)
ov = rect(s1, 7.75, 0, 5.584, 7.5, fill=GREEN_D, alpha=32)
divider = rect(s1, 7.63, 0, 0.075, 7.5, fill=GOLD)

badge = tb(s1, 0.85, 1.00, 6.4, 0.4,
           [{'runs': [{'t': "BIOTECHNOLOGIE APPLIQUÉE  ·  SCIENCES AGRONOMIQUES",
                       'sz': 12, 'c': GOLD, 'b': True, 'spc': 300}]}])
title = tb(s1, 0.85, 1.55, 6.55, 2.7,
           [{'runs': [{'t': "L'EMBRYOGENÈSE", 'sz': 40, 'c': WHITE, 'b': True, 'f': F_TITLE}],
             'line': 1.04},
            {'runs': [{'t': "SOMATIQUE", 'sz': 40, 'c': WHITE, 'b': True, 'f': F_TITLE}],
             'line': 1.04},
            {'runs': [{'t': "CHEZ LES PLANTES", 'sz': 40, 'c': SAGE, 'b': True, 'f': F_TITLE}],
             'line': 1.04}])
rule = rect(s1, 0.88, 4.42, 1.6, 0.05, fill=GOLD)
subtitle = tb(s1, 0.85, 4.68, 6.4, 0.9,
              [{'runs': [{'t': "Principes, Étapes et Applications en Ingénierie Agronomique",
                          'sz': 17, 'c': IVORY, 'i': True}], 'line': 1.15}])
info = tb(s1, 0.85, 5.62, 6.4, 1.7,
          [{'runs': [{'t': "Présenté par : ", 'sz': 12.5, 'c': GOLD, 'b': True},
                     {'t': "[Ton Nom et Prénom]", 'sz': 12.5, 'c': WHITE}], 'space_after': 5},
           {'runs': [{'t': "Module : ", 'sz': 12.5, 'c': GOLD, 'b': True},
                     {'t': "Biotechnologie Appliquée", 'sz': 12.5, 'c': IVORY}], 'space_after': 5},
           {'runs': [{'t': "Filière : ", 'sz': 12.5, 'c': GOLD, 'b': True},
                     {'t': "Sciences Agronomiques", 'sz': 12.5, 'c': IVORY}], 'space_after': 5},
           {'runs': [{'t': "3ème Année Ingénieur", 'sz': 12.5, 'c': IVORY}]}])

inset = pic(s1, f"{A}/embryo_single.jpg", 8.55, 4.55, 2.95, 2.05, line=WHITE, line_w=2.25)
inset_cap = tb(s1, 8.55, 6.68, 2.95, 0.35,
               [{'runs': [{'t': "Embryon somatique observé in vitro", 'sz': 9, 'c': IVORY, 'i': True}]}])

notes(s1, """Bonjour Monsieur le professeur, bonjour à tous mes chers camarades.

Aujourd'hui, nous allons explorer l'une des techniques les plus fascinantes de la biotechnologie végétale : l'embryogenèse somatique.

Nous verrons comment cette méthode permet, à partir d'une simple cellule, de recréer une plante entière, et comment elle révolutionne aujourd'hui l'agriculture et la foresterie.""")

add_transition(s1, 'fade')
add_animations(s1, [
    S(photo, 'fade', 'auto', 0, 700), S(ov, 'fade', 'with', 0, 700),
    S(divider, 'wipe_b', 'after', 0, 500),
    S(badge, 'fade', 'after', 200, 500),
    S(title, 'fade', 'after', 300, 600),
    S(rule, 'wipe_l', 'after', 200, 500),
    S(subtitle, 'fade', 'after', 200, 500),
    S(info, 'fade', 'after', 250, 500),
    S(inset, 'zoom', 'after', 350, 600), S(inset_cap, 'fade', 'with', 0, 500),
])

# ================================================================ SLIDE 2
s2 = new_slide()
rect(s2, 0, 0, 13.334, 7.5, fill=IVORY)
header(s2, "Définition et principe fondamental", "De la cellule somatique à la plante entière")

def_card = rect(s2, 0.55, 1.42, 5.55, 1.42, fill=WHITE, line=SAGE, line_w=0.75, round_=0.07)
def_acc = rect(s2, 0.55, 1.42, 0.075, 1.42, fill=GOLD)
def_lbl = tb(s2, 0.85, 1.58, 5.0, 0.3,
             [{'runs': [{'t': "DÉFINITION", 'sz': 10.5, 'c': GOLD, 'b': True, 'spc': 250}]}])
def_txt = tb(s2, 0.85, 1.92, 5.05, 0.85,
             [{'runs': [{'t': "Formation d'un embryon à partir de ", 'sz': 13.5, 'c': CHARCOAL},
                        {'t': "cellules végétatives (somatiques)", 'sz': 13.5, 'c': GREEN_D, 'b': True},
                        {'t': ", sans fécondation.", 'sz': 13.5, 'c': CHARCOAL}], 'line': 1.12}])

tot_card = rect(s2, 0.55, 3.05, 5.55, 2.42, fill=WHITE, line=SAGE, line_w=0.75, round_=0.06)
tot_acc = rect(s2, 0.55, 3.05, 0.075, 2.42, fill=TEAL)
tot_lbl = tb(s2, 0.85, 3.21, 5.1, 0.55,
             [{'runs': [{'t': "LE PRINCIPE CLÉ : LA ", 'sz': 10.5, 'c': TEAL, 'b': True, 'spc': 200},
                        {'t': "TOTIPOTENCE", 'sz': 10.5, 'c': GREEN_D, 'b': True, 'spc': 200}]}])
tot_bul = bullets(s2, 0.85, 3.62, 5.05, 1.75, [
    [{'t': "Capacité d'une cellule différenciée à ", 'sz': 11.5},
     {'t': "redevenir indifférenciée", 'sz': 11.5, 'b': True, 'c': GREEN_D}, {'t': ".", 'sz': 11.5}],
    [{'t': "Aptitude à ", 'sz': 11.5},
     {'t': "régénérer un organisme complet", 'sz': 11.5, 'b': True, 'c': GREEN_D}, {'t': ".", 'sz': 11.5}],
    [{'t': "Différence avec l'embryon zygotique : ", 'sz': 11.5},
     {'t': "absence de fécondation", 'sz': 11.5, 'b': True, 'c': GREEN_D},
     {'t': " et de recombinaison génétique liée à la reproduction sexuée.", 'sz': 11.5}],
], gap=6, dot=TEAL)

chip = rect(s2, 0.55, 5.72, 4.35, 0.62, fill=GOLD, round_=0.5)
shape_text(chip, "Concept clé : Totipotence cellulaire", 13, WHITE, bold=True)

ph_a = pic(s2, f"{A}/callus_ms.jpg", 6.45, 1.42, 3.20, 1.72, line=WHITE, line_w=2)
ph_b = pic(s2, f"{A}/embryo_heart.jpg", 9.85, 1.42, 3.0, 1.72, line=WHITE, line_w=2)
cap_a = tb(s2, 6.45, 3.16, 3.2, 0.3,
           [{'runs': [{'t': "Cal embryogène sur milieu gélosé", 'sz': 9, 'c': GREY_TX, 'i': True}]}])
cap_b = tb(s2, 9.85, 3.16, 3.0, 0.3,
           [{'runs': [{'t': "Embryons somatiques (stade précoce)", 'sz': 9, 'c': GREY_TX, 'i': True}]}])

chain = []
labels = [("Cellule somatique\n(végétative)", WHITE, GREEN_D),
          ("Reprogrammation\ncellulaire", WHITE, GREEN_D),
          ("Embryon\nsomatique", WHITE, GREEN_D),
          ("Plante entière\nrégénérée", GREEN_E, WHITE)]
cx, cw, chh = 6.45, 1.42, 1.05
cy = 3.62
for i, (lab, fillc, txtc) in enumerate(labels):
    x = cx + i * (cw + 0.24)
    nd = rect(s2, x, cy, cw, chh, fill=fillc,
              line=(GREEN_E if fillc == WHITE else None), line_w=1.25, round_=0.12)
    tfl = nd.text_frame
    tfl.word_wrap = True
    tfl.margin_left = In(0.04); tfl.margin_right = In(0.04)
    tfl.vertical_anchor = MSO_ANCHOR.MIDDLE
    pl = tfl.paragraphs[0]; pl.alignment = PP_ALIGN.CENTER
    for j, part in enumerate(lab.split("\n")):
        if j > 0:
            pl = tfl.add_paragraph(); pl.alignment = PP_ALIGN.CENTER
        r = pl.add_run(); r.text = part
        r.font.size = Pt(10.5); r.font.bold = True
        r.font.name = F_BODY; r.font.color.rgb = C(txtc)
    chain.append(nd)
    if i < 3:
        ar = shape(s2, MSO_SHAPE.CHEVRON, x + cw + 0.015, cy + chh / 2 - 0.13, 0.21, 0.26, fill=GOLD)
        chain.append(ar)

cmp1 = rect(s2, 6.45, 5.05, 3.05, 1.28, fill=WHITE, line=TEAL, line_w=1.0, round_=0.09)
tb(s2, 6.62, 5.18, 2.75, 1.05,
   [{'runs': [{'t': "Embryogenèse zygotique", 'sz': 10.5, 'c': TEAL, 'b': True}], 'space_after': 3},
    {'runs': [{'t': "Ovule fécondé · recombinaison génétique parentale", 'sz': 9.5, 'c': CHARCOAL}], 'line': 1.1}])
cmp2 = rect(s2, 9.80, 5.05, 3.05, 1.28, fill=TEAL, round_=0.09)
tb(s2, 9.97, 5.18, 2.75, 1.05,
   [{'runs': [{'t': "Embryogenèse somatique", 'sz': 10.5, 'c': WHITE, 'b': True}], 'space_after': 3},
    {'runs': [{'t': "Sans fécondation · multiplication clonale possible", 'sz': 9.5, 'c': IVORY}], 'line': 1.1}])

footer(s2, 2)
notes(s2, """Qu'est-ce que l'embryogenèse somatique ? C'est tout simplement la capacité de créer un embryon sans passer par la reproduction sexuée. On part d'une cellule ordinaire, comme une cellule de feuille, pour en faire un arbre complet.

Tout cela repose sur un principe fondamental associé aux travaux de Haberlandt en 1902 : la totipotence cellulaire. C'est la capacité d'une cellule végétale à reprogrammer son activité génétique, puis à se différencier pour former une plante entière dans des conditions appropriées.

Contrairement à la reproduction sexuée classique, cette méthode ne repose pas sur la fécondation ni sur le brassage génétique parental. Elle permet donc une multiplication clonale, même si des variations somaclonales peuvent parfois apparaître.""")

add_transition(s2, 'morph')
add_animations(s2, [
    S(def_card, 'fade', 'click', 0, 400), S(def_acc, 'wipe_l', 'with', 0, 400),
    S(def_lbl, 'fade', 'with', 0, 400), S(def_txt, 'fade', 'after', 150, 500),
    S(tot_card, 'fade', 'click', 0, 400), S(tot_acc, 'wipe_l', 'with', 0, 400),
    S(tot_lbl, 'fade', 'with', 0, 400),
    S(tot_bul, 'fade', 'after', 150, 500),
    S(ph_a, 'fade', 'after', 150, 500), S(cap_a, 'fade', 'with', 0, 500),
    S(ph_b, 'fade', 'after', 150, 500), S(cap_b, 'fade', 'with', 0, 500),
    S(chain[0], 'wipe_l', 'click', 0, 400),
    S(chain[1], 'wipe_l', 'after', 100, 300), S(chain[2], 'fade', 'after', 100, 400),
    S(chain[3], 'wipe_l', 'after', 100, 300), S(chain[4], 'fade', 'after', 100, 400),
    S(chain[5], 'wipe_l', 'after', 100, 300), S(chain[6], 'fade', 'after', 100, 400),
    S(cmp1, 'fade', 'after', 200, 400), S(cmp2, 'fade', 'after', 150, 400),
    S(chip, 'zoom', 'after', 250, 500),
])

# ================================================================ SLIDE 3
s3 = new_slide()
rect(s3, 0, 0, 13.334, 7.5, fill=IVORY)
header(s3, "Les étapes du processus", "Un cycle rigoureux en cinq étapes, de l'explant à la plante acclimatée")

stages = [
    ("1", "Induction", f"{A}/stage_induction.jpg",
     ["Formation de cals embryogènes.", "Milieu riche en auxines (2,4-D selon le protocole)."]),
    ("2", "Maintenance", f"{A}/stage_liquid.jpg",
     ["Multiplication des structures embryogènes.", "Possibilité de culture en milieu liquide."]),
    ("3", "Maturation", f"{A}/embryo_cluster.jpg",
     ["Développement jusqu'au stade cotylédonaire.", "Réduction ou retrait des auxines selon le protocole."]),
    ("4", "Germination", f"{A}/stage_rooted.jpg",
     ["Développement de la plantule.", "Transition vers l'autotrophie."]),
    ("5", "Acclimatation", f"{A}/stage_acclim.jpg",
     ["Transfert progressif en serre.", "Adaptation aux conditions ex vitro et au sol."]),
]
cw, gap, cy, chh = 2.25, 0.27, 1.62, 4.42
s3_groups = []
for i, (num, name, img, pts) in enumerate(stages):
    x = 0.5 + i * (cw + gap)
    card = rect(s3, x, cy, cw, chh, fill=WHITE, line=SAGE, line_w=0.75, round_=0.05)
    ph = pic(s3, img, x + 0.09, cy + 0.09, cw - 0.18, 1.18)
    circ = oval(s3, x + 0.09, cy + 1.02, 0.52, 0.52, fill=GREEN_D, line=WHITE, line_w=1.5)
    shape_text(circ, num, 16, WHITE)
    nm = tb(s3, x + 0.70, cy + 1.14, cw - 0.78, 0.4,
            [{'runs': [{'t': name, 'sz': 13.5, 'c': GREEN_D, 'b': True, 'f': F_TITLE}]}])
    bl = bullets(s3, x + 0.14, cy + 1.72, cw - 0.28, chh - 1.85, pts, sz=9.8, gap=5, line=1.08)
    s3_groups.append((card, ph, circ, nm, bl))
    if i < 4:
        shape(s3, MSO_SHAPE.CHEVRON, x + cw + 0.025, cy + 1.95, 0.22, 0.30, fill=GOLD)

bar = rect(s3, 0.5, 6.32, 12.33, 0.13, fill=SAGE, round_=0.5)
dots = []
for i in range(5):
    x = 0.5 + i * (cw + gap) + cw / 2 - 0.09
    dots.append(oval(s3, x, 6.285, 0.2, 0.2, fill=GREEN_D, line=WHITE, line_w=1))
prog = tb(s3, 0.5, 6.62, 12.33, 0.35,
          [{'runs': [{'t': "Explant → Call embryogène → Embryon → Plantule → Plante acclimatée",
                      'sz': 11, 'c': GREEN_D, 'b': True}], 'align': PP_ALIGN.CENTER}])

footer(s3, 3)
notes(s3, """Concrètement, comment se déroule ce processus au laboratoire ? Il suit un cycle rigoureux en cinq étapes.

D'abord, l'Induction : on place un explant sur un milieu riche en auxines, souvent le 2,4-D, pour créer un tissu appelé cal embryogène.

Ensuite, la Maintenance, où l'on multiplie ces cellules, souvent en bioréacteurs liquides.

Pour la Maturation, on modifie les conditions de culture et l'on réduit ou retire les auxines selon le protocole, afin de permettre à l'embryon d'atteindre le stade cotylédonaire, semblable à certaines étapes du développement d'une graine.

Enfin, la Germination et l'Acclimatation, où la plantule sort du milieu stérile pour s'habituer progressivement aux conditions naturelles en serre.""")

add_transition(s3, 'push_up')
specs3 = []
for i, (card, ph, circ, nm, bl) in enumerate(s3_groups):
    trig = 'click' if i == 0 else 'after'
    base = 0 if i == 0 else 150
    specs3 += [
        S(card, 'fade', trig, base, 400),
        S(ph, 'fade', 'with', 0, 500),
        S(circ, 'zoom', 'after', 100, 400),
        S(nm, 'fade', 'with', 0, 400),
        S(bl, 'fade', 'after', 100, 500),
    ]
specs3.append(S(bar, 'wipe_l', 'after', 200, 700))
for d in dots:
    specs3.append(S(d, 'zoom', 'with', 0, 400))
specs3.append(S(prog, 'fade', 'after', 200, 500))
add_animations(s3, specs3)

# ================================================================ SLIDE 4
s4 = new_slide()
rect(s4, 0, 0, 13.334, 7.5, fill=IVORY)
header(s4, "Facteurs de réussite et défis techniques", "Maîtriser le processus : facteurs clés, brunissement et conformité clonale")

cardA = rect(s4, 0.5, 1.42, 6.05, 3.15, fill=WHITE, line=SAGE, line_w=0.75, round_=0.05)
accA = rect(s4, 0.5, 1.42, 0.075, 3.15, fill=GREEN_E)
lblA = tb(s4, 0.78, 1.56, 5.6, 0.35,
          [{'runs': [{'t': "1 · L'ÉQUILIBRE HORMONAL", 'sz': 12, 'c': GREEN_D, 'b': True, 'spc': 150}]}])
beam = rect(s4, 1.35, 2.42, 3.30, 0.055, fill=CHARCOAL)
fulcrum = shape(s4, MSO_SHAPE.ISOSCELES_TRIANGLE, 2.83, 2.48, 0.34, 0.34, fill=GOLD)
base = rect(s4, 2.30, 2.86, 1.40, 0.055, fill=CHARCOAL)
wireL = rect(s4, 1.42, 2.10, 0.045, 0.34, fill=CHARCOAL)
wireR = rect(s4, 4.55, 2.10, 0.045, 0.34, fill=CHARCOAL)
panL = rect(s4, 0.98, 1.72, 0.95, 0.40, fill=TEAL, round_=0.35)
shape_text(panL, "Auxines", 10.5, WHITE)
panR = rect(s4, 4.11, 1.72, 0.95, 0.40, fill=GREEN_E, round_=0.35)
shape_text(panR, "Cytokinines", 10.5, WHITE)
txtA = tb(s4, 0.78, 3.12, 5.5, 1.3,
          [{'runs': [{'t': "Le ratio ", 'sz': 11.5, 'c': CHARCOAL},
                     {'t': "Auxine / Cytokinine", 'sz': 11.5, 'c': TEAL, 'b': True},
                     {'t': " influence le développement et le devenir cellulaire.", 'sz': 11.5, 'c': CHARCOAL}],
            'line': 1.15, 'space_after': 5},
           {'runs': [{'t': "Exigences variables selon l'espèce, l'explant et le protocole de culture.",
                      'sz': 10.5, 'c': GREY_TX, 'i': True}], 'line': 1.1}])

cardB = rect(s4, 0.5, 4.78, 6.05, 2.05, fill=WHITE, line=SAGE, line_w=0.75, round_=0.05)
accB = rect(s4, 0.5, 4.78, 0.075, 2.05, fill=GREEN_E)
lblB = tb(s4, 0.78, 4.92, 4.2, 0.35,
          [{'runs': [{'t': "2 · LE CHOIX DE L'EXPLANT", 'sz': 12, 'c': GREEN_D, 'b': True, 'spc': 150}]}])
bulB = bullets(s4, 0.78, 5.32, 4.15, 1.35, [
    [{'t': "Importance des ", 'sz': 11.5}, {'t': "tissus juvéniles", 'sz': 11.5, 'b': True, 'c': GREEN_D}, {'t': ".", 'sz': 11.5}],
    [{'t': "Exemples : ", 'sz': 11.5}, {'t': "embryons immatures", 'sz': 11.5, 'b': True, 'c': GREEN_D},
     {'t': " et ", 'sz': 11.5}, {'t': "apex", 'sz': 11.5, 'b': True, 'c': GREEN_D}, {'t': ".", 'sz': 11.5}],
], gap=6)
phB = pic(s4, f"{A}/callus_tube.jpg", 5.05, 4.95, 1.32, 1.72, line=SAGE, line_w=1)
capB = tb(s4, 5.05, 6.68, 1.32, 0.28,
          [{'runs': [{'t': "Culture in vitro", 'sz': 8, 'c': GREY_TX, 'i': True}], 'align': PP_ALIGN.CENTER}])

cardC = rect(s4, 6.80, 1.42, 6.03, 3.60, fill=WHITE, line=SAGE, line_w=0.75, round_=0.05)
accC = rect(s4, 6.80, 1.42, 0.075, 3.60, fill=TERRA)
lblC = tb(s4, 7.08, 1.56, 5.6, 0.35,
          [{'runs': [{'t': "3 · LE DÉFI MAJEUR : LE BRUNISSEMENT", 'sz': 12, 'c': TERRA, 'b': True, 'spc': 120}]}])
bulC = bullets(s4, 7.08, 1.98, 5.5, 1.05, [
    [{'t': "Oxydation des ", 'sz': 11.5}, {'t': "polyphénols", 'sz': 11.5, 'b': True, 'c': TERRA},
     {'t': " libérés par les tissus blessés.", 'sz': 11.5}],
    [{'t': "Risque de toxicité et de perte de viabilité des cultures.", 'sz': 11.5}],
], gap=5, dot=TERRA)
cal_ok = oval(s4, 7.25, 3.10, 1.05, 1.05, fill=SAGE, line=GREEN_E, line_w=1.5)
tb(s4, 7.05, 4.20, 1.45, 0.3, [{'runs': [{'t': "Cal sain", 'sz': 9.5, 'c': GREEN_D, 'b': True}], 'align': PP_ALIGN.CENTER}])
arr = shape(s4, MSO_SHAPE.RIGHT_ARROW, 8.48, 3.47, 0.55, 0.32, fill=TERRA)
cal_ko = oval(s4, 9.20, 3.10, 1.05, 1.05, fill=BROWN, line=TERRA, line_w=1.5)
tb(s4, 8.95, 4.20, 1.55, 0.3, [{'runs': [{'t': "Cal bruni", 'sz': 9.5, 'c': TERRA, 'b': True}], 'align': PP_ALIGN.CENTER}])
sol_lbl = tb(s4, 10.55, 2.98, 2.1, 0.3,
             [{'runs': [{'t': "SOLUTIONS :", 'sz': 10, 'c': GOLD, 'b': True, 'spc': 150}]}])
sol1 = rect(s4, 10.55, 3.32, 2.10, 0.42, fill=GOLD, round_=0.5)
shape_text(sol1, "Charbon actif", 10.5, WHITE)
sol2 = rect(s4, 10.55, 3.84, 2.10, 0.42, fill=GOLD, round_=0.5)
shape_text(sol2, "Acide ascorbique", 10.5, WHITE)

cardD = rect(s4, 6.80, 5.22, 6.03, 1.61, fill=WHITE, line=SAGE, line_w=0.75, round_=0.05)
accD = rect(s4, 6.80, 5.22, 0.075, 1.61, fill=TERRA)
lblD = tb(s4, 7.08, 5.36, 5.6, 0.35,
          [{'runs': [{'t': "4 · RISQUE DE VARIATIONS SOMACLONALES", 'sz': 12, 'c': TERRA, 'b': True, 'spc': 120}]}])
bulD = bullets(s4, 7.08, 5.76, 5.55, 1.0, [
    [{'t': "Modifications ", 'sz': 11.5}, {'t': "génétiques ou épigénétiques", 'sz': 11.5, 'b': True, 'c': GREEN_D},
     {'t': " possibles durant la culture in vitro.", 'sz': 11.5}],
    [{'t': "Risque d'altération de la conformité clonale.", 'sz': 11.5}],
], gap=5, dot=TERRA)

footer(s4, 4)
notes(s4, """En tant que futurs ingénieurs, il est crucial de comprendre que cette technique ne réussit pas automatiquement.

Le facteur numéro un est l'équilibre hormonal. C'est le dosage des phytohormones qui influence l'orientation du développement cellulaire.

Le deuxième défi majeur, particulièrement chez certaines espèces comme le palmier dattier, est le brunissement. Les tissus blessés libèrent des composés phénoliques qui peuvent s'oxyder et endommager les cultures. Pour limiter ce phénomène, on peut utiliser du charbon actif dans les milieux ou des antioxydants comme l'acide ascorbique.

Enfin, il faut surveiller le risque de variation somaclonale, c'est-à-dire des modifications génétiques ou épigénétiques qui peuvent apparaître sous l'effet des conditions de culture in vitro et altérer la conformité du clone.""")

add_transition(s4, 'wipe_r')
add_animations(s4, [
    S(cardA, 'fade', 'click', 0, 400), S(accA, 'wipe_l', 'with', 0, 400), S(lblA, 'fade', 'with', 0, 400),
    S(panL, 'fade', 'after', 100, 400), S(panR, 'fade', 'with', 0, 400),
    S(beam, 'wipe_l', 'after', 100, 400), S(fulcrum, 'fade', 'with', 0, 400),
    S(base, 'wipe_l', 'with', 0, 400), S(wireL, 'wipe_b', 'with', 0, 300), S(wireR, 'wipe_b', 'with', 0, 300),
    S(txtA, 'fade', 'after', 150, 500),
    S(cardB, 'fade', 'click', 0, 400), S(accB, 'wipe_l', 'with', 0, 400), S(lblB, 'fade', 'with', 0, 400),
    S(bulB, 'fade', 'after', 100, 500), S(phB, 'fade', 'after', 100, 500), S(capB, 'fade', 'with', 0, 400),
    S(cardC, 'fade', 'click', 0, 400), S(accC, 'wipe_l', 'with', 0, 400), S(lblC, 'fade', 'with', 0, 400),
    S(bulC, 'fade', 'after', 100, 500),
    S(cal_ok, 'fade', 'after', 150, 400),
    S(arr, 'wipe_l', 'after', 150, 350),
    S(cal_ko, 'fade', 'after', 100, 400),
    S(sol_lbl, 'fade', 'after', 200, 300),
    S(sol1, 'zoom', 'after', 100, 400), S(sol2, 'zoom', 'after', 120, 400),
    S(cardD, 'fade', 'click', 0, 400), S(accD, 'wipe_l', 'with', 0, 400), S(lblD, 'fade', 'with', 0, 400),
    S(bulD, 'fade', 'after', 100, 500),
])

# ================================================================ SLIDE 5
s5 = new_slide()
rect(s5, 0, 0, 13.334, 7.5, fill=IVORY)
header(s5, "Applications concrètes : cas pratiques", "Deux exemples issus de la littérature scientifique")

pL = rect(s5, 0.5, 1.42, 6.05, 5.42, fill=GREEN_D, round_=0.04)
hL = tb(s5, 0.80, 1.62, 5.4, 0.75,
        [{'runs': [{'t': "Le palmier dattier", 'sz': 18, 'c': WHITE, 'b': True, 'f': F_TITLE}], 'space_after': 2},
         {'runs': [{'t': "Lutte contre le Bayoud", 'sz': 12.5, 'c': GOLD, 'b': True}]}])
phL = pic(s5, f"{A}/oasis_algeria.jpg", 0.80, 2.42, 5.45, 1.95, line=GOLD, line_w=1.25)
bulL = bullets(s5, 0.80, 4.55, 5.45, 1.9, [
    [{'t': "Multiplication de clones présentant une ", 'sz': 11.5, 'c': IVORY},
     {'t': "résistance au Bayoud", 'sz': 11.5, 'c': WHITE, 'b': True}, {'t': ".", 'sz': 11.5, 'c': IVORY}],
    [{'t': "Exemple cité : variété ", 'sz': 11.5, 'c': IVORY},
     {'t': "Degla Beida", 'sz': 11.5, 'c': WHITE, 'b': True}, {'t': ".", 'sz': 11.5, 'c': IVORY}],
    [{'t': "Maladie associée au champignon tellurique ", 'sz': 11.5, 'c': IVORY},
     {'t': "Fusarium oxysporum", 'sz': 11.5, 'c': WHITE, 'b': True, 'i': True}, {'t': ".", 'sz': 11.5, 'c': IVORY}],
    [{'t': "Intérêt : ", 'sz': 11.5, 'c': IVORY},
     {'t': "multiplication rapide", 'sz': 11.5, 'c': WHITE, 'b': True},
     {'t': " du matériel végétal.", 'sz': 11.5, 'c': IVORY}],
], gap=6, dot=GOLD)
phF = pic(s5, f"{A}/fusarium.jpg", 5.00, 5.55, 1.25, 1.05, line=WHITE, line_w=1.25)
capF = tb(s5, 3.30, 6.68, 2.95, 0.28,
          [{'runs': [{'t': "Culture de Fusarium sp. au laboratoire", 'sz': 8.5, 'c': SAGE, 'i': True}],
            'align': PP_ALIGN.RIGHT}])
chipL = rect(s5, 4.85, 1.62, 1.40, 0.40, fill=GOLD, round_=0.5)
shape_text(chipL, "Algérie", 10.5, WHITE)

pR = rect(s5, 6.80, 1.42, 6.03, 5.42, fill=WHITE, line=SAGE, line_w=0.75, round_=0.04)
hR = tb(s5, 7.10, 1.62, 5.4, 0.75,
        [{'runs': [{'t': "Sylviculture & reforestation", 'sz': 18, 'c': GREEN_D, 'b': True, 'f': F_TITLE}], 'space_after': 2},
         {'runs': [{'t': "Les conifères — production de plants d'élite", 'sz': 12.5, 'c': TEAL, 'b': True}]}])
phR = pic(s5, f"{A}/spruce_seedling.jpg", 7.10, 2.42, 5.43, 1.95, line=TEAL, line_w=1.25)
bulR = bullets(s5, 7.10, 4.55, 5.43, 1.9, [
    [{'t': "Production ", 'sz': 11.5}, {'t': "à grande échelle", 'sz': 11.5, 'b': True, 'c': GREEN_D},
     {'t': " de conifères d'élite.", 'sz': 11.5}],
    [{'t': "Exemple : ", 'sz': 11.5}, {'t': "épinette blanche", 'sz': 11.5, 'b': True, 'c': GREEN_D},
     {'t': " (", 'sz': 11.5}, {'t': "Picea glauca", 'sz': 11.5, 'i': True, 'c': TEAL}, {'t': ").", 'sz': 11.5}],
    [{'t': "Référence d'application : ", 'sz': 11.5},
     {'t': "pépinière de Saint-Modeste, Québec", 'sz': 11.5, 'b': True, 'c': GREEN_D}, {'t': ".", 'sz': 11.5}],
    [{'t': "Intérêt : croissance, qualité du bois et reboisement.", 'sz': 11.5}],
], gap=6, dot=TEAL)
chipR = rect(s5, 11.30, 1.62, 1.40, 0.40, fill=TEAL, round_=0.5)
shape_text(chipR, "Québec", 10.5, WHITE)

footer(s5, 5)
notes(s5, """Pourquoi utilisons-nous cette technologie à si grande échelle ? J'aimerais vous donner deux exemples concrets tirés de la littérature scientifique.

Premièrement, la sauvegarde du palmier dattier. La maladie du Bayoud représente une menace importante pour les palmeraies au Maghreb. La multiplication traditionnelle par rejets est limitée en vitesse et en quantité. L'embryogenèse somatique peut permettre une multiplication à grande échelle de matériel végétal sélectionné, notamment de génotypes présentant une résistance au Bayoud. La résistance de chaque variété ou clone doit cependant être confirmée par des évaluations appropriées.

Deuxièmement, la foresterie. Au Québec, des pépinières et programmes de production utilisent des techniques de multiplication végétative et d'embryogenèse somatique pour produire des conifères sélectionnés, notamment l'épinette blanche. Cette technologie peut contribuer à la multiplication de matériel génétique d'intérêt pour le reboisement et la production forestière.""")

add_transition(s5, 'push_up')
add_animations(s5, [
    S(pL, 'fade', 'click', 0, 400), S(hL, 'fade', 'with', 0, 400),
    S(phL, 'fade', 'after', 150, 600),
    S(bulL, 'fade', 'after', 200, 500),
    S(phF, 'zoom', 'after', 200, 500), S(capF, 'fade', 'with', 0, 400),
    S(chipL, 'zoom', 'after', 150, 400),
    S(pR, 'fade', 'click', 0, 400), S(hR, 'fade', 'with', 0, 400),
    S(phR, 'fade', 'after', 150, 600),
    S(bulR, 'fade', 'after', 200, 500),
    S(chipR, 'zoom', 'after', 150, 400),
])

# ================================================================ SLIDE 6
s6 = new_slide()
rect(s6, 0, 0, 13.334, 7.5, fill=IVORY)
header(s6, "Avantages, limites et perspectives", "Un bilan critique et les technologies de demain")

cardAv = rect(s6, 0.5, 1.42, 6.05, 2.72, fill=WHITE, line=SAGE, line_w=0.75, round_=0.05)
hdAv = rect(s6, 0.5, 1.42, 6.05, 0.44, fill=GREEN_E, round_=0.05)
tb(s6, 0.78, 1.50, 5.5, 0.32,
   [{'runs': [{'t': "AVANTAGES", 'sz': 12.5, 'c': WHITE, 'b': True, 'spc': 250}]}])
bulAv = bullets(s6, 0.78, 2.02, 4.35, 2.0, [
    [{'t': "Taux de multiplication potentiellement très élevé.", 'sz': 10.8}],
    [{'t': "Structure bipolaire : tige et racine.", 'sz': 10.8}],
    [{'t': "Potentiel d'automatisation.", 'sz': 10.8}],
    [{'t': "Adaptation à la culture en bioréacteurs.", 'sz': 10.8}],
    [{'t': "Production clonale de matériel sélectionné.", 'sz': 10.8}],
], gap=3.5, dot=GREEN_E)
phBr = pic(s6, f"{A}/bioreactor.jpg", 5.22, 2.02, 1.18, 1.95, line=SAGE, line_w=1)
capBr = tb(s6, 5.10, 3.99, 1.42, 0.28,
           [{'runs': [{'t': "Bioréacteur de laboratoire", 'sz': 8, 'c': GREY_TX, 'i': True}],
             'align': PP_ALIGN.CENTER}])

cardLim = rect(s6, 0.5, 4.34, 6.05, 2.50, fill=WHITE, line=SAGE, line_w=0.75, round_=0.05)
hdLim = rect(s6, 0.5, 4.34, 6.05, 0.44, fill=TERRA, round_=0.05)
tb(s6, 0.78, 4.42, 5.5, 0.32,
   [{'runs': [{'t': "LIMITES ET DÉFIS", 'sz': 12.5, 'c': WHITE, 'b': True, 'spc': 250}]}])
bulLim = bullets(s6, 0.78, 4.94, 5.5, 1.8, [
    [{'t': "Coût de production élevé.", 'sz': 10.8}],
    [{'t': "Besoin d'infrastructures de laboratoire.", 'sz': 10.8}],
    [{'t': "Protocoles complexes et parfois longs.", 'sz': 10.8}],
    [{'t': "Risque de variations somaclonales.", 'sz': 10.8}],
    [{'t': "Acclimatation parfois délicate.", 'sz': 10.8}],
], gap=3.5, dot=TERRA)

cardFut = rect(s6, 6.80, 1.42, 6.03, 5.42, fill=WHITE, line=TEAL, line_w=1.0, round_=0.05)
hdFut = rect(s6, 6.80, 1.42, 6.03, 0.44, fill=TEAL, round_=0.05)
tb(s6, 7.08, 1.50, 5.5, 0.32,
   [{'runs': [{'t': "PERSPECTIVES FUTURES", 'sz': 12.5, 'c': WHITE, 'b': True, 'spc': 250}]}])

lblCryo = tb(s6, 7.08, 2.02, 5.4, 0.32,
             [{'runs': [{'t': "Cryoconservation", 'sz': 13.5, 'c': GREEN_D, 'b': True, 'f': F_TITLE}]}])
phCryo = pic(s6, f"{A}/cryo_dewar.jpg", 7.08, 2.42, 2.05, 1.72, line=TEAL, line_w=1)
bulCryo = bullets(s6, 9.30, 2.42, 3.35, 1.75, [
    [{'t': "Conservation à très basse température.", 'sz': 10.8}],
    [{'t': "Azote liquide ≈ ", 'sz': 10.8}, {'t': "−196 °C", 'sz': 10.8, 'b': True, 'c': TEAL}, {'t': ".", 'sz': 10.8}],
    [{'t': "Réussite dépendante du matériel végétal et du protocole.", 'sz': 10.8, 'i': True, 'c': GREY_TX}],
], gap=4, dot=TEAL)

lblSyn = tb(s6, 7.08, 4.42, 5.4, 0.32,
            [{'runs': [{'t': "Semences artificielles (Synseeds)", 'sz': 13.5, 'c': GREEN_D, 'b': True, 'f': F_TITLE}]}])
seed_out = oval(s6, 7.45, 4.92, 1.15, 1.15, fill=SAGE, line=TEAL, line_w=1.25, alpha=70)
seed_in = oval(s6, 7.72, 5.19, 0.61, 0.61, fill=GREEN_E, line=WHITE, line_w=1.25)
shoot = rect(s6, 8.00, 4.62, 0.05, 0.32, fill=GREEN_D)
leafL = shape(s6, MSO_SHAPE.TEAR, 7.78, 4.52, 0.24, 0.24, fill=GREEN_E, rot=225)
leafR = shape(s6, MSO_SHAPE.TEAR, 8.03, 4.52, 0.24, 0.24, fill=GREEN_E, rot=315)
capSyn = tb(s6, 7.08, 6.14, 1.9, 0.5,
            [{'runs': [{'t': "Embryon encapsulé dans un gel d'alginate", 'sz': 8, 'c': GREY_TX, 'i': True}],
              'align': PP_ALIGN.CENTER, 'line': 1.05}])
bulSyn = bullets(s6, 9.30, 4.92, 3.35, 1.6, [
    [{'t': "Encapsulation d'embryons somatiques.", 'sz': 10.8}],
    [{'t': "Utilisation possible d'un ", 'sz': 10.8}, {'t': "gel d'alginate", 'sz': 10.8, 'b': True, 'c': TEAL}, {'t': ".", 'sz': 10.8}],
    [{'t': "Manipulation et propagation facilitées.", 'sz': 10.8}],
], gap=4, dot=TEAL)

footer(s6, 6)
notes(s6, """Pour résumer la valeur de cette technique, comparons-la à l'organogenèse classique.

L'embryogenèse somatique possède l'avantage de produire une structure bipolaire, avec un axe comprenant les structures caulinaires et racinaires. Cela peut faciliter la régénération d'une plante entière sans passer par une étape d'enracinement séparée dans certains protocoles.

De plus, elle peut être adaptée aux bioréacteurs pour une production à grande échelle.

Cependant, elle reste coûteuse en termes d'infrastructure, de main-d'œuvre et de maîtrise technique.

L'avenir de cette technique comprend notamment la cryoconservation à environ -196 degrés Celsius dans l'azote liquide, qui permet de conserver du matériel végétal pendant de longues périodes, ainsi que les semences artificielles, où l'embryon somatique est encapsulé, par exemple dans un gel d'alginate, afin de faciliter sa manipulation et son utilisation.""")

add_transition(s6, 'fade')
add_animations(s6, [
    S(cardAv, 'fade', 'click', 0, 400), S(hdAv, 'wipe_l', 'with', 0, 400),
    S(bulAv, 'fade', 'after', 100, 500),
    S(phBr, 'fade', 'after', 100, 500), S(capBr, 'fade', 'with', 0, 400),
    S(cardLim, 'fade', 'click', 0, 400), S(hdLim, 'wipe_l', 'with', 0, 400),
    S(bulLim, 'fade', 'after', 100, 500),
    S(cardFut, 'fade', 'click', 0, 400), S(hdFut, 'wipe_l', 'with', 0, 400),
    S(lblCryo, 'fade', 'after', 100, 400),
    S(phCryo, 'fade', 'after', 100, 500), S(bulCryo, 'fade', 'with', 0, 500),
    S(lblSyn, 'fade', 'after', 200, 400),
    S(seed_out, 'zoom', 'after', 100, 450), S(seed_in, 'zoom', 'with', 0, 450),
    S(shoot, 'wipe_b', 'after', 100, 300), S(leafL, 'zoom', 'with', 0, 300), S(leafR, 'zoom', 'with', 0, 300),
    S(capSyn, 'fade', 'after', 100, 400), S(bulSyn, 'fade', 'with', 0, 500),
])

# ================================================================ SLIDE 7
s7 = new_slide()
ph7 = pic(s7, f"{A}/finale_cultures.jpg", 0, 0, 13.334, 7.5)
ov7 = rect(s7, 0, 0, 13.334, 7.5, fill=GREEN_D, alpha=72)
kicker = tb(s7, 1.9, 1.55, 9.5, 0.45,
            [{'runs': [{'t': "CONCLUSION", 'sz': 14, 'c': GOLD, 'b': True, 'spc': 450}],
              'align': PP_ALIGN.CENTER}])
concl = tb(s7, 1.9, 2.25, 9.5, 1.6,
           [{'runs': [{'t': "L'embryogenèse somatique est un outil biotechnologique incontournable pour l'agriculture et la foresterie modernes face aux défis climatiques et sanitaires.",
                       'sz': 21, 'c': WHITE, 'b': False}], 'align': PP_ALIGN.CENTER, 'line': 1.25}])
rule7 = rect(s7, 5.87, 4.15, 1.6, 0.05, fill=GOLD)
merci = tb(s7, 1.9, 4.45, 9.5, 0.75,
           [{'runs': [{'t': "Merci de votre attention !", 'sz': 30, 'c': WHITE, 'b': True, 'f': F_TITLE}],
             'align': PP_ALIGN.CENTER}])
quest = tb(s7, 1.9, 5.35, 9.5, 0.5,
           [{'runs': [{'t': "Questions ?", 'sz': 17, 'c': SAGE, 'i': True}], 'align': PP_ALIGN.CENTER}])
foot7 = tb(s7, 1.9, 6.75, 9.5, 0.35,
           [{'runs': [{'t': "Biotechnologie Appliquée — 3ème Année Ingénieur — Sciences Agronomiques",
                       'sz': 10, 'c': SAGE}], 'align': PP_ALIGN.CENTER}])

notes(s7, """En conclusion, l'embryogenèse somatique est bien plus qu'une simple curiosité de laboratoire. C'est un outil d'ingénierie agronomique majeur pour faire face aux défis actuels : changement climatique, émergence de nouvelles maladies et besoin de reboisement massif.

Elle permet de multiplier du matériel végétal sélectionné, de contribuer à la conservation de ressources génétiques et d'ouvrir de nouvelles perspectives pour l'agriculture et la foresterie.

Je vous remercie pour votre attention. Je suis maintenant à votre disposition si vous avez des questions.""")

add_transition(s7, 'fade_black')
add_animations(s7, [
    S(ph7, 'fade', 'auto', 0, 800), S(ov7, 'fade', 'with', 0, 800),
    S(kicker, 'fade', 'after', 300, 500),
    S(concl, 'fade', 'after', 300, 700),
    S(rule7, 'wipe_l', 'after', 200, 500),
    S(merci, 'fade', 'after', 300, 600),
    S(quest, 'fade', 'after', 250, 500),
    S(foot7, 'fade', 'after', 200, 400),
])

# ================================================================ save
OUT = "Embryogenese_Somatique_Presentation.pptx"
prs.save(OUT)
print("saved", OUT, "slides:", len(prs.slides.__iter__.__self__._sldIdLst))
