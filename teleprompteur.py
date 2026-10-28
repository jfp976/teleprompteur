#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Téléprompteur — affichage défilant en gros caractères Arial.

Touches :
  Molette souris   vitesse de défilement (vitesse du karaoké quand il est actif)
  Espace           pause / reprise
  ↑ / ↓            paragraphe précédent / suivant
  + / -            agrandir / réduire les caractères
  *                mode karaoké (le texte lu se colore en jaune)
  ← / →            karaoké moins vite / plus vite
  Clic             modifier le texte (Échap pour revenir au défilement)
  F1 aide · F2 voix on/off · F3 ouvrir · F4 langue · F11 plein écran · Échap quitter

Dépendances : voir requirements.txt (pypdf, python-docx, pywin32, edge-tts, pygame).
"""

import os
import sys
import re
import math
import bisect
import unicodedata
import time
import queue
import asyncio
import tempfile
import threading
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, font as tkfont

os.environ.setdefault("PYGAME_HIDE_SUPPORT_PROMPT", "1")

LANGUES = ["fr", "en", "es", "ar"]

# --------------------------------------------------------------------------
#  Textes de l'interface
# --------------------------------------------------------------------------
TXT = {
    "fr": {
        "titre": "Téléprompteur",
        "nom_langue": "Français",
        "code": "FR",
        "vitesse": "Vitesse :",
        "mpm": "mots/min",
        "ligne": "Ligne",
        "pause": "⏸ Pause",
        "lecture": "▶ Lecture",
        "fin": "■ Fin",
        "edition": "✎ Édition (Échap)",
        "karaoke": "Karaoké",
        "son": "Voix",
        "on": "ON", "off": "OFF",
        "taille": "Taille",
        "raccourcis": "Espace pause · Molette vitesse · ↑↓ paragraphe · Début/Fin · PgPréc/PgSuiv · +/− taille · * karaoké · ←/→ vit. karaoké · Clic modifier · F1 aide · F2 voix · F3 ouvrir · F4 langue · F11 écran · Échap quitter",
        "ouvrir": "Ouvrir un texte",
        "documents": "Documents (txt, pdf, doc, docx)",
        "tous": "Tous les fichiers",
        "texte_brut": "Texte (*.txt)",
        "word": "Word (*.docx)",
        "enregistrer": "Enregistrer le texte modifié",
        "modifie": "Le texte a été modifié.\nVoulez-vous enregistrer les modifications ?",
        "quitter": "Quitter le téléprompteur ?",
        "erreur": "Erreur",
        "err_lecture": "Impossible de lire le fichier :\n{}",
        "err_module": "Le module « {} » n'est pas installé.\nInstallez-le avec :\n\npip install {}",
        "err_doc": "Lecture des fichiers .doc impossible : Microsoft Word n'est pas disponible "
                   "(et antiword non plus).\nEnregistrez le document au format .docx puis réessayez.",
        "err_vide": "Aucun texte n'a pu être extrait de ce fichier (PDF scanné ?).",
        "err_ecriture": "Impossible d'enregistrer :\n{}",
        "chargement": "Chargement…",
        "aucune_voix": "aucune voix de synthèse disponible",
        "voix_repli": "Voix neuronale indisponible (Internet ?) — voix Windows utilisée",
        "aide_titre": "Téléprompteur — Aide",
        "aide": [
            ("Molette de la souris", "régler la vitesse de défilement (et celle du karaoké quand il est actif)"),
            ("Barre d'espace", "pause / reprise de la lecture"),
            ("Flèche ↑", "réafficher le paragraphe précédent"),
            ("Flèche ↓", "afficher le paragraphe suivant"),
            ("Touche +", "agrandir les caractères"),
            ("Touche −", "réduire les caractères"),
            ("Touche *", "mode karaoké : le texte se colore progressivement en jaune"),
            ("Flèche ←", "karaoké moins rapide"),
            ("Flèche →", "karaoké plus rapide"),
            ("Clic dans le texte", "modifier le texte à cet endroit (la lecture se met en pause) ; "
                                   "Échap pour revenir au défilement"),
            ("Début", "revenir au début du texte"),
            ("Fin", "aller à la fin du texte"),
            ("Page préc.", "remonter d'un écran (début d'écran)"),
            ("Page suiv.", "descendre d'un écran (bas d'écran)"),
            ("F1", "afficher cette aide"),
            ("F2", "activer / couper la voix de synthèse"),
            ("F3", "ouvrir un fichier (.txt, .pdf, .doc, .docx)"),
            ("F4", "changer de langue : Français → English → Español → العربية"),
            ("F11", "plein écran / fenêtre"),
            ("Échap", "quitter (l'enregistrement des modifications est proposé)"),
        ],
        "aide_pied": "La ligne d'état, en bas de l'écran, indique la vitesse de lecture imposée en "
                     "mots par minute, le numéro de la ligne en cours de lecture (repérée par la flèche "
                     "jaune) et le pourcentage d'avancement dans le texte.\n\n"
                     "Fermer l'aide : Échap ou F1.",
        "accueil": [
            "Bienvenue dans le téléprompteur.",
            "Appuyez sur F3 pour ouvrir un texte au format txt, pdf, doc ou docx, puis sur la barre "
            "d'espace pour lancer le défilement.",
            "La molette de la souris règle la vitesse. Les flèches haut et bas permettent de passer "
            "d'un paragraphe à l'autre. Les touches plus et moins changent la taille des caractères.",
            "La touche étoile active le mode karaoké : le texte se colore en jaune au fil de la "
            "lecture, à une vitesse réglable avec les flèches gauche et droite.",
            "F2 active la voix de synthèse, F4 change de langue, et F1 affiche l'aide complète.",
            "Pour corriger le texte, il suffit de cliquer dessus. La touche Échap ramène au défilement.",
        ],
    },
    "en": {
        "titre": "Teleprompter",
        "nom_langue": "English",
        "code": "EN",
        "vitesse": "Speed:",
        "mpm": "words/min",
        "ligne": "Line",
        "pause": "⏸ Paused",
        "lecture": "▶ Playing",
        "fin": "■ End",
        "edition": "✎ Editing (Esc)",
        "karaoke": "Karaoke",
        "son": "Voice",
        "on": "ON", "off": "OFF",
        "taille": "Size",
        "raccourcis": "Space pause · Wheel speed · ↑↓ paragraph · Home/End · PgUp/PgDn · +/− size · * karaoke · ←/→ karaoke speed · Click edit · F1 help · F2 voice · F3 open · F4 language · F11 screen · Esc quit",
        "ouvrir": "Open a text",
        "documents": "Documents (txt, pdf, doc, docx)",
        "tous": "All files",
        "texte_brut": "Text (*.txt)",
        "word": "Word (*.docx)",
        "enregistrer": "Save the edited text",
        "modifie": "The text has been modified.\nDo you want to save your changes?",
        "quitter": "Quit the teleprompter?",
        "erreur": "Error",
        "err_lecture": "Unable to read the file:\n{}",
        "err_module": "The “{}” module is not installed.\nInstall it with:\n\npip install {}",
        "err_doc": "Cannot read .doc files: Microsoft Word is not available (nor antiword).\n"
                   "Save the document as .docx and try again.",
        "err_vide": "No text could be extracted from this file (scanned PDF?).",
        "err_ecriture": "Unable to save:\n{}",
        "chargement": "Loading…",
        "aucune_voix": "no speech voice available",
        "voix_repli": "Neural voice unavailable (Internet?) — using Windows voice",
        "aide_titre": "Teleprompter — Help",
        "aide": [
            ("Mouse wheel", "adjust scrolling speed (and karaoke speed when it is on)"),
            ("Space bar", "pause / resume"),
            ("Up arrow", "show the previous paragraph"),
            ("Down arrow", "show the next paragraph"),
            ("+ key", "larger characters"),
            ("− key", "smaller characters"),
            ("* key", "karaoke mode: the text is progressively coloured yellow"),
            ("Left arrow", "slower karaoke"),
            ("Right arrow", "faster karaoke"),
            ("Click in the text", "edit the text at that point (reading pauses); Esc to resume scrolling"),
            ("Home", "go to the beginning of the text"),
            ("End", "go to the end of the text"),
            ("Page Up", "go back one screen (top of screen)"),
            ("Page Down", "go forward one screen (bottom of screen)"),
            ("F1", "show this help"),
            ("F2", "turn the speech voice on / off"),
            ("F3", "open a file (.txt, .pdf, .doc, .docx)"),
            ("F4", "change language: Français → English → Español → العربية"),
            ("F11", "full screen / window"),
            ("Esc", "quit (you will be offered to save your changes)"),
        ],
        "aide_pied": "The status line at the bottom of the screen shows the imposed reading speed in "
                     "words per minute, the number of the line being read (marked by the yellow arrows)"
                     " and the progress through the text as a percentage.\n\n"
                     "Close help: Esc or F1.",
        "accueil": [
            "Welcome to the teleprompter.",
            "Press F3 to open a txt, pdf, doc or docx file, then press the space bar to start scrolling.",
            "The mouse wheel sets the speed. The up and down arrows move from one paragraph to "
            "another. The plus and minus keys change the size of the characters.",
            "The star key turns on karaoke mode: the text turns yellow as you read, at a speed you "
            "set with the left and right arrows.",
            "F2 turns on the speech voice, F4 changes the language and F1 shows the full help.",
            "To correct the text, just click on it. The Esc key brings you back to scrolling.",
        ],
    },
    "es": {
        "titre": "Teleprompter",
        "nom_langue": "Español",
        "code": "ES",
        "vitesse": "Velocidad:",
        "mpm": "palabras/min",
        "ligne": "Línea",
        "pause": "⏸ Pausa",
        "lecture": "▶ Lectura",
        "fin": "■ Fin",
        "edition": "✎ Edición (Esc)",
        "karaoke": "Karaoke",
        "son": "Voz",
        "on": "SÍ", "off": "NO",
        "taille": "Tamaño",
        "raccourcis": "Espacio pausa · Rueda velocidad · ↑↓ párrafo · Inicio/Fin · RePág/AvPág · +/− tamaño · * karaoke · ←/→ vel. karaoke · Clic editar · F1 ayuda · F2 voz · F3 abrir · F4 idioma · F11 pantalla · Esc salir",
        "ouvrir": "Abrir un texto",
        "documents": "Documentos (txt, pdf, doc, docx)",
        "tous": "Todos los archivos",
        "texte_brut": "Texto (*.txt)",
        "word": "Word (*.docx)",
        "enregistrer": "Guardar el texto modificado",
        "modifie": "El texto ha sido modificado.\n¿Desea guardar los cambios?",
        "quitter": "¿Salir del teleprompter?",
        "erreur": "Error",
        "err_lecture": "No se puede leer el archivo:\n{}",
        "err_module": "El módulo «{}» no está instalado.\nInstálelo con:\n\npip install {}",
        "err_doc": "No se pueden leer archivos .doc: Microsoft Word no está disponible (ni antiword).\n"
                   "Guarde el documento como .docx e inténtelo de nuevo.",
        "err_vide": "No se ha podido extraer texto de este archivo (¿PDF escaneado?).",
        "err_ecriture": "No se puede guardar:\n{}",
        "chargement": "Cargando…",
        "aucune_voix": "ninguna voz de síntesis disponible",
        "voix_repli": "Voz neuronal no disponible (¿Internet?) — se usa la voz de Windows",
        "aide_titre": "Teleprompter — Ayuda",
        "aide": [
            ("Rueda del ratón", "ajustar la velocidad de desplazamiento (y la del karaoke cuando está activo)"),
            ("Barra espaciadora", "pausa / reanudar"),
            ("Flecha ↑", "volver al párrafo anterior"),
            ("Flecha ↓", "pasar al párrafo siguiente"),
            ("Tecla +", "letra más grande"),
            ("Tecla −", "letra más pequeña"),
            ("Tecla *", "modo karaoke: el texto se colorea de amarillo progresivamente"),
            ("Flecha ←", "karaoke más lento"),
            ("Flecha →", "karaoke más rápido"),
            ("Clic en el texto", "modificar el texto en ese punto (la lectura se pausa); "
                                 "Esc para volver al desplazamiento"),
            ("Inicio", "volver al principio del texto"),
            ("Fin", "ir al final del texto"),
            ("Re Pág", "retroceder una pantalla (principio de pantalla)"),
            ("Av Pág", "avanzar una pantalla (final de pantalla)"),
            ("F1", "mostrar esta ayuda"),
            ("F2", "activar / desactivar la voz de síntesis"),
            ("F3", "abrir un archivo (.txt, .pdf, .doc, .docx)"),
            ("F4", "cambiar de idioma: Français → English → Español → العربية"),
            ("F11", "pantalla completa / ventana"),
            ("Esc", "salir (se ofrece guardar los cambios)"),
        ],
        "aide_pied": "La línea de estado, en la parte inferior de la pantalla, indica la velocidad de "
                     "lectura impuesta en palabras por minuto, el número de la línea que se está leyendo "
                     "(señalada por las flechas amarillas) y el porcentaje de avance en el texto.\n\n"
                     "Cerrar la ayuda: Esc o F1.",
        "accueil": [
            "Bienvenido al teleprompter.",
            "Pulse F3 para abrir un archivo txt, pdf, doc o docx y luego la barra espaciadora para "
            "iniciar el desplazamiento.",
            "La rueda del ratón ajusta la velocidad. Las flechas arriba y abajo pasan de un párrafo a "
            "otro. Las teclas más y menos cambian el tamaño de la letra.",
            "La tecla asterisco activa el modo karaoke: el texto se vuelve amarillo a medida que se "
            "lee, a una velocidad que se ajusta con las flechas izquierda y derecha.",
            "F2 activa la voz de síntesis, F4 cambia el idioma y F1 muestra la ayuda completa.",
            "Para corregir el texto, basta con hacer clic en él. La tecla Esc vuelve al desplazamiento.",
        ],
    },
    "ar": {
        "titre": "الملقّن",
        "nom_langue": "العربية",
        "code": "AR",
        "vitesse": "السرعة:",
        "mpm": "كلمة/دقيقة",
        "ligne": "سطر",
        "pause": "⏸ إيقاف مؤقت",
        "lecture": "▶ قراءة",
        "fin": "■ النهاية",
        "edition": "✎ تحرير (Esc)",
        "karaoke": "كاريوكي",
        "son": "صوت",
        "on": "تشغيل", "off": "إيقاف",
        "taille": "الحجم",
        "raccourcis": "مسافة إيقاف · العجلة السرعة · ↑↓ فقرة · Home/End البداية/النهاية · PgUp/PgDn شاشة · +/− الحجم · * كاريوكي · ←/→ سرعة الكاريوكي · نقرة تحرير · F1 مساعدة · F2 صوت · F3 فتح · F4 اللغة · F11 ملء الشاشة · Esc خروج",
        "ouvrir": "فتح نص",
        "documents": "مستندات (txt, pdf, doc, docx)",
        "tous": "كل الملفات",
        "texte_brut": "نص (*.txt)",
        "word": "Word (*.docx)",
        "enregistrer": "حفظ النص المعدَّل",
        "modifie": "تم تعديل النص.\nهل تريد حفظ التعديلات؟",
        "quitter": "هل تريد الخروج من الملقّن؟",
        "erreur": "خطأ",
        "err_lecture": "تعذّرت قراءة الملف:\n{}",
        "err_module": "الوحدة «{}» غير مثبّتة.\nثبّتها بالأمر:\n\npip install {}",
        "err_doc": "تعذّرت قراءة ملفات doc لأن برنامج Microsoft Word غير متوفر.\n"
                   "احفظ المستند بصيغة docx ثم أعد المحاولة.",
        "err_vide": "لم يُستخرج أي نص من هذا الملف (ملف PDF ممسوح ضوئيًا؟).",
        "err_ecriture": "تعذّر الحفظ:\n{}",
        "chargement": "جارٍ التحميل…",
        "aucune_voix": "لا يتوفر أي صوت تركيبي",
        "voix_repli": "الصوت العصبي غير متاح (الإنترنت؟) — يُستخدم صوت Windows",
        "aide_titre": "الملقّن — مساعدة",
        "aide": [
            ("عجلة الفأرة", "ضبط سرعة التمرير (وسرعة الكاريوكي عند تفعيله)"),
            ("مفتاح المسافة", "إيقاف القراءة مؤقتًا / استئنافها"),
            ("السهم ↑", "عرض الفقرة السابقة"),
            ("السهم ↓", "عرض الفقرة التالية"),
            ("Home", "العودة إلى بداية النص"),
            ("End", "الانتقال إلى نهاية النص"),
            ("Page Up", "الرجوع شاشةً واحدة (بداية الشاشة)"),
            ("Page Down", "التقدّم شاشةً واحدة (أسفل الشاشة)"),
            ("المفتاح +", "تكبير الحروف"),
            ("المفتاح −", "تصغير الحروف"),
            ("المفتاح *", "وضع الكاريوكي: يتلوّن النص بالأصفر تدريجيًا، من اليمين إلى اليسار في النص العربي"),
            ("السهم ←", "إبطاء الكاريوكي"),
            ("السهم →", "تسريع الكاريوكي"),
            ("النقر على النص", "تعديل النص في هذا الموضع (تتوقف القراءة مؤقتًا)؛ Esc للعودة إلى التمرير"),
            ("F1", "عرض هذه المساعدة"),
            ("F2", "تشغيل الصوت التركيبي / إيقافه"),
            ("F3", "فتح ملف (txt, pdf, doc, docx)"),
            ("F4", "تغيير اللغة: Français ← English ← Español ← العربية"),
            ("F11", "ملء الشاشة / نافذة"),
            ("Esc", "الخروج (مع اقتراح حفظ التعديلات)"),
        ],
        "aide_pied": "يعرض سطر الحالة في أسفل الشاشة سرعة القراءة المفروضة بالكلمات في الدقيقة، "
                     "ورقم السطر الجاري قراءته (المشار إليه بالسهمين الأصفرين)، ونسبة التقدّم في النص.\n\n"
                     "إغلاق المساعدة: Esc أو F1.",
        "accueil": [
            "مرحبًا بكم في الملقّن.",
            "اضغط F3 لفتح نص بصيغة txt أو pdf أو doc أو docx، ثم اضغط مفتاح المسافة لبدء التمرير.",
            "تضبط عجلة الفأرة السرعة. وينقلك السهمان إلى الأعلى وإلى الأسفل من فقرة إلى أخرى، "
            "ويغيّر المفتاحان زائد وناقص حجم الحروف.",
            "يفعّل مفتاح النجمة وضع الكاريوكي: يتلوّن النص بالأصفر من اليمين إلى اليسار أثناء القراءة، "
            "بسرعة تُضبط بالسهمين الأيمن والأيسر.",
            "يفعّل F2 الصوت التركيبي، ويغيّر F4 اللغة، ويعرض F1 المساعدة الكاملة.",
            "لتصحيح النص يكفي النقر عليه، ويعيدك المفتاح Esc إلى التمرير.",
        ],
    },
}


class ErreurLecture(Exception):
    pass


class ModuleManquant(Exception):
    def __init__(self, module, paquet):
        super().__init__(module)
        self.module, self.paquet = module, paquet


# --------------------------------------------------------------------------
#  Lecture des fichiers
# --------------------------------------------------------------------------
def reformer_paragraphes(texte):
    """Recolle les lignes coupées (PDF) en vrais paragraphes."""
    paras, cour = [], ""
    for ligne in texte.replace("\r", "").split("\n"):
        s = ligne.strip()
        if not s:
            if cour:
                paras.append(cour)
                cour = ""
            continue
        if not cour:
            cour = s
        elif cour.endswith("-") and not cour.endswith(" -") and s[:1].islower():
            cour = cour[:-1] + s
        elif re.search(r"[.!?:»\"…]$", cour) and (s[:1].isupper() or s[:1] in "«\"—–-•·"):
            paras.append(cour)
            cour = s
        else:
            cour += " " + s
    if cour:
        paras.append(cour)
    return "\n".join(paras)


def lire_txt(chemin):
    with open(chemin, "rb") as f:
        brut = f.read()
    for enc in ("utf-8-sig", "utf-16") if brut[:2] in (b"\xff\xfe", b"\xfe\xff") else ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return brut.decode(enc)
        except UnicodeDecodeError:
            continue
    return brut.decode("latin-1", errors="replace")


def lire_pdf(chemin):
    try:
        from pypdf import PdfReader
    except ImportError:
        raise ModuleManquant("pypdf", "pypdf")
    lecteur = PdfReader(chemin)
    pages = [(p.extract_text() or "") for p in lecteur.pages]
    return reformer_paragraphes("\n\n".join(pages))


def lire_docx(chemin):
    try:
        import docx
    except ImportError:
        raise ModuleManquant("python-docx", "python-docx")
    doc = docx.Document(chemin)
    return "\n".join(p.text for p in doc.paragraphs)


def lire_doc(chemin):
    if sys.platform == "win32":
        try:
            import win32com.client
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            try:
                d = word.Documents.Open(os.path.abspath(chemin), ReadOnly=True,
                                        AddToRecentFiles=False, ConfirmConversions=False)
                texte = d.Content.Text
                d.Close(False)
            finally:
                word.Quit()
            return texte.replace("\r\x07", "\n").replace("\r", "\n").replace("\x0b", "\n")
        except Exception:
            pass
    try:
        sortie = subprocess.run(["antiword", "-w", "0", chemin], capture_output=True, check=True)
        return sortie.stdout.decode("utf-8", errors="replace")
    except Exception:
        raise ErreurLecture("doc")


def lire_fichier(chemin):
    ext = os.path.splitext(chemin)[1].lower()
    if ext == ".pdf":
        t = lire_pdf(chemin)
    elif ext == ".docx":
        t = lire_docx(chemin)
    elif ext == ".doc":
        t = lire_doc(chemin)
    else:
        t = lire_txt(chemin)
    t = t.replace("\r\n", "\n").replace("\r", "\n").replace("\t", " ").replace("\f", "\n")
    t = re.sub(r"\n{3,}", "\n\n", t).strip("\n")
    return t


# --------------------------------------------------------------------------
#  Écriture de droite à gauche (arabe)
# --------------------------------------------------------------------------
# Tk n'applique pas l'algorithme bidirectionnel : on calcule nous-mêmes la forme
# visuelle des lignes (lettres liées par arabic-reshaper, mots et lettres remis dans
# l'ordre d'affichage), et l'on garde la correspondance « caractère du texte →
# colonne affichée » pour que le karaoké avance de droite à gauche.

ARABE = re.compile(r"[؀-ۿݐ-ݿࢠ-ࣿﭐ-﷿ﹰ-﻿]")
MIROIR = str.maketrans("()[]{}<>«»", ")(][}{><»«")


def est_arabe(texte):
    """Vrai si le texte est majoritairement écrit en alphabet arabe."""
    lettres = [c for c in texte[:20000] if c.isalpha()]
    if not lettres:
        return False
    return sum(1 for c in lettres if ARABE.match(c)) > 0.4 * len(lettres)


_FACONNEUR = None


def faconneur():
    global _FACONNEUR
    if _FACONNEUR is None:
        try:
            import arabic_reshaper
        except ImportError:
            raise ModuleManquant("arabic-reshaper", "arabic-reshaper")
        _FACONNEUR = arabic_reshaper.ArabicReshaper(
            configuration={"delete_harakat": False, "support_ligatures": False})
    return _FACONNEUR


def _mot_visuel(mot):
    """Retourne (texte affiché, position affichée de chaque caractère, grappes affichées)."""
    if not ARABE.search(mot):
        return mot, list(range(len(mot))), [], False
    forme = faconneur().reshape(mot)
    if len(forme) != len(mot):                      # sécurité : pas de correspondance 1 pour 1
        forme = mot
    # grappes = lettre + signes diacritiques qui la suivent
    grappes, i = [], 0
    while i < len(forme):
        j = i + 1
        while j < len(forme) and unicodedata.combining(forme[j]):
            j += 1
        grappes.append((i, j))
        i = j
    visuel, pos, sorties, col = [], [0] * len(mot), [], 0
    for a, b in reversed(grappes):
        sorties.append((col, col + (b - a)))
        for k in range(a, b):
            pos[k] = col + (k - a)
            visuel.append(forme[k].translate(MIROIR))
        col += b - a
    return "".join(visuel), pos, sorties, True


def mise_en_page_rtl(source, mesurer, largeur):
    """Coupe le texte en lignes affichables de droite à gauche.

    Retourne :
      lignes : [(texte affiché, début logique, fin logique, début de paragraphe, fin de paragraphe)]
      mots   : [(début logique, fin logique, n° de ligne (1…), [colonne de chaque caractère])]
      grappes: {n° de ligne: [(colonne début, colonne fin), …]} pour les lettres arabes
    """
    espace = mesurer(" ")
    lignes, mots, grappes = [], [], {}
    decalage = 0
    for brute in source.split("\n"):
        jetons = []
        for m in re.finditer(r"\S+", brute):
            vis, pos, gr, rtl = _mot_visuel(m.group())
            jetons.append({"debut": decalage + m.start(), "fin": decalage + m.end(),
                           "vis": vis, "pos": pos, "gr": gr, "rtl": rtl, "l": mesurer(vis)})
        decalage += len(brute) + 1
        if not jetons:
            continue
        # répartition en lignes (premier mot logique = à droite)
        paquets, cour, l_cour = [], [], 0
        for j in jetons:
            besoin = j["l"] if not cour else l_cour + espace + j["l"]
            if cour and besoin > largeur:
                paquets.append(cour)
                cour, l_cour = [j], j["l"]
            else:
                cour.append(j)
                l_cour = besoin
        paquets.append(cour)
        for n, paquet in enumerate(paquets):
            # suites de mots latins/chiffres gardées dans leur ordre, le reste inversé
            series = []
            for j in paquet:
                if not j["rtl"] and series and not series[-1][0]["rtl"]:
                    series[-1].append(j)
                else:
                    series.append([j])
            ordre = [j for serie in reversed(series) for j in serie]
            no = len(lignes) + 1
            morceaux, col, gr_ligne = [], 0, []
            for j in ordre:
                if morceaux:
                    morceaux.append(" ")
                    col += 1
                morceaux.append(j["vis"])
                mots.append((j["debut"], j["fin"], no, [col + p for p in j["pos"]]))
                gr_ligne += [(col + a, col + b) for a, b in j["gr"]]
                col += len(j["vis"])
            grappes[no] = gr_ligne
            lignes.append(("".join(morceaux), paquet[0]["debut"], paquet[-1]["fin"],
                           n == 0, n == len(paquets) - 1))
    mots.sort(key=lambda m: m[0])
    return lignes, mots, grappes


# --------------------------------------------------------------------------
#  Voix de synthèse
# --------------------------------------------------------------------------
class Voix:
    """Voix neuronale (edge-tts) si disponible, sinon voix Windows (SAPI), sinon pyttsx3."""

    NEURONALES = {"fr": ("fr-FR-DeniseNeural", "Denise"),
                  "en": ("en-GB-SoniaNeural", "Sonia"),
                  "es": ("es-ES-ElviraNeural", "Elvira"),
                  "ar": ("ar-SA-ZariyahNeural", "Zariyah")}
    SAPI_PREF = {"fr": ["virginie", "hortense", "julie", "denise"],
                 "en": ["zira", "hazel", "susan", "jenny", "aria", "david"],
                 "es": ["helena", "laura", "sabina", "elvira", "pablo"],
                 "ar": ["hoda", "salma", "zariyah", "naayf", "hamed"]}
    SAPI_LCID = {"fr": {"40c", "c0c", "80c", "100c"},
                 "en": {"409", "809", "c09", "1009", "1409"},
                 "es": {"c0a", "40a", "80a", "2c0a", "540a"},
                 "ar": {"401", "801", "c01", "1001", "1401", "1801", "1c01", "2001", "2401", "2801",
                        "2c01", "3001", "3401", "3801", "3c01", "4001"}}

    def __init__(self, signaler):
        self.signaler = signaler          # exécute une fonction dans le fil principal
        self.lang = "fr"
        self.moteur = None
        self.nom = "—"
        self.message = None
        self._jeton = 0
        self._etat = "repos"
        self._debut = 0.0
        self._fichiers = []
        self._pyx_fini = False
        self._pyx_moteur = None
        self._pyx_pos = 0
        self._decalage = 0                # espaces retirés en tête du texte prononcé
        self._longueur = 0
        self._mots = []                   # edge : (début ms, durée ms, car. début, car. fin)
        self.edge_ok = False
        self.sapi = None
        self.jetons_sapi = []
        self.pyx_ok = False

        try:
            import edge_tts  # noqa: F401
            import pygame
            pygame.mixer.init()
            self.pygame = pygame
            self.edge_ok = True
        except Exception:
            self.pygame = None

        if sys.platform == "win32":
            try:
                import win32com.client
                self.sapi = win32com.client.Dispatch("SAPI.SpVoice")
                self.jetons_sapi = self._lister_sapi()
            except Exception:
                self.sapi = None

        if not self.sapi:
            try:
                import pyttsx3  # noqa: F401
                self.pyx_ok = True
            except Exception:
                pass

        self.choisir_langue("fr")

    # --- choix de la voix ---------------------------------------------------
    def _lister_sapi(self):
        import win32com.client
        jetons, vus = [], set()
        sources = [self.sapi.GetVoices()]
        try:
            cat = win32com.client.Dispatch("SAPI.SpObjectTokenCategory")
            cat.SetId(r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices", False)
            sources.append(cat.EnumerateTokens())
        except Exception:
            pass
        for src in sources:
            for i in range(src.Count):
                j = src.Item(i)
                try:
                    d = j.GetDescription()
                except Exception:
                    continue
                if d not in vus:
                    vus.add(d)
                    jetons.append(j)
        return jetons

    def _choisir_sapi(self, lang):
        meilleur = None
        for pref in self.SAPI_PREF[lang]:
            for j in self.jetons_sapi:
                if pref in j.GetDescription().lower():
                    meilleur = j
                    break
            if meilleur:
                break
        if not meilleur:
            for j in self.jetons_sapi:
                try:
                    codes = {c.strip().lower().lstrip("0") for c in j.GetAttribute("Language").split(";")}
                except Exception:
                    continue
                if codes & self.SAPI_LCID[lang]:
                    meilleur = j
                    break
        if meilleur:
            try:
                self.sapi.Voice = meilleur
                return meilleur.GetDescription()
            except Exception:
                pass
        return None

    def choisir_langue(self, lang):
        self.lang = lang
        nom_sapi = self._choisir_sapi(lang) if self.sapi else None
        if self.edge_ok:
            self.moteur, self.nom = "edge", self.NEURONALES[lang][1] + " (neuronale)"
        elif self.sapi and nom_sapi:
            self.moteur, self.nom = "sapi", nom_sapi
        elif self.sapi:
            self.moteur, self.nom = "sapi", "Windows"
        elif self.pyx_ok:
            self.moteur, self.nom = "pyttsx3", "pyttsx3"
        else:
            self.moteur, self.nom = None, None

    # --- vitesse ------------------------------------------------------------
    @staticmethod
    def _rate_edge(mpm):
        pct = int(round((mpm / 165.0 - 1) * 100))
        return f"{max(-50, min(100, pct)):+d}%"

    @staticmethod
    def _rate_sapi(mpm):
        return max(-10, min(10, int(round(10 * math.log(mpm / 170.0) / math.log(3)))))

    # --- lecture ------------------------------------------------------------
    def parler(self, texte, mpm):
        self.arreter()
        self._decalage = len(texte) - len(texte.lstrip())
        texte = texte.strip()
        self._longueur = len(texte)
        self._mots = []
        self._pyx_pos = 0
        if not texte or not self.moteur:
            return
        mpm = max(80, min(320, mpm or 150))
        self._jeton += 1
        j = self._jeton
        self._debut = time.perf_counter()
        if self.moteur == "edge":
            self._etat = "generation"
            threading.Thread(target=self._generer_edge, daemon=True,
                             args=(texte, j, self.NEURONALES[self.lang][0], self._rate_edge(mpm), mpm)).start()
        elif self.moteur == "sapi":
            try:
                self.sapi.Rate = self._rate_sapi(mpm)
                self.sapi.Speak(texte, 3)          # asynchrone + purge
                self._etat = "lecture"
            except Exception:
                self._etat = "repos"
        elif self.moteur == "pyttsx3":
            self._pyx_fini = False
            self._etat = "lecture"
            threading.Thread(target=self._pyx, args=(texte, j, mpm), daemon=True).start()

    def _generer_edge(self, texte, j, voix, rate, mpm):
        try:
            import edge_tts
            fd, chemin = tempfile.mkstemp(prefix="telep_", suffix=".mp3")
            os.close(fd)

            bruts = []

            async def gen():
                try:
                    comm = edge_tts.Communicate(texte, voix, rate=rate, boundary="WordBoundary")
                except TypeError:                       # anciennes versions d'edge-tts
                    comm = edge_tts.Communicate(texte, voix, rate=rate)
                with open(chemin, "wb") as f:
                    async for morceau in comm.stream():
                        if morceau["type"] == "audio":
                            f.write(morceau["data"])
                        elif morceau["type"] == "WordBoundary":
                            bruts.append((morceau["offset"] / 1e4, morceau["duration"] / 1e4, morceau["text"]))

            asyncio.run(gen())
            if os.path.getsize(chemin) == 0:
                raise IOError("vide")
            mots = self._aligner(texte, bruts)
            self.signaler(lambda: self._jouer(chemin, j, mots))
        except Exception:
            self.signaler(lambda: self._repli(texte, j, mpm))

    @staticmethod
    def _aligner(texte, bruts):
        """Retrouve la position de chaque mot prononcé dans le texte."""
        mots, curseur, bas = [], 0, texte.lower()
        for debut, duree, mot in bruts:
            mot = (mot or "").strip()
            if not mot:
                continue
            i = bas.find(mot.lower(), curseur)
            if i < 0 or i - curseur > 80:
                continue
            mots.append((debut, max(duree, 1.0), i, i + len(mot)))
            curseur = i + len(mot)
        return mots

    def _jouer(self, chemin, j, mots=None):
        if j != self._jeton:
            self._supprimer(chemin)
            return
        self._mots = mots or []
        try:
            self.pygame.mixer.music.load(chemin)
            self.pygame.mixer.music.play()
            self._fichiers.append(chemin)
            self._etat = "lecture"
            self._debut = time.perf_counter()
        except Exception:
            self._etat = "repos"

    def _repli(self, texte, j, mpm):
        if j != self._jeton:
            return
        self.edge_ok = False
        self.choisir_langue(self.lang)
        self.message = "voix_repli"
        if self.moteur:
            self.parler(texte, mpm)
        else:
            self._etat = "repos"

    def _pyx(self, texte, j, mpm):
        try:
            import pyttsx3
            eng = pyttsx3.init()
            self._pyx_moteur = eng
            for v in eng.getProperty("voices"):
                infos = (v.id + " " + (v.name or "") + " " + " ".join(map(str, v.languages or []))).lower()
                if self.lang in infos:
                    eng.setProperty("voice", v.id)
                    break
            eng.setProperty("rate", mpm)

            def mot(name, location, length):
                if j == self._jeton:
                    self._pyx_pos = location + length
            try:
                eng.connect("started-word", mot)
            except Exception:
                pass
            eng.say(texte)
            eng.runAndWait()
        except Exception:
            pass
        if j == self._jeton:
            self._pyx_fini = True

    def arreter(self):
        self._jeton += 1
        if self._etat == "repos" and self.moteur != "edge":
            return
        self._etat = "repos"
        try:
            if self.moteur == "edge" and self.pygame:
                self.pygame.mixer.music.stop()
                try:
                    self.pygame.mixer.music.unload()
                except Exception:
                    pass
                self._nettoyer()
            elif self.moteur == "sapi":
                self.sapi.Speak("", 3)
            elif self.moteur == "pyttsx3" and self._pyx_moteur:
                self._pyx_moteur.stop()
        except Exception:
            pass

    def occupe(self):
        return self._etat != "repos"

    def position(self):
        """Nombre de caractères déjà prononcés dans le texte passé à parler(), ou None
        si la voix n'a pas encore commencé (génération en cours) ou est arrêtée."""
        if self._etat != "lecture":
            return None
        try:
            if self.moteur == "edge":
                t = self.pygame.mixer.music.get_pos()       # ms depuis le début de la lecture
                if t < 0:
                    return None
                p = 0
                for debut, duree, a, b in self._mots:
                    if debut > t:
                        break
                    p = b if t >= debut + duree else a + (b - a) * (t - debut) / duree
                return self._decalage + p
            if self.moteur == "sapi":
                st = self.sapi.Status
                if st.RunningState == 1:
                    return self._decalage + self._longueur
                return self._decalage + st.InputWordPosition + st.InputWordLength
            return self._decalage + self._pyx_pos
        except Exception:
            return None

    def vient_de_finir(self):
        """Appelée régulièrement : True quand l'énoncé en cours s'est terminé normalement."""
        if self._etat != "lecture" or time.perf_counter() - self._debut < 0.4:
            return False
        try:
            if self.moteur == "edge":
                fini = not self.pygame.mixer.music.get_busy()
            elif self.moteur == "sapi":
                fini = self.sapi.Status.RunningState == 1
            else:
                fini = self._pyx_fini
        except Exception:
            fini = True
        if fini:
            self._etat = "repos"
            if self.moteur == "edge":
                try:
                    self.pygame.mixer.music.unload()
                except Exception:
                    pass
                self._nettoyer()
        return fini

    @staticmethod
    def _supprimer(chemin):
        try:
            os.remove(chemin)
        except OSError:
            pass

    def _nettoyer(self):
        reste = []
        for f in self._fichiers:
            try:
                os.remove(f)
            except OSError:
                reste.append(f)
        self._fichiers = reste

    def fermer(self):
        self.arreter()
        if self.pygame:
            try:
                self.pygame.mixer.quit()
            except Exception:
                pass
        self._nettoyer()


# --------------------------------------------------------------------------
#  Application
# --------------------------------------------------------------------------
class Teleprompteur:
    TICK_MS = 16
    FOND = "#000000"
    ENCRE = "#ffffff"
    JAUNE = "#ffd400"
    GRIS = "#8c8c8c"

    def __init__(self, root, chemin=None):
        self.root = root
        self.lang = "fr"
        self.taille = 56
        self.vitesse_px = 45.0       # défilement, pixels / seconde
        self.vitesse_car = 14.0      # karaoké, caractères / seconde
        self.pause = True
        self.fin = False
        self.karaoke = False
        self.pos_kara = 0.0
        self._kara_int = 0
        self.edition = False
        self.son = False
        self.chemin = None
        self.accueil = True
        self.rtl = False             # texte arabe affiché de droite à gauche
        self.source = ""             # texte logique (ordre de lecture)
        self.rtl_mots, self.rtl_debuts, self.rtl_lignes = [], [], [(0, 0)]
        self.ligne_voix = None       # n° du paragraphe en cours de lecture vocale
        self.base_voix = 0           # position (car.) dans le texte du début de l'énoncé en cours
        self._relance_id = None
        self.plein_ecran = False
        self._reste = 0.0
        self._dernier = time.perf_counter()
        self._prochain_etat = 0.0
        self._stats = None
        self._evts = queue.Queue()
        self._aide = None
        self._config_id = None
        self._derniere_activite = time.perf_counter()

        self._construire()
        self.voix = Voix(self._signaler)
        self._charger_texte("\n".join(TXT[self.lang]["accueil"]), accueil=True)
        if chemin:
            self.root.after(200, lambda: self.ouvrir(chemin))
        self.root.after(self.TICK_MS, self._tick)

    # ------------------------------------------------------------------ UI
    def t(self, cle):
        return TXT[self.lang][cle]

    def _construire(self):
        r = self.root
        r.title(self.t("titre"))
        r.configure(bg=self.FOND)
        r.geometry("1200x800")
        try:
            r.state("zoomed")
        except tk.TclError:
            try:
                r.attributes("-zoomed", True)
            except tk.TclError:
                pass

        self.police = tkfont.Font(family="Arial", size=self.taille)
        self.police_etat = tkfont.Font(family="Arial", size=11, weight="bold")

        self.police_raccourcis = tkfont.Font(family="Arial", size=10)
        self._metriques()

        cadre = tk.Frame(r, bg=self.FOND)
        cadre.pack(side="top", fill="both", expand=True)

        self.reglette = tk.Canvas(cadre, width=46, bg=self.FOND, highlightthickness=0, bd=0)
        self.reglette.pack(side="left", fill="y")
        self.reglette_d = tk.Canvas(cadre, width=46, bg=self.FOND, highlightthickness=0, bd=0)
        self.reglette_d.pack(side="right", fill="y")

        # Barre du bas posée par-dessus le texte : la masquer ne décale pas la lecture.
        self.bas = tk.Frame(r, bg="#1c1c1c")
        self.barre = tk.Label(self.bas, anchor="w", bg="#1c1c1c", fg="#e6e6e6", font=self.police_etat,
                              padx=6, pady=5)
        self.barre.pack(side="left")
        self.ligne_raccourcis = tk.Label(self.bas, anchor="w", justify="left", bg="#1c1c1c", fg="#9a9a9a",
                                         font=self.police_raccourcis, padx=4, pady=4)
        self.ligne_raccourcis.pack(side="left", fill="x", expand=True)
        self.bas.bind("<Configure>", lambda e: self._ajuster_raccourcis())
        self.bas_visible = False
        self._afficher_bas(True)

        self.texte = tk.Text(cadre, wrap="word", font=self.police, bg=self.FOND, fg=self.ENCRE,
                             bd=0, highlightthickness=0, padx=40, pady=0, insertwidth=0,
                             insertbackground=self.JAUNE, selectbackground="#3a5a8c",
                             undo=True, maxundo=-1, cursor="arrow")
        self.texte.pack(side="left", fill="both", expand=True)

        self.texte.tag_configure("para")
        self.texte.tag_configure("alire", foreground=self.GRIS)
        self.texte.tag_configure("lu", foreground=self.JAUNE)
        self.texte.tag_configure("marge_haut")
        self.texte.tag_configure("marge_bas")

        t = self.texte
        t.bind("<Key>", self._touche)
        t.bind("<Button-1>", self._clic)
        t.bind("<MouseWheel>", self._molette)
        t.bind("<Button-4>", lambda e: self._molette(e, +1))
        t.bind("<Button-5>", lambda e: self._molette(e, -1))
        t.bind("<Configure>", self._redimension)
        for seq in ("<B1-Motion>", "<Double-Button-1>", "<Triple-Button-1>", "<Button-2>", "<Button-3>"):
            t.bind(seq, lambda e: None if self.edition else "break")
        t.bind("<Motion>", lambda e: self._activite())
        for w in (self.reglette, self.reglette_d, self.bas, self.barre, self.ligne_raccourcis):
            w.bind("<MouseWheel>", self._molette)
            w.bind("<Button-4>", lambda e: self._molette(e, +1))
            w.bind("<Button-5>", lambda e: self._molette(e, -1))
            w.bind("<Motion>", lambda e: self._activite())
        r.protocol("WM_DELETE_WINDOW", self.quitter)
        t.focus_set()

    # --- barre du bas : visible à l'arrêt, masquée après quelques secondes de défilement
    DELAI_MASQUAGE = 4.0

    def _activite(self):
        self._derniere_activite = time.perf_counter()
        if not self.bas_visible:
            self._afficher_bas(True)

    def _afficher_bas(self, visible):
        if visible == self.bas_visible:
            return
        self.bas_visible = visible
        if visible:
            self.bas.place(relx=0, rely=1.0, anchor="sw", relwidth=1.0)
            self.bas.lift()
        else:
            self.bas.place_forget()

    def _gerer_masquage(self, maintenant):
        en_cours = not (self.pause or self.edition or self.fin)
        self._afficher_bas(not en_cours or maintenant - self._derniere_activite < self.DELAI_MASQUAGE)

    def _ajuster_raccourcis(self):
        place = self.bas.winfo_width() - self.barre.winfo_reqwidth() - 6
        self.ligne_raccourcis.configure(wraplength=max(150, place))

    def _signaler(self, fonction):
        """Appelable depuis n'importe quel fil : exécute `fonction` dans le fil Tk."""
        self._evts.put(fonction)

    # ---------------------------------------------------- géométrie / marges
    @property
    def y_guide(self):
        return int(self.hauteur_ligne * 0.75) + 14     # au niveau de la première ligne

    def _metriques(self):
        self.hauteur_ligne = self.police.metrics("linespace")
        self.ascent = self.police.metrics("ascent")

    def _appliquer_mise_en_forme(self):
        """Marges, interlignes et (en arabe) recalcul des lignes affichées."""
        t = self.texte
        self._metriques()
        if self.rtl and not self.edition:
            self._afficher_rtl()
        h = max(t.winfo_height(), 200)
        yg = self.y_guide
        para = int(self.hauteur_ligne * 0.45)
        inter = int(self.hauteur_ligne * 0.08)
        t.configure(spacing3=para, spacing2=inter)
        t.tag_configure("suite", spacing3=inter)            # lignes arabes d'un même paragraphe
        for tag in ("marge_haut", "marge_bas"):
            t.tag_remove(tag, "1.0", "end")
        t.tag_configure("marge_haut", spacing1=max(0, yg - int(self.hauteur_ligne * 0.55)))
        t.tag_configure("marge_bas", spacing3=max(0, h - yg - int(self.hauteur_ligne * 0.5)))
        t.tag_add("marge_haut", "1.0", "1.0 lineend +1c")
        t.tag_add("marge_bas", "end-1c linestart", "end")
        t.tag_raise("marge_haut")
        t.tag_raise("marge_bas")
        self._stats = None
        self._dessiner_guide()

    def _dessiner_guide(self):
        y = self.y_guide
        s = max(10, min(22, self.hauteur_ligne // 4))
        g, d = self.reglette, self.reglette_d
        g.delete("all")
        d.delete("all")
        g.create_polygon(8, y - s, 8 + int(s * 1.5), y, 8, y + s, fill=self.JAUNE, outline="")
        x = int(d.cget("width")) - 8
        d.create_polygon(x, y - s, x - int(s * 1.5), y, x, y + s, fill=self.JAUNE, outline="")

    def _redimension(self, _e=None):
        if self._config_id:
            self.root.after_cancel(self._config_id)

        def appliquer():
            self._config_id = None
            if self.edition:
                return
            pos = self._pos_lecture()
            self._appliquer_mise_en_forme()
            self._placer(self._idx(pos))
            if self.karaoke:
                self._kara_depuis(int(self.pos_kara))
        self._config_id = self.root.after(120, appliquer)

    # --------------------------------------------------- texte et affichage
    def _texte(self):
        """Texte logique (ordre de lecture), quel que soit le sens d'affichage."""
        if self.rtl and not self.edition:
            return self.source
        return self.texte.get("1.0", "end-1c")

    def _installer_texte(self, contenu):
        """Remplace le texte affiché (le drapeau « modifié » est conservé)."""
        t = self.texte
        modifie = t.edit_modified()
        self.source = contenu
        self.rtl = est_arabe(contenu)
        if self.rtl:
            try:
                faconneur()
            except ModuleManquant as e:
                self.rtl = False
                messagebox.showerror(self.t("erreur"), self.t("err_module").format(e.module, e.paquet),
                                     parent=self.root)
        t.tag_configure("droite", justify="right")
        if not self.rtl:
            t.delete("1.0", "end")
            t.insert("1.0", contenu)
        self._appliquer_mise_en_forme()          # (re)construit les lignes arabes
        t.edit_reset()
        t.edit_modified(modifie)

    def _afficher_rtl(self):
        t = self.texte
        modifie = t.edit_modified()
        largeur = max(200, t.winfo_width() - 2 * int(t.cget("padx")) - 30)
        lignes, self.rtl_mots, grappes = mise_en_page_rtl(self.source, self.police.measure, largeur)
        self.rtl_lignes = [(a, b) for _, a, b, _, _ in lignes] or [(0, 0)]
        self.rtl_debuts = [m[0] for m in self.rtl_mots]
        t.delete("1.0", "end")
        t.insert("1.0", "\n".join(l[0] for l in lignes))
        t.tag_configure("droite", justify="right", wrap="none")
        t.tag_add("droite", "1.0", "end")
        # Chaque lettre arabe forme son propre segment d'affichage (balises alternées),
        # ce qui empêche le système de la réordonner une seconde fois.
        for no, liste in grappes.items():
            for k, tag in ((0, "g0"), (1, "g1")):
                plages = []
                for a, b in liste[k::2]:
                    plages += [f"{no}.{a}", f"{no}.{b}"]
                if plages:
                    t.tag_add(tag, *plages)
        for no, l in enumerate(lignes, 1):
            if not l[4]:
                t.tag_add("suite", f"{no}.0", f"{no}.end +1c")
        t.edit_reset()
        t.edit_modified(modifie)

    # ------------------------------------------------- positions logiques
    # Une « position » est un nombre de caractères depuis le début du texte logique.
    def _nb(self, a, b, quoi="chars"):
        r = self.texte.count(a, b, quoi)
        if not r:
            return 0
        return r[0] if isinstance(r, tuple) else int(r)

    def _idx(self, pos):
        """Position logique → index Tk affiché."""
        pos = int(pos)
        if not (self.rtl and not self.edition):
            return self.texte.index(f"1.0 + {pos} chars")
        if not self.rtl_mots:
            return "1.0"
        k = bisect.bisect_right(self.rtl_debuts, pos) - 1
        if k < 0:
            k = 0
        debut, fin, ligne, cols = self.rtl_mots[k]
        if pos >= fin and k + 1 < len(self.rtl_mots):
            debut, fin, ligne, cols = self.rtl_mots[k + 1]
            return f"{ligne}.{cols[0]}"
        i = min(max(pos - debut, 0), len(cols) - 1)
        return f"{ligne}.{cols[i]}"

    def _pos_debut_ligne(self, idx):
        """Position logique du début de la ligne affichée qui contient idx."""
        if not (self.rtl and not self.edition):
            return self._nb("1.0", self.texte.index(f"{idx} display linestart"))
        no = int(self.texte.index(idx).split(".")[0])
        no = max(1, min(no, len(self.rtl_lignes)))
        return self.rtl_lignes[no - 1][0]

    def _index_guide(self):
        return self.texte.index(f"@2,{self.y_guide}")

    def _pos_lecture(self):
        return int(self.pos_kara) if self.karaoke else self._pos_debut_ligne(self._index_guide())

    def _paragraphes(self):
        return self._stats_texte()["paras"]

    def _debut_mot_pos(self, pos):
        src = self._texte()
        while 0 < pos <= len(src) and src[pos - 1].isalnum():
            pos -= 1
        return pos

    def _centre_ligne(self, idx):
        info = self.texte.dlineinfo(idx)
        if not info:
            return None
        return info[1] + info[4] - int(self.ascent * 0.38)

    def _placer(self, idx):
        """Amène la ligne contenant `idx` à hauteur des flèches de lecture."""
        t = self.texte
        t.see(idx)
        t.update_idletasks()
        for _ in range(3):
            y = self._centre_ligne(idx)
            if y is None:
                t.see(idx)
                t.update_idletasks()
                continue
            d = y - self.y_guide
            if abs(d) < 2:
                break
            t.yview_scroll(int(d), "pixels")
            t.update_idletasks()
        self._reste = 0.0
        self.fin = False

    # ------------------------------------------------------------ chargement
    def _charger_texte(self, contenu, accueil=False):
        t = self.texte
        if hasattr(self, "voix"):
            self.voix.arreter()
        t.edit_modified(False)
        self._installer_texte(contenu)
        t.edit_modified(False)
        self.accueil = accueil
        self.pause = True
        self.fin = False
        self.pos_kara = 0.0
        t.update_idletasks()
        if self.karaoke:
            self._kara_depuis(0)
        t.yview_moveto(0)
        self._reste = 0.0
        self._maj_etat(force=True)

    def ouvrir(self, chemin=None):
        if self.edition:
            self._sortir_edition()
        if not chemin:
            self.pause = True
            self.voix.arreter()
            chemin = filedialog.askopenfilename(
                parent=self.root, title=self.t("ouvrir"),
                filetypes=[(self.t("documents"), "*.txt *.pdf *.doc *.docx"),
                           ("Texte", "*.txt"), ("PDF", "*.pdf"), ("Word", "*.doc *.docx"),
                           (self.t("tous"), "*.*")])
            self.texte.focus_set()
            if not chemin:
                return
        if not self._proposer_enregistrement():
            return
        self.barre.configure(text=self.t("chargement"))
        self.root.update_idletasks()
        try:
            contenu = lire_fichier(chemin)
        except ModuleManquant as e:
            messagebox.showerror(self.t("erreur"), self.t("err_module").format(e.module, e.paquet), parent=self.root)
            return self._maj_etat(force=True)
        except ErreurLecture:
            messagebox.showerror(self.t("erreur"), self.t("err_doc"), parent=self.root)
            return self._maj_etat(force=True)
        except Exception as e:
            messagebox.showerror(self.t("erreur"), self.t("err_lecture").format(e), parent=self.root)
            return self._maj_etat(force=True)
        if not contenu.strip():
            messagebox.showwarning(self.t("erreur"), self.t("err_vide"), parent=self.root)
            return self._maj_etat(force=True)
        self.chemin = chemin
        if est_arabe(contenu) and self.lang != "ar":
            self._definir_langue("ar")           # interface et voix en arabe
        self._charger_texte(contenu)
        self.root.title(f"{self.t('titre')} — {os.path.basename(chemin)}")
        self.texte.focus_set()

    # ---------------------------------------------------------- enregistrer
    def _enregistrer(self):
        if self.chemin:
            dossier = os.path.dirname(self.chemin)
            base, ext = os.path.splitext(os.path.basename(self.chemin))
            ext = ext.lower()
            ext = ext if ext in (".txt", ".docx") else (".docx" if ext == ".doc" else ".txt")
        else:
            dossier, base, ext = os.path.expanduser("~"), "texte", ".txt"
        types = [(self.t("texte_brut"), "*.txt"), (self.t("word"), "*.docx")]
        if ext == ".docx":
            types.reverse()
        chemin = filedialog.asksaveasfilename(parent=self.root, title=self.t("enregistrer"),
                                              initialdir=dossier, initialfile=base + ext,
                                              defaultextension=ext, filetypes=types)
        if not chemin:
            return False
        contenu = self._texte()
        try:
            if chemin.lower().endswith(".docx"):
                try:
                    import docx
                    from docx.oxml.ns import qn
                except ImportError:
                    raise ModuleManquant("python-docx", "python-docx")
                d = docx.Document()
                style = d.styles["Normal"]
                style.font.name = "Arial"
                for p in contenu.split("\n"):
                    para = d.add_paragraph(p)
                    if ARABE.search(p):                      # paragraphe de droite à gauche
                        ppr = para._p.get_or_add_pPr()
                        bidi = ppr.makeelement(qn("w:bidi"), {})
                        ppr.append(bidi)
                d.save(chemin)
            else:
                with open(chemin, "w", encoding="utf-8", newline="\r\n" if sys.platform == "win32" else "\n") as f:
                    f.write(contenu)
        except ModuleManquant as e:
            messagebox.showerror(self.t("erreur"), self.t("err_module").format(e.module, e.paquet), parent=self.root)
            return False
        except Exception as e:
            messagebox.showerror(self.t("erreur"), self.t("err_ecriture").format(e), parent=self.root)
            return False
        self.chemin = chemin
        self.texte.edit_modified(False)
        self.root.title(f"{self.t('titre')} — {os.path.basename(chemin)}")
        return True

    def _proposer_enregistrement(self):
        """True si on peut continuer (enregistré ou abandon des modifications)."""
        if not self.texte.edit_modified():
            return True
        rep = messagebox.askyesnocancel(self.t("titre"), self.t("modifie"), parent=self.root)
        if rep is None:
            return False
        if rep:
            return self._enregistrer()
        return True

    def quitter(self):
        if self.edition:
            self._sortir_edition()
        self.pause = True
        self.voix.arreter()
        if self.texte.edit_modified():
            if not self._proposer_enregistrement():
                self.texte.focus_set()
                return
        elif not messagebox.askyesno(self.t("titre"), self.t("quitter"), parent=self.root):
            self.texte.focus_set()
            return
        self.voix.fermer()
        self.root.destroy()

    # --------------------------------------------------------------- clavier
    def _touche(self, e):
        self._activite()
        k = e.keysym
        if k == "F1":
            self.afficher_aide()
            return "break"
        if k == "F2":
            self.basculer_son()
            return "break"
        if k == "F3":
            self.ouvrir()
            return "break"
        if k == "F4":
            self.changer_langue()
            return "break"
        if k == "F11":
            self.basculer_plein_ecran()
            return "break"

        if self.edition:
            if k == "Escape":
                self._sortir_edition()
                return "break"
            return None                       # saisie normale

        if k == "space":
            self.basculer_pause()
        elif k in ("Up", "KP_Up"):
            self.paragraphe(-1)
        elif k in ("Down", "KP_Down"):
            self.paragraphe(+1)
        elif k in ("plus", "KP_Add", "equal"):
            self.changer_taille(+4)
        elif k in ("minus", "KP_Subtract", "underscore"):
            self.changer_taille(-4)
        elif k in ("asterisk", "KP_Multiply"):
            self.basculer_karaoke()
        elif k in ("Left", "KP_Left"):
            self.regler_karaoke(-1)
        elif k in ("Right", "KP_Right"):
            self.regler_karaoke(+1)
        elif k in ("Home", "KP_Home"):
            self._aller(0)
        elif k in ("End", "KP_End"):
            self._aller(self._pos_debut_ligne("end-1c"))
        elif k in ("Prior", "KP_Prior"):
            self.page(-1)
        elif k in ("Next", "KP_Next"):
            self.page(+1)
        elif k == "Escape":
            if self.plein_ecran:
                self.basculer_plein_ecran()
            else:
                self.quitter()
        return "break"                        # aucune saisie hors mode édition

    def _molette(self, e, sens=None):
        self._activite()
        if self.edition:
            return None
        if sens is None:
            sens = 1 if e.delta > 0 else -1
            crans = max(1, abs(e.delta) // 120) if abs(e.delta) >= 120 else 1
        else:
            crans = 1
        if self.karaoke:
            self.vitesse_car = max(1.0, min(80.0, self.vitesse_car + sens * crans))
            self._planifier_relance()
        else:
            self.vitesse_px = max(0.0, min(900.0, self.vitesse_px + sens * crans * 5))
        self._maj_etat(force=True)
        return "break"

    # -------------------------------------------------------------- actions
    def basculer_pause(self):
        if self.fin and self.pause:
            self._aller(0)
        self.pause = not self.pause
        self._dernier = time.perf_counter()
        self._activite()
        if self.pause:
            self.voix.arreter()
        else:
            self._demarrer_voix()
        self._maj_etat(force=True)

    def _aller(self, pos):
        """Amène la position logique `pos` sur la ligne de lecture."""
        self._placer(self._idx(pos))
        if self.karaoke:
            self._kara_depuis(pos)
        self._redemarrer_voix()
        self._maj_etat(force=True)

    def _para_de(self, pos):
        """Indice du paragraphe qui contient pos (ou du dernier qui commence avant)."""
        paras = self._paragraphes()
        k = bisect.bisect_right([p[0] for p in paras], pos) - 1
        return k, paras

    def paragraphe(self, sens):
        pos = self._pos_lecture()
        k, paras = self._para_de(pos)
        if not paras:
            return
        if sens < 0:
            if k >= 0 and pos > paras[k][0]:
                cible = k                         # d'abord : début du paragraphe en cours
            else:
                cible = k - 1
        else:
            cible = k + 1
        if 0 <= cible < len(paras):
            self._aller(paras[cible][0])

    def page(self, sens):
        """Page suiv. : la ligne du bas de l'écran vient sur le repère de lecture.
        Page préc. : la ligne lue redescend en bas de l'écran (on remonte d'un écran)."""
        t = self.texte
        hauteur = t.winfo_height() - self.y_guide - self.hauteur_ligne
        if hauteur < self.hauteur_ligne:
            hauteur = self.hauteur_ligne
        ici = self._index_guide()
        bas = t.index(f"@2,{self.y_guide + hauteur}")
        n = max(1, self._nb(f"{ici} display linestart", f"{bas} display linestart", "displaylines"))
        idx = t.index(f"{ici} {'+' if sens > 0 else '-'} {n} display lines")
        self._aller(self._pos_debut_ligne(idx))

    def changer_taille(self, delta):
        nouvelle = max(16, min(220, self.taille + delta))
        if nouvelle == self.taille:
            return
        pos = self._pos_lecture()
        self.taille = nouvelle
        self.police.configure(size=nouvelle)
        self.texte.update_idletasks()
        self._appliquer_mise_en_forme()
        self._placer(self._idx(pos))
        if self.karaoke:
            self._kara_depuis(int(self.pos_kara))
        self._maj_etat(force=True)

    def basculer_karaoke(self):
        if self.karaoke:
            self.karaoke = False
            self.texte.tag_remove("lu", "1.0", "end")
            self.texte.tag_remove("alire", "1.0", "end")
        else:
            pos = self._pos_lecture()
            self.karaoke = True
            self._kara_depuis(pos)
        self._maj_etat(force=True)

    def regler_karaoke(self, sens):
        self.vitesse_car = max(1.0, min(80.0, self.vitesse_car + sens))
        self._planifier_relance()
        self._maj_etat(force=True)

    def _planifier_relance(self):
        """Karaoké + voix : après un changement de vitesse, la voix reprend au mot en cours
        avec le nouveau débit (petit délai pour regrouper plusieurs appuis)."""
        if not (self.karaoke and self.son and not self.pause and not self.edition):
            return
        if self._relance_id:
            self.root.after_cancel(self._relance_id)

        def relancer():
            self._relance_id = None
            if self.karaoke and self.son and not self.pause and not self.edition:
                self._kara_depuis(self._debut_mot_pos(int(self.pos_kara)))
                self._redemarrer_voix()
        self._relance_id = self.root.after(600, relancer)

    # ---------------------------------------------------------------- karaoké
    def _colorier(self, a, b):
        """Passe en jaune les caractères logiques a … b-1 (de droite à gauche en arabe)."""
        if b <= a:
            return
        t = self.texte
        if not (self.rtl and not self.edition):
            i, j = self._idx(a), self._idx(b)
            t.tag_add("lu", i, j)
            t.tag_remove("alire", i, j)
            return
        plages = []
        k = max(0, bisect.bisect_right(self.rtl_debuts, a) - 1)
        while k < len(self.rtl_mots):
            debut, fin, ligne, cols = self.rtl_mots[k]
            if debut >= b:
                break
            lo, hi = max(a, debut), min(b, fin)
            if hi > lo:
                if lo == debut and hi == fin:             # mot entier : une seule plage
                    plages += [f"{ligne}.{min(cols)}", f"{ligne}.{max(cols) + 1}"]
                else:
                    for i in range(lo, hi):
                        c = cols[i - debut]
                        plages += [f"{ligne}.{c}", f"{ligne}.{c + 1}"]
            k += 1
        if plages:
            t.tag_add("lu", *plages)
            t.tk.call(t._w, "tag", "remove", "alire", *plages)

    def _kara_depuis(self, pos):
        t = self.texte
        total = self._stats_texte()["car"]
        pos = max(0, min(int(pos), total))
        self.pos_kara = float(pos)
        self._kara_int = pos
        t.tag_remove("lu", "1.0", "end")
        t.tag_add("alire", "1.0", "end")
        self._colorier(0, pos)

    def basculer_son(self):
        self.son = not self.son
        if self.son:
            if not self.voix.moteur:
                self.son = False
                messagebox.showinfo(self.t("titre"), self.t("aucune_voix"), parent=self.root)
            elif not self.pause and not self.edition:
                self._demarrer_voix()
        else:
            self.voix.arreter()
        self._maj_etat(force=True)

    def _definir_langue(self, lang):
        self.lang = lang
        self.voix.choisir_langue(lang)
        titre = self.t("titre") + (f" — {os.path.basename(self.chemin)}" if self.chemin else "")
        self.root.title(titre)

    def changer_langue(self):
        self._definir_langue(LANGUES[(LANGUES.index(self.lang) + 1) % len(LANGUES)])
        if self.accueil and not self.texte.edit_modified() and not self.edition:
            self._charger_texte("\n".join(self.t("accueil")), accueil=True)
        else:
            self._redemarrer_voix()
        if self._aide and self._aide.winfo_exists():
            self._aide.destroy()
            self.afficher_aide()
        self._maj_etat(force=True)

    def basculer_plein_ecran(self):
        self.plein_ecran = not self.plein_ecran
        self.root.attributes("-fullscreen", self.plein_ecran)

    # --------------------------------------------------------------- édition
    def _clic(self, e):
        if self.edition:
            return None
        t = self.texte
        self.pause = True
        self.voix.arreter()
        if self.rtl:
            # en arabe, on corrige le texte dans son ordre naturel (aligné à droite)
            pos = self._pos_debut_ligne(t.index(f"@{e.x},{e.y}"))
            self.edition = True
            modifie = t.edit_modified()
            t.delete("1.0", "end")
            t.insert("1.0", self.source)
            t.tag_add("droite", "1.0", "end")
            t.tag_configure("droite", wrap="word")
            t.edit_reset()
            t.edit_modified(modifie)
            self._appliquer_mise_en_forme()
            t.mark_set("insert", f"1.0 + {pos} chars")
            self._placer("insert")
        self.edition = True
        t.configure(insertwidth=max(3, self.taille // 14), cursor="xterm")
        t.focus_set()
        self._maj_etat(force=True)
        return "break" if self.rtl else None  # sinon le curseur se place au point cliqué

    def _sortir_edition(self):
        t = self.texte
        pos = self._nb("1.0", "insert")
        contenu = t.get("1.0", "end-1c")
        self.edition = False
        t.configure(insertwidth=0, cursor="arrow")
        t.tag_remove("sel", "1.0", "end")
        if t.edit_modified():
            self.accueil = False
        t.tag_configure("droite", wrap="none")
        self._installer_texte(contenu)
        total = len(contenu)
        if self.karaoke:
            self.pos_kara = min(self.pos_kara, float(total))
            self._kara_depuis(int(self.pos_kara))
            pos = int(self.pos_kara)
        self._placer(self._idx(min(pos, total)))
        t.focus_set()
        self._maj_etat(force=True)

    # ------------------------------------------------------------------ voix
    def _mpm_voix(self):
        return self._stats_texte()["mpm"]

    def _demarrer_voix(self):
        if not (self.son and not self.pause and not self.edition):
            return
        pos = self._pos_lecture()
        if self.karaoke:
            pos = self._debut_mot_pos(pos)
        k, paras = self._para_de(pos)
        if k < 0 or pos >= paras[k][1]:
            self.ligne_voix = k
            return self._voix_suivante()
        self.ligne_voix = k
        self.base_voix = pos
        self.voix.parler(self._texte()[pos:paras[k][1]], self._mpm_voix())

    def _redemarrer_voix(self):
        self.voix.arreter()
        self._demarrer_voix()

    def _voix_suivante(self):
        paras = self._paragraphes()
        k = (self.ligne_voix if self.ligne_voix is not None else -1) + 1
        if k >= len(paras):
            self.ligne_voix = None
            if self.karaoke:                    # la voix a tout lu : fin du karaoké
                self.pos_kara = float(self._stats_texte()["car"])
            return
        self.ligne_voix = k
        a, b = paras[k]
        self.base_voix = a
        self.voix.parler(self._texte()[a:b], self._mpm_voix())

    # ------------------------------------------------------------ animation
    def _tick(self):
        maintenant = time.perf_counter()
        dt = min(maintenant - self._dernier, 0.1)
        self._dernier = maintenant
        try:
            while True:
                self._evts.get_nowait()()
        except queue.Empty:
            pass
        try:
            if not self.pause and not self.edition:
                if self.karaoke and self.son and self.voix.moteur:
                    self._karaoke_voix(dt)
                elif self.karaoke:
                    self._avancer_karaoke(dt)
                else:
                    self._avancer_defilement(dt)
            if self.son and not self.pause and not self.edition and self.voix.vient_de_finir():
                self._voix_suivante()
            self._gerer_masquage(maintenant)
            if maintenant >= self._prochain_etat:
                self._maj_etat()
        finally:
            self.root.after(self.TICK_MS, self._tick)

    def _defiler(self, px):
        self._reste += px
        n = int(self._reste)
        if n:
            self._reste -= n
            avant = self.texte.yview()
            self.texte.yview_scroll(n, "pixels")
            return self.texte.yview() != avant
        return True

    def _avancer_defilement(self, dt):
        if self.vitesse_px <= 0:
            return
        if not self._defiler(self.vitesse_px * dt):
            self._terminer()

    def _avancer_karaoke(self, dt):
        """Karaoké seul : la coloration avance à la vitesse réglée."""
        total = self._stats_texte()["car"]
        self.pos_kara = min(float(total), self.pos_kara + self.vitesse_car * dt)
        self._rendu_karaoke(dt)

    def _karaoke_voix(self, dt):
        """Karaoké + voix : la coloration suit le mot effectivement prononcé."""
        total = self._stats_texte()["car"]
        p = self.voix.position()
        if p is not None and self.ligne_voix is not None:
            cible = min(float(total), float(self.base_voix + p))
            if cible > self.pos_kara:
                self.pos_kara = cible
        self._rendu_karaoke(dt)

    def _rendu_karaoke(self, dt):
        t = self.texte
        total = self._stats_texte()["car"]
        n = int(self.pos_kara)
        if n > self._kara_int:
            self._colorier(self._kara_int, n)
        self._kara_int = n
        idx = self._idx(min(n, max(total - 1, 0)))
        y = self._centre_ligne(idx)
        if y is None:
            t.see(idx)
        else:
            ecart = y - self.y_guide
            if abs(ecart) > 1:
                self._defiler(ecart * min(1.0, dt * 4.0))
        if n >= total:
            self._terminer()

    def _terminer(self):
        self.pause = True
        self.fin = True
        self._maj_etat(force=True)

    # ------------------------------------------------------------ ligne d'état
    def _stats_texte(self):
        if self._stats is None:
            contenu = self._texte()
            mots = max(1, len(re.findall(r"\w+(?:['’-]\w+)*", contenu)))
            paras, d = [], 0
            for l in contenu.split("\n"):
                if l.strip():
                    paras.append((d, d + len(l)))
                d += len(l) + 1
            pix = max(1, self._nb("1.0", "end", "ypixels"))
            lignes = max(1, self._nb("1.0", "end", "displaylines"))
            self._stats = {"mots": mots, "car": len(contenu), "pix": pix, "lignes": lignes, "paras": paras}
        s = self._stats
        if self.karaoke:
            car_par_mot = max(1.0, s["car"] / s["mots"])
            mpm = self.vitesse_car * 60.0 / car_par_mot
        else:
            mpm = self.vitesse_px * 60.0 * s["mots"] / s["pix"]
        s["mpm"] = int(round(mpm))
        return s

    def _nom_court_voix(self):
        """« Microsoft Hortense Desktop - French » → « Hortense », « ScanSoft Virginie_Dri40 » → « Virginie »."""
        nom = self.voix.nom or ""
        nom = nom.split(" (")[0].split(" - ")[0]
        for prefixe in ("Microsoft ", "ScanSoft ", "Nuance "):
            nom = nom.replace(prefixe, "")
        nom = nom.split("_")[0].split()
        return nom[0] if nom else self.t("son")

    def _maj_etat(self, force=False):
        self._prochain_etat = time.perf_counter() + 0.25
        if self.voix.message:
            msg = self.t(self.voix.message)
            self.voix.message = None
            self.barre.configure(text=msg)
            self._prochain_etat = time.perf_counter() + 3.0
            return
        s = self._stats_texte()
        pos = self._pos_lecture()
        ligne = self._nb("1.0", self._idx(pos), "displaylines") + 1
        avance = 0 if not s["car"] else round(100.0 * pos / s["car"])
        if self.fin:
            avance = 100
        if self.edition:
            mode = self.t("edition")
        elif self.fin:
            mode = self.t("fin")
        else:
            mode = self.t("pause") if self.pause else self.t("lecture")
        morceaux = [
            f"{s['mpm']} {self.t('mpm')}",
            f"{self.t('ligne')} {min(ligne, s['lignes'])}/{s['lignes']}",
            f"{min(100, max(0, avance))} %",
            mode,
        ]
        if self.karaoke:
            morceaux.append("★ " + self.t("karaoke"))
        if self.son:
            morceaux.append("♪ " + self._nom_court_voix())
        morceaux += [f"{self.taille} pt", self.t("code")]
        self.barre.configure(text=" │ ".join(morceaux) + " ║")
        # retour à la ligne (si l'écran est étroit) uniquement entre deux raccourcis
        elements = [e.replace(" ", "\u00a0") for e in self.t("raccourcis").split(" · ")]
        self.ligne_raccourcis.configure(text="\u00a0· ".join(elements))
        self._ajuster_raccourcis()

    # ------------------------------------------------------------------ aide
    def afficher_aide(self):
        if self._aide and self._aide.winfo_exists():
            self._aide.destroy()
            self._aide = None
            self.texte.focus_set()
            return
        w = tk.Toplevel(self.root)
        self._aide = w
        w.title(self.t("aide_titre"))
        w.configure(bg="#111111")
        w.transient(self.root)
        w.geometry("980x760")
        zone = tk.Text(w, wrap="word", bg="#111111", fg="#eeeeee", bd=0, highlightthickness=0,
                       font=("Arial", 16), padx=30, pady=24, spacing3=8, cursor="arrow")
        zone.pack(fill="both", expand=True)
        zone.tag_configure("titre", font=("Arial", 24, "bold"), foreground=self.JAUNE, spacing3=20)
        zone.tag_configure("cle", font=("Arial", 16, "bold"), foreground=self.JAUNE)
        zone.tag_configure("pied", foreground="#bbbbbb", spacing1=18)
        zone.insert("end", self.t("aide_titre") + "\n", "titre")
        for cle, desc in self.t("aide"):
            zone.insert("end", cle, "cle")
            zone.insert("end", "  —  " + desc + "\n")
        zone.insert("end", self.t("aide_pied"), "pied")
        if self.lang == "ar":
            zone.tag_configure("droite", justify="right")
            zone.tag_add("droite", "1.0", "end")
        zone.configure(state="disabled")

        def fermer(_e=None):
            w.destroy()
            self._aide = None
            self.texte.focus_set()
            return "break"
        for seq in ("<Escape>", "<F1>"):
            w.bind(seq, fermer)
        w.bind("<F4>", lambda e: (self.changer_langue(), "break")[1])
        w.protocol("WM_DELETE_WINDOW", fermer)
        w.focus_set()


def main():
    root = tk.Tk()
    chemin = sys.argv[1] if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]) else None
    Teleprompteur(root, chemin)
    root.mainloop()


if __name__ == "__main__":
    main()
