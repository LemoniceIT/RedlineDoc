"""
Internationalization (i18n) module for RedlineDoc.

Provides a lightweight, dependency-free translation system.
All user-facing strings are collected here and translated into five languages:
    en  –  English  (default)
    es  –  Español
    fr  –  Français
    it  –  Italiano
    de  –  Deutsch

Usage:
    from i18n import t, set_language, get_language, LANGUAGES
    label = t("btn.compare")           # returns translated string
    set_language("it")                 # switches to Italian
"""

from __future__ import annotations

# ---------------------------------------------------------------------------
# Language registry
# ---------------------------------------------------------------------------

# Ordered mapping: code → display name (used by the GUI selector)
LANGUAGES: dict[str, str] = {
    "en": "English",
    "es": "Español",
    "fr": "Français",
    "it": "Italiano",
    "de": "Deutsch",
}

_current_language: str = "en"


def set_language(code: str) -> None:
    """Set the active language. *code* must be one of the LANGUAGES keys."""
    global _current_language
    if code not in LANGUAGES:
        raise ValueError(f"Unsupported language code: {code!r}")
    _current_language = code


def get_language() -> str:
    """Return the current language code."""
    return _current_language


def t(key: str) -> str:
    """Return the translated string for *key* in the active language.

    Falls back to English if the key is missing for the current language,
    then to the raw key itself as a last resort.
    """
    lang_table = _TRANSLATIONS.get(_current_language, _TRANSLATIONS["en"])
    text = lang_table.get(key)
    if text is not None:
        return text
    # Fallback to English
    text = _TRANSLATIONS["en"].get(key)
    return text if text is not None else key


# ---------------------------------------------------------------------------
# Translation tables
# ---------------------------------------------------------------------------

_TRANSLATIONS: dict[str, dict[str, str]] = {
    # ==================================================================
    # ENGLISH  (default)
    # ==================================================================
    "en": {
        # --- App ---
        "app.startup": "Application started: {name} v{version}",
        "app.ready": "Ready for document loading.",
        "app.closing": "Closing application…",

        # --- File selection ---
        "file.no_selection": "No file selected",
        "file.dialog_original": "Select the Original Document",
        "file.dialog_modified": "Select the Modified Document",
        "file.loaded_original": "Original document loaded: {path}",
        "file.loaded_modified": "Modified document loaded: {path}",

        # --- Supported file-type labels (file dialog filters) ---
        "filetype.doc": "Word 97-2003 Documents (.doc)",
        "filetype.docx": "Word Documents (.docx)",
        "filetype.pdf": "PDF Documents (.pdf)",

        # --- Buttons ---
        "btn.original_doc": "\U0001f4c4 Original Document",
        "btn.modified_doc": "\U0001f4c4 Modified Document",
        "btn.config": "\u2699\ufe0f Configuration",
        "btn.compare": "\U0001f50d Compare",
        "btn.credits": "\u2139 Credits",
        "btn.clear": "Clear",
        "btn.cancel": "Cancel",
        "btn.apply": "Apply",
        "btn.close": "Close",

        # --- Config dialog ---
        "config.title": "Comparison Configuration",
        "config.header": "\u2699\ufe0f Comparison Options Configuration",
        "config.description": (
            "Select which elements to ignore during\n"
            "text extraction for comparison."
        ),
        "config.ignore_formatting": "Ignore formatting",
        "config.ignore_images": "Ignore images",
        "config.ignore_header": "Ignore header",
        "config.ignore_footer": "Ignore footer",

        # --- Credits dialog ---
        "credits.title": "\u2139 Credits & Information",
        "credits.developed_by": "Developed by {author}",
        "credits.description": (
            "Comparison and analysis of Word (.doc/.docx) and PDF documents"
        ),
        "credits.license_info": (
            "License: {license}\n"
            "Attribution required — No warranty provided"
        ),
        "credits.tech_header": "Libraries and Technologies Used:\n",
        "credits.tech_customtkinter": "CustomTkinter (MIT License) — cyber/dark GUI",
        "credits.tech_python_docx": "python-docx (MIT License) — .docx text extraction",
        "credits.tech_pywin32": "pywin32 (Microsoft Public License) — .doc text extraction",
        "credits.tech_pdfplumber": "pdfplumber (MIT License) — .pdf text extraction",
        "credits.tech_lxml": "lxml (BSD License) — python-docx dependency",
        "credits.tech_difflib": "difflib / tempfile / webbrowser (Python stdlib, PSF License)",
        "credits.tech_tkinter": "tkinter.filedialog (Python stdlib, PSF License)",

        # --- Terminal ---
        "terminal.header": "  > TERMINAL LOG [READ-ONLY]",

        # --- Status bar ---
        "status.ready": "READY",
        "status.processing": "PROCESSING…",

        # --- Comparison workflow ---
        "compare.in_progress": "\u23f3 Comparison in progress…",
        "compare.both_required": (
            "\u274c You must load both documents before comparing."
        ),
        "compare.ext_mismatch": (
            "\u274c Error: both documents must have the same extension "
            "({ext1} != {ext2}). Operation cancelled."
        ),
        "compare.options": "Options: header={header}, footer={footer}",
        "compare.yes": "yes",
        "compare.no": "no",
        "compare.start_original": "Starting text extraction of Original Document…",
        "compare.done_chars": "Extraction completed. Characters extracted: {count}",
        "compare.start_modified": "Starting text extraction of Modified Document…",
        "compare.generating": "Generating comparison report with difflib.HtmlDiff…",
        "compare.report_generated": "Report generated: {path}",
        "compare.opening_browser": "Opening report in default browser…",
        "compare.completed": (
            "\u2705 Comparison completed. The report has been opened in the browser."
        ),
        "compare.error": "\u274c Error during comparison: {error}",

        # --- Document processor ---
        "doc.original": "Original Document",
        "doc.modified": "Modified Document",
        "doc.unsupported_format": "Unsupported format: {ext}",
        "doc.header_tag": "[header]",
        "doc.footer_tag": "[footer]",
        "doc.module_info": (
            "document_processor.py is a library module.\n"
            "Start the application with:  python gui.py"
        ),

        # --- Extraction library descriptions ---
        "lib.pywin32": "Text extraction from .doc documents via Windows COM",
        "lib.python_docx": "Text extraction from .docx documents",
        "lib.pdfplumber": "Text extraction from .pdf documents",
        "lib.difflib": "Text comparison and HTML report generation",

        # --- Language selector ---
        "lang.label": "\U0001f310 Language:",
    },

    # ==================================================================
    # ESPAÑOL
    # ==================================================================
    "es": {
        "app.startup": "Aplicación iniciada: {name} v{version}",
        "app.ready": "Lista para la carga de documentos.",
        "app.closing": "Cerrando aplicación…",

        "file.no_selection": "Ningún archivo seleccionado",
        "file.dialog_original": "Seleccionar el Documento Original",
        "file.dialog_modified": "Seleccionar el Documento Modificado",
        "file.loaded_original": "Documento original cargado: {path}",
        "file.loaded_modified": "Documento modificado cargado: {path}",

        "filetype.doc": "Documentos Word 97-2003 (.doc)",
        "filetype.docx": "Documentos Word (.docx)",
        "filetype.pdf": "Documentos PDF (.pdf)",

        "btn.original_doc": "\U0001f4c4 Documento Original",
        "btn.modified_doc": "\U0001f4c4 Documento Modificado",
        "btn.config": "\u2699\ufe0f Configuración",
        "btn.compare": "\U0001f50d Comparar",
        "btn.credits": "\u2139 Créditos",
        "btn.clear": "Limpiar",
        "btn.cancel": "Cancelar",
        "btn.apply": "Aplicar",
        "btn.close": "Cerrar",

        "config.title": "Configuración de Comparación",
        "config.header": "\u2699\ufe0f Configuración de Opciones de Comparación",
        "config.description": (
            "Seleccione qué elementos ignorar durante\n"
            "la extracción de texto para la comparación."
        ),
        "config.ignore_formatting": "Ignorar formato",
        "config.ignore_images": "Ignorar imágenes",
        "config.ignore_header": "Ignorar encabezado",
        "config.ignore_footer": "Ignorar pie de página",

        "credits.title": "\u2139 Créditos e Información",
        "credits.developed_by": "Desarrollado por {author}",
        "credits.description": (
            "Comparación y análisis de documentos Word (.doc/.docx) y PDF"
        ),
        "credits.license_info": (
            "Licencia: {license}\n"
            "Se requiere atribución — Sin garantía"
        ),
        "credits.tech_header": "Bibliotecas y Tecnologías Utilizadas:\n",
        "credits.tech_customtkinter": "CustomTkinter (Licencia MIT) — interfaz gráfica cyber/dark",
        "credits.tech_python_docx": "python-docx (Licencia MIT) — extracción de texto .docx",
        "credits.tech_pywin32": "pywin32 (Licencia Pública de Microsoft) — extracción de texto .doc",
        "credits.tech_pdfplumber": "pdfplumber (Licencia MIT) — extracción de texto .pdf",
        "credits.tech_lxml": "lxml (Licencia BSD) — dependencia de python-docx",
        "credits.tech_difflib": "difflib / tempfile / webbrowser (stdlib Python, Licencia PSF)",
        "credits.tech_tkinter": "tkinter.filedialog (stdlib Python, Licencia PSF)",

        "terminal.header": "  > REGISTRO TERMINAL [SOLO LECTURA]",

        "status.ready": "LISTO",
        "status.processing": "PROCESANDO…",

        "compare.in_progress": "\u23f3 Comparación en curso…",
        "compare.both_required": (
            "\u274c Debe cargar ambos documentos antes de comparar."
        ),
        "compare.ext_mismatch": (
            "\u274c Error: ambos documentos deben tener la misma extensión "
            "({ext1} != {ext2}). Operación cancelada."
        ),
        "compare.options": "Opciones: encabezado={header}, pie de página={footer}",
        "compare.yes": "sí",
        "compare.no": "no",
        "compare.start_original": "Iniciando extracción de texto del Documento Original…",
        "compare.done_chars": "Extracción completada. Caracteres extraídos: {count}",
        "compare.start_modified": "Iniciando extracción de texto del Documento Modificado…",
        "compare.generating": "Generando informe de comparación con difflib.HtmlDiff…",
        "compare.report_generated": "Informe generado: {path}",
        "compare.opening_browser": "Abriendo informe en el navegador predeterminado…",
        "compare.completed": (
            "\u2705 Comparación completada. El informe se ha abierto en el navegador."
        ),
        "compare.error": "\u274c Error durante la comparación: {error}",

        "doc.original": "Documento Original",
        "doc.modified": "Documento Modificado",
        "doc.unsupported_format": "Formato no soportado: {ext}",
        "doc.header_tag": "[encabezado]",
        "doc.footer_tag": "[pie de página]",
        "doc.module_info": (
            "document_processor.py es un módulo de biblioteca.\n"
            "Inicie la aplicación con:  python gui.py"
        ),

        "lib.pywin32": "Extracción de texto de documentos .doc mediante Windows COM",
        "lib.python_docx": "Extracción de texto de documentos .docx",
        "lib.pdfplumber": "Extracción de texto de documentos .pdf",
        "lib.difflib": "Comparación textual y generación de informes HTML",

        "lang.label": "\U0001f310 Idioma:",
    },

    # ==================================================================
    # FRANÇAIS
    # ==================================================================
    "fr": {
        "app.startup": "Application démarrée : {name} v{version}",
        "app.ready": "Prête pour le chargement des documents.",
        "app.closing": "Fermeture de l'application…",

        "file.no_selection": "Aucun fichier sélectionné",
        "file.dialog_original": "Sélectionner le Document Original",
        "file.dialog_modified": "Sélectionner le Document Modifié",
        "file.loaded_original": "Document original chargé : {path}",
        "file.loaded_modified": "Document modifié chargé : {path}",

        "filetype.doc": "Documents Word 97-2003 (.doc)",
        "filetype.docx": "Documents Word (.docx)",
        "filetype.pdf": "Documents PDF (.pdf)",

        "btn.original_doc": "\U0001f4c4 Document Original",
        "btn.modified_doc": "\U0001f4c4 Document Modifié",
        "btn.config": "\u2699\ufe0f Configuration",
        "btn.compare": "\U0001f50d Comparer",
        "btn.credits": "\u2139 Crédits",
        "btn.clear": "Effacer",
        "btn.cancel": "Annuler",
        "btn.apply": "Appliquer",
        "btn.close": "Fermer",

        "config.title": "Configuration de la Comparaison",
        "config.header": "\u2699\ufe0f Configuration des Options de Comparaison",
        "config.description": (
            "Sélectionnez les éléments à ignorer lors\n"
            "de l'extraction du texte pour la comparaison."
        ),
        "config.ignore_formatting": "Ignorer la mise en forme",
        "config.ignore_images": "Ignorer les images",
        "config.ignore_header": "Ignorer l'en-tête",
        "config.ignore_footer": "Ignorer le pied de page",

        "credits.title": "\u2139 Crédits et Informations",
        "credits.developed_by": "Développé par {author}",
        "credits.description": (
            "Comparaison et analyse de documents Word (.doc/.docx) et PDF"
        ),
        "credits.license_info": (
            "Licence : {license}\n"
            "Attribution obligatoire — Aucune garantie"
        ),
        "credits.tech_header": "Bibliothèques et Technologies Utilisées :\n",
        "credits.tech_customtkinter": "CustomTkinter (Licence MIT) — interface graphique cyber/dark",
        "credits.tech_python_docx": "python-docx (Licence MIT) — extraction de texte .docx",
        "credits.tech_pywin32": "pywin32 (Licence Publique Microsoft) — extraction de texte .doc",
        "credits.tech_pdfplumber": "pdfplumber (Licence MIT) — extraction de texte .pdf",
        "credits.tech_lxml": "lxml (Licence BSD) — dépendance de python-docx",
        "credits.tech_difflib": "difflib / tempfile / webbrowser (stdlib Python, Licence PSF)",
        "credits.tech_tkinter": "tkinter.filedialog (stdlib Python, Licence PSF)",

        "terminal.header": "  > JOURNAL TERMINAL [LECTURE SEULE]",

        "status.ready": "PRÊT",
        "status.processing": "TRAITEMENT…",

        "compare.in_progress": "\u23f3 Comparaison en cours…",
        "compare.both_required": (
            "\u274c Vous devez charger les deux documents avant de comparer."
        ),
        "compare.ext_mismatch": (
            "\u274c Erreur : les deux documents doivent avoir la même extension "
            "({ext1} != {ext2}). Opération annulée."
        ),
        "compare.options": "Options : en-tête={header}, pied de page={footer}",
        "compare.yes": "oui",
        "compare.no": "non",
        "compare.start_original": "Début de l'extraction du texte du Document Original…",
        "compare.done_chars": "Extraction terminée. Caractères extraits : {count}",
        "compare.start_modified": "Début de l'extraction du texte du Document Modifié…",
        "compare.generating": "Génération du rapport de comparaison avec difflib.HtmlDiff…",
        "compare.report_generated": "Rapport généré : {path}",
        "compare.opening_browser": "Ouverture du rapport dans le navigateur par défaut…",
        "compare.completed": (
            "\u2705 Comparaison terminée. Le rapport a été ouvert dans le navigateur."
        ),
        "compare.error": "\u274c Erreur lors de la comparaison : {error}",

        "doc.original": "Document Original",
        "doc.modified": "Document Modifié",
        "doc.unsupported_format": "Format non pris en charge : {ext}",
        "doc.header_tag": "[en-tête]",
        "doc.footer_tag": "[pied de page]",
        "doc.module_info": (
            "document_processor.py est un module de bibliothèque.\n"
            "Lancez l'application avec :  python gui.py"
        ),

        "lib.pywin32": "Extraction de texte de documents .doc via Windows COM",
        "lib.python_docx": "Extraction de texte de documents .docx",
        "lib.pdfplumber": "Extraction de texte de documents .pdf",
        "lib.difflib": "Comparaison textuelle et génération de rapports HTML",

        "lang.label": "\U0001f310 Langue :",
    },

    # ==================================================================
    # ITALIANO
    # ==================================================================
    "it": {
        "app.startup": "Avvio applicazione: {name} v{version}",
        "app.ready": "Pronta per il caricamento dei documenti.",
        "app.closing": "Chiusura applicazione…",

        "file.no_selection": "Nessun file selezionato",
        "file.dialog_original": "Seleziona il Documento Originale",
        "file.dialog_modified": "Seleziona il Documento Modificato",
        "file.loaded_original": "Documento originale caricato: {path}",
        "file.loaded_modified": "Documento modificato caricato: {path}",

        "filetype.doc": "Documenti Word 97-2003 (.doc)",
        "filetype.docx": "Documenti Word (.docx)",
        "filetype.pdf": "Documenti PDF (.pdf)",

        "btn.original_doc": "\U0001f4c4 Documento Originale",
        "btn.modified_doc": "\U0001f4c4 Documento Modificato",
        "btn.config": "\u2699\ufe0f Configurazione",
        "btn.compare": "\U0001f50d Confronta",
        "btn.credits": "\u2139 Crediti",
        "btn.clear": "Pulisci",
        "btn.cancel": "Annulla",
        "btn.apply": "Applica",
        "btn.close": "Chiudi",

        "config.title": "Configurazione Confronto",
        "config.header": "\u2699\ufe0f Configurazione Opzioni di Confronto",
        "config.description": (
            "Seleziona quali elementi ignorare durante\n"
            "l'estrazione del testo confrontato."
        ),
        "config.ignore_formatting": "Ignora formattazione",
        "config.ignore_images": "Ignora immagini",
        "config.ignore_header": "Ignora intestazione",
        "config.ignore_footer": "Ignora piè di pagina",

        "credits.title": "\u2139 Crediti & Informazioni",
        "credits.developed_by": "Sviluppato da {author}",
        "credits.description": (
            "Confronto e analisi di documenti Word (.doc/.docx) e PDF"
        ),
        "credits.license_info": (
            "Licenza: {license}\n"
            "Citazione obbligatoria — Nessuna responsabilità"
        ),
        "credits.tech_header": "Librerie e Tecnologie Utilizzate:\n",
        "credits.tech_customtkinter": "CustomTkinter (Licenza MIT) — interfaccia grafica cyber/dark",
        "credits.tech_python_docx": "python-docx (Licenza MIT) — estrazione testo .docx",
        "credits.tech_pywin32": "pywin32 (Licenza Pubblica Microsoft) — estrazione testo .doc",
        "credits.tech_pdfplumber": "pdfplumber (Licenza MIT) — estrazione testo .pdf",
        "credits.tech_lxml": "lxml (Licenza BSD) — dipendenza di python-docx",
        "credits.tech_difflib": "difflib / tempfile / webbrowser (stdlib Python, Licenza PSF)",
        "credits.tech_tkinter": "tkinter.filedialog (stdlib Python, Licenza PSF)",

        "terminal.header": "  > TERMINALE LOG [SOLA LETTURA]",

        "status.ready": "PRONTO",
        "status.processing": "ELABORAZIONE…",

        "compare.in_progress": "\u23f3 Confronto in corso…",
        "compare.both_required": (
            "\u274c Devi caricare entrambi i documenti prima di confrontare."
        ),
        "compare.ext_mismatch": (
            "\u274c Errore: i due documenti devono avere la stessa estensione "
            "({ext1} != {ext2}). Operazione annullata."
        ),
        "compare.options": "Opzioni: intestazione={header}, piè di pagina={footer}",
        "compare.yes": "sì",
        "compare.no": "no",
        "compare.start_original": "Avvio estrazione testo del Documento Originale…",
        "compare.done_chars": "Estrazione completata. Caratteri estratti: {count}",
        "compare.start_modified": "Avvio estrazione testo del Documento Modificato…",
        "compare.generating": "Generazione report di confronto con difflib.HtmlDiff…",
        "compare.report_generated": "Report generato: {path}",
        "compare.opening_browser": "Apertura del report nel browser predefinito…",
        "compare.completed": (
            "\u2705 Confronto completato. Il report è stato aperto nel browser."
        ),
        "compare.error": "\u274c Errore durante il confronto: {error}",

        "doc.original": "Documento Originale",
        "doc.modified": "Documento Modificato",
        "doc.unsupported_format": "Formato non supportato: {ext}",
        "doc.header_tag": "[intestazione]",
        "doc.footer_tag": "[piè di pagina]",
        "doc.module_info": (
            "document_processor.py è un modulo libreria.\n"
            "Avvia l'applicazione con:  python gui.py"
        ),

        "lib.pywin32": "Estrazione testo da documenti .doc tramite Windows COM",
        "lib.python_docx": "Estrazione testo da documenti .docx",
        "lib.pdfplumber": "Estrazione testo da documenti .pdf",
        "lib.difflib": "Confronto testuale e generazione report HTML",

        "lang.label": "\U0001f310 Lingua:",
    },

    # ==================================================================
    # DEUTSCH
    # ==================================================================
    "de": {
        "app.startup": "Anwendung gestartet: {name} v{version}",
        "app.ready": "Bereit zum Laden der Dokumente.",
        "app.closing": "Anwendung wird geschlossen…",

        "file.no_selection": "Keine Datei ausgewählt",
        "file.dialog_original": "Originaldokument auswählen",
        "file.dialog_modified": "Geändertes Dokument auswählen",
        "file.loaded_original": "Originaldokument geladen: {path}",
        "file.loaded_modified": "Geändertes Dokument geladen: {path}",

        "filetype.doc": "Word 97-2003 Dokumente (.doc)",
        "filetype.docx": "Word-Dokumente (.docx)",
        "filetype.pdf": "PDF-Dokumente (.pdf)",

        "btn.original_doc": "\U0001f4c4 Originaldokument",
        "btn.modified_doc": "\U0001f4c4 Geändertes Dokument",
        "btn.config": "\u2699\ufe0f Konfiguration",
        "btn.compare": "\U0001f50d Vergleichen",
        "btn.credits": "\u2139 Credits",
        "btn.clear": "Löschen",
        "btn.cancel": "Abbrechen",
        "btn.apply": "Übernehmen",
        "btn.close": "Schließen",

        "config.title": "Vergleichskonfiguration",
        "config.header": "\u2699\ufe0f Konfiguration der Vergleichsoptionen",
        "config.description": (
            "Wählen Sie aus, welche Elemente bei der\n"
            "Textextraktion für den Vergleich ignoriert werden sollen."
        ),
        "config.ignore_formatting": "Formatierung ignorieren",
        "config.ignore_images": "Bilder ignorieren",
        "config.ignore_header": "Kopfzeile ignorieren",
        "config.ignore_footer": "Fußzeile ignorieren",

        "credits.title": "\u2139 Credits & Informationen",
        "credits.developed_by": "Entwickelt von {author}",
        "credits.description": (
            "Vergleich und Analyse von Word-Dokumenten (.doc/.docx) und PDF"
        ),
        "credits.license_info": (
            "Lizenz: {license}\n"
            "Namensnennung erforderlich — Keine Gewährleistung"
        ),
        "credits.tech_header": "Verwendete Bibliotheken und Technologien:\n",
        "credits.tech_customtkinter": "CustomTkinter (MIT-Lizenz) — Cyber/Dark-GUI",
        "credits.tech_python_docx": "python-docx (MIT-Lizenz) — .docx-Textextraktion",
        "credits.tech_pywin32": "pywin32 (Microsoft Public License) — .doc-Textextraktion",
        "credits.tech_pdfplumber": "pdfplumber (MIT-Lizenz) — .pdf-Textextraktion",
        "credits.tech_lxml": "lxml (BSD-Lizenz) — Abhängigkeit von python-docx",
        "credits.tech_difflib": "difflib / tempfile / webbrowser (Python-Stdlib, PSF-Lizenz)",
        "credits.tech_tkinter": "tkinter.filedialog (Python-Stdlib, PSF-Lizenz)",

        "terminal.header": "  > TERMINAL-LOG [SCHREIBGESCHÜTZT]",

        "status.ready": "BEREIT",
        "status.processing": "VERARBEITUNG…",

        "compare.in_progress": "\u23f3 Vergleich läuft…",
        "compare.both_required": (
            "\u274c Sie müssen beide Dokumente laden, bevor Sie vergleichen."
        ),
        "compare.ext_mismatch": (
            "\u274c Fehler: Beide Dokumente müssen die gleiche Erweiterung haben "
            "({ext1} != {ext2}). Vorgang abgebrochen."
        ),
        "compare.options": "Optionen: Kopfzeile={header}, Fußzeile={footer}",
        "compare.yes": "ja",
        "compare.no": "nein",
        "compare.start_original": "Starte Textextraktion des Originaldokuments…",
        "compare.done_chars": "Extraktion abgeschlossen. Extrahierte Zeichen: {count}",
        "compare.start_modified": "Starte Textextraktion des geänderten Dokuments…",
        "compare.generating": "Vergleichsbericht wird mit difflib.HtmlDiff erstellt…",
        "compare.report_generated": "Bericht erstellt: {path}",
        "compare.opening_browser": "Bericht wird im Standardbrowser geöffnet…",
        "compare.completed": (
            "\u2705 Vergleich abgeschlossen. Der Bericht wurde im Browser geöffnet."
        ),
        "compare.error": "\u274c Fehler beim Vergleich: {error}",

        "doc.original": "Originaldokument",
        "doc.modified": "Geändertes Dokument",
        "doc.unsupported_format": "Nicht unterstütztes Format: {ext}",
        "doc.header_tag": "[Kopfzeile]",
        "doc.footer_tag": "[Fußzeile]",
        "doc.module_info": (
            "document_processor.py ist ein Bibliotheksmodul.\n"
            "Starten Sie die Anwendung mit:  python gui.py"
        ),

        "lib.pywin32": "Textextraktion aus .doc-Dokumenten über Windows COM",
        "lib.python_docx": "Textextraktion aus .docx-Dokumenten",
        "lib.pdfplumber": "Textextraktion aus .pdf-Dokumenten",
        "lib.difflib": "Textvergleich und HTML-Berichterstellung",

        "lang.label": "\U0001f310 Sprache:",
    },
}
