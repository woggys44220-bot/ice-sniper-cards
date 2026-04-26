from __future__ import annotations

import re
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from urllib.parse import quote_plus

import cv2
import pytesseract
from openpyxl import Workbook
from PIL import Image

PHOTO_DIR = Path("photos")
OUTPUT_FILE = Path("resultats_cartes_hockey.xlsx")

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}

KNOWN_BRANDS = {
    "upper deck": "Upper Deck",
    "o-pee-chee": "O-Pee-Chee",
    "opc": "O-Pee-Chee",
    "panini": "Panini",
    "fleer": "Fleer",
    "score": "Score",
    "parkhurst": "Parkhurst",
    "in the game": "In The Game",
    "sp authentic": "SP Authentic",
    "spx": "SPx",
    "mvp": "MVP",
    "ud": "Upper Deck",
}

KNOWN_SERIES = {
    "young guns": "Young Guns",
    "canvas": "Canvas",
    "platinum": "Platinum",
    "artifacts": "Artifacts",
    "credentials": "Credentials",
    "series one": "Series 1",
    "series 1": "Series 1",
    "series two": "Series 2",
    "series 2": "Series 2",
}

KNOWN_INSERTS = {
    "autograph": "Autograph",
    "auto": "Autograph",
    "jersey": "Jersey",
    "patch": "Patch",
    "parallel": "Parallel",
    "rookie": "Rookie",
}

KNOWN_TEAMS = {
    "canadiens": "Montreal Canadiens",
    "montreal": "Montreal Canadiens",
    "bruins": "Boston Bruins",
    "boston": "Boston Bruins",
    "maple leafs": "Toronto Maple Leafs",
    "leafs": "Toronto Maple Leafs",
    "toronto": "Toronto Maple Leafs",
    "rangers": "New York Rangers",
    "new york": "New York Rangers",
    "blackhawks": "Chicago Blackhawks",
    "chicago": "Chicago Blackhawks",
    "oilers": "Edmonton Oilers",
    "edmonton": "Edmonton Oilers",
    "canucks": "Vancouver Canucks",
    "vancouver": "Vancouver Canucks",
    "flames": "Calgary Flames",
    "calgary": "Calgary Flames",
    "avalanche": "Colorado Avalanche",
    "colorado": "Colorado Avalanche",
    "kings": "Los Angeles Kings",
    "los angeles": "Los Angeles Kings",
    "devils": "New Jersey Devils",
    "new jersey": "New Jersey Devils",
    "flyers": "Philadelphia Flyers",
    "philadelphia": "Philadelphia Flyers",
    "penguins": "Pittsburgh Penguins",
    "pittsburgh": "Pittsburgh Penguins",
    "sabres": "Buffalo Sabres",
    "buffalo": "Buffalo Sabres",
    "red wings": "Detroit Red Wings",
    "detroit": "Detroit Red Wings",
    "senators": "Ottawa Senators",
    "ottawa": "Ottawa Senators",
    "jets": "Winnipeg Jets",
    "winnipeg": "Winnipeg Jets",
    "wild": "Minnesota Wild",
    "minnesota": "Minnesota Wild",
}


@dataclass
class CardResult:
    fichier_photo: str
    joueur: str
    annee: str
    marque: str
    serie: str
    insert: str
    numero_carte: str
    equipe: str
    texte_ocr: str
    requete_prix: str
    lien_ebay: str
    lien_130point: str
    prix_bas: str
    prix_moyen: str
    prix_haut: str
    statut: str


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def run_ocr(path: Path) -> str:
    """Lit le texte réel de l'image via PIL + OpenCV + pytesseract."""
    texts: list[str] = []

    with Image.open(path) as img:
        pil_text = pytesseract.image_to_string(img, lang="eng", config="--psm 6")
        texts.append(pil_text)

    cv_img = cv2.imread(str(path))
    if cv_img is None:
        return normalize_text(" ".join(texts))

    gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
    upscaled = cv2.resize(gray, None, fx=1.8, fy=1.8, interpolation=cv2.INTER_CUBIC)
    denoised = cv2.GaussianBlur(upscaled, (3, 3), 0)

    _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    ocr_binary = pytesseract.image_to_string(binary, lang="eng", config="--psm 6")
    texts.append(ocr_binary)

    adaptive = cv2.adaptiveThreshold(
        denoised,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31,
        5,
    )
    ocr_adaptive = pytesseract.image_to_string(adaptive, lang="eng", config="--psm 11")
    texts.append(ocr_adaptive)

    return normalize_text(" ".join(texts))


def extract_year(ocr_text: str) -> str:
    match = re.search(r"\b(19\d{2}|20[0-2]\d)\b", ocr_text)
    return match.group(1) if match else "à vérifier"


def extract_card_number(ocr_text: str) -> str:
    patterns = [
        r"(?:no\.?|#|num(?:ber)?\s*)([a-z]{0,3}\d{1,4})\b",
        r"\b([A-Z]{1,3}-?\d{1,4})\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, ocr_text, flags=re.IGNORECASE)
        if match:
            return match.group(1).upper().replace("-", "")
    return "à vérifier"


def match_from_dict(ocr_text: str, mapping: dict[str, str]) -> str:
    lowered = ocr_text.lower()
    for key, value in mapping.items():
        if key in lowered:
            return value
    return "à vérifier"


def extract_player(ocr_text: str) -> str:
    lines = [ln.strip() for ln in re.split(r"[\n\r]+", ocr_text) if ln.strip()]
    for line in lines:
        if len(line.split()) == 2 and all(part.isalpha() for part in line.split()):
            return line.title()

    tokens = re.findall(r"[A-Za-z]{3,}", ocr_text)
    for i in range(len(tokens) - 1):
        first, last = tokens[i], tokens[i + 1]
        if first.lower() in {"hockey", "card", "national", "league", "upper", "deck"}:
            continue
        return f"{first.title()} {last.title()}"

    return "à vérifier"


def build_price_query(
    joueur: str,
    annee: str,
    marque: str,
    serie: str,
    insert: str,
    numero_carte: str,
    equipe: str,
) -> str:
    elements = [
        joueur,
        annee,
        marque,
        serie,
        insert,
        f"#{numero_carte}" if numero_carte != "à vérifier" else "",
        equipe,
        "hockey card",
    ]
    clean = [e.strip() for e in elements if e and e != "à vérifier"]
    return " ".join(clean) if clean else "hockey card à vérifier"


def build_links(query: str) -> tuple[str, str]:
    encoded = quote_plus(query)
    ebay = f"https://www.ebay.com/sch/i.html?_nkw={encoded}&LH_Sold=1&LH_Complete=1"
    point130 = f"https://130point.com/sales/?q={encoded}"
    return ebay, point130


def analyze_image(path: Path) -> CardResult:
    try:
        texte_ocr = run_ocr(path)
    except pytesseract.TesseractNotFoundError:
        texte_ocr = "Tesseract non trouvé. Installez Tesseract OCR et ajoutez-le au PATH Windows."
    except Exception as exc:  # fallback simple pour éviter un crash global
        texte_ocr = f"OCR erreur: {exc}"

    joueur = extract_player(texte_ocr)
    annee = extract_year(texte_ocr)
    marque = match_from_dict(texte_ocr, KNOWN_BRANDS)
    serie = match_from_dict(texte_ocr, KNOWN_SERIES)
    insert = match_from_dict(texte_ocr, KNOWN_INSERTS)
    numero_carte = extract_card_number(texte_ocr)
    equipe = match_from_dict(texte_ocr, KNOWN_TEAMS)

    requete = build_price_query(joueur, annee, marque, serie, insert, numero_carte, equipe)
    lien_ebay, lien_130point = build_links(requete)

    champs = [joueur, annee, marque, serie, insert, numero_carte, equipe]
    unresolved = sum(value == "à vérifier" for value in champs)
    statut = "ok" if unresolved <= 2 else "OCR partiel (champs à vérifier)"

    return CardResult(
        fichier_photo=path.name,
        joueur=joueur,
        annee=annee,
        marque=marque,
        serie=serie,
        insert=insert,
        numero_carte=numero_carte,
        equipe=equipe,
        texte_ocr=texte_ocr if texte_ocr else "à vérifier",
        requete_prix=requete,
        lien_ebay=lien_ebay,
        lien_130point=lien_130point,
        prix_bas="à vérifier",
        prix_moyen="à vérifier",
        prix_haut="à vérifier",
        statut=statut,
    )


def iter_images(folder: Path) -> Iterable[Path]:
    for file in sorted(folder.iterdir()):
        if file.is_file() and file.suffix.lower() in IMAGE_EXTENSIONS:
            yield file


def export_to_excel(results: list[CardResult], output_file: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "cartes_hockey"

    headers = list(CardResult.__annotations__.keys())
    ws.append(headers)

    for result in results:
        ws.append(list(asdict(result).values()))

    wb.save(output_file)


def main() -> None:
    PHOTO_DIR.mkdir(exist_ok=True)

    images = list(iter_images(PHOTO_DIR))
    results = [analyze_image(image) for image in images]

    if not results:
        results.append(
            CardResult(
                fichier_photo="(aucune image détectée)",
                joueur="à vérifier",
                annee="à vérifier",
                marque="à vérifier",
                serie="à vérifier",
                insert="à vérifier",
                numero_carte="à vérifier",
                equipe="à vérifier",
                texte_ocr="à vérifier",
                requete_prix="hockey card à vérifier",
                lien_ebay="https://www.ebay.com",
                lien_130point="https://130point.com",
                prix_bas="à vérifier",
                prix_moyen="à vérifier",
                prix_haut="à vérifier",
                statut="ajoutez des photos dans le dossier photos/",
            )
        )

    export_to_excel(results, OUTPUT_FILE)
    print(f"Analyse terminée. Fichier généré: {OUTPUT_FILE.resolve()}")


if __name__ == "__main__":
    main()
