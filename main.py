from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Iterable
from urllib.parse import quote_plus

from openpyxl import Workbook


PHOTO_DIR = Path("photos")
OUTPUT_FILE = Path("resultats_cartes_hockey.xlsx")

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}

KNOWN_BRANDS = {
    "upperdeck": "Upper Deck",
    "upper_deck": "Upper Deck",
    "opc": "O-Pee-Chee",
    "opeechee": "O-Pee-Chee",
    "panini": "Panini",
    "fleer": "Fleer",
    "score": "Score",
    "parkhurst": "Parkhurst",
    "itg": "In The Game",
    "sp": "SP",
    "mvp": "MVP",
}

KNOWN_TEAMS = {
    "canadiens": "Montreal Canadiens",
    "habs": "Montreal Canadiens",
    "bruins": "Boston Bruins",
    "mapleleafs": "Toronto Maple Leafs",
    "leafs": "Toronto Maple Leafs",
    "rangers": "New York Rangers",
    "blackhawks": "Chicago Blackhawks",
    "oilers": "Edmonton Oilers",
    "canucks": "Vancouver Canucks",
    "flames": "Calgary Flames",
    "avalanche": "Colorado Avalanche",
    "kings": "Los Angeles Kings",
    "devils": "New Jersey Devils",
    "flyers": "Philadelphia Flyers",
    "penguins": "Pittsburgh Penguins",
    "sabres": "Buffalo Sabres",
    "redwings": "Detroit Red Wings",
    "senators": "Ottawa Senators",
    "jets": "Winnipeg Jets",
    "wild": "Minnesota Wild",
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
    requete_prix: str
    lien_ebay: str
    lien_130point: str
    prix_bas: str
    prix_moyen: str
    prix_haut: str
    statut: str


def normalize_token(token: str) -> str:
    return re.sub(r"[^a-z0-9]", "", token.lower())


def tokenize_stem(stem: str) -> list[str]:
    parts = re.split(r"[_\-\s]+", stem)
    return [p for p in parts if p]


def find_year(tokens: list[str]) -> str:
    for token in tokens:
        if re.fullmatch(r"(19|20)\d{2}", token):
            return token
        if re.fullmatch(r"(19|20)\d{2}(19|20)\d{2}", token):
            return f"{token[:4]}-{token[4:]}"
    return "à vérifier"


def find_card_number(tokens: list[str]) -> str:
    for token in tokens:
        clean = token.replace("#", "")
        if re.fullmatch(r"[A-Za-z]{0,3}\d{1,4}", clean):
            if any(ch.isdigit() for ch in clean):
                return clean.upper()
    return "à vérifier"


def find_brand(tokens: list[str]) -> str:
    for token in tokens:
        key = normalize_token(token)
        if key in KNOWN_BRANDS:
            return KNOWN_BRANDS[key]
    return "à vérifier"


def find_team(tokens: list[str]) -> str:
    for token in tokens:
        key = normalize_token(token)
        if key in KNOWN_TEAMS:
            return KNOWN_TEAMS[key]
    return "à vérifier"


def guess_player(tokens: list[str]) -> str:
    if not tokens:
        return "à vérifier"

    stopwords = {
        "front",
        "back",
        "recto",
        "verso",
        "scan",
        "card",
        "hockey",
    }

    player_parts: list[str] = []
    for token in tokens:
        clean = normalize_token(token)
        if not clean or clean in stopwords:
            continue
        if re.fullmatch(r"(19|20)\d{2}", clean):
            break
        if clean in KNOWN_BRANDS or clean in KNOWN_TEAMS:
            break
        if re.fullmatch(r"[a-z]{0,3}\d{1,4}", clean):
            break
        player_parts.append(token.title())
        if len(player_parts) >= 3:
            break

    return " ".join(player_parts) if player_parts else "à vérifier"


def guess_series(tokens: list[str], marque: str) -> str:
    if marque == "à vérifier":
        return "à vérifier"

    cleaned = [normalize_token(t) for t in tokens]
    if "youngguns" in cleaned:
        return "Young Guns"
    if "rookie" in cleaned:
        return "Rookie"
    if "canvas" in cleaned:
        return "Canvas"
    return "à vérifier"


def guess_insert(tokens: list[str]) -> str:
    cleaned = {normalize_token(t) for t in tokens}
    insert_map = {
        "autograph": "Autograph",
        "auto": "Autograph",
        "jersey": "Jersey",
        "patch": "Patch",
        "parallel": "Parallel",
        "rookie": "Rookie",
    }
    for key, value in insert_map.items():
        if key in cleaned:
            return value
    return "à vérifier"


def build_price_query(
    joueur: str,
    annee: str,
    marque: str,
    serie: str,
    numero_carte: str,
    equipe: str,
) -> str:
    elements = [joueur, annee, marque, serie, f"#{numero_carte}" if numero_carte != "à vérifier" else "", equipe, "hockey card"]
    clean = [e.strip() for e in elements if e and e != "à vérifier"]
    return " ".join(clean) if clean else "hockey card à vérifier"


def build_links(query: str) -> tuple[str, str]:
    encoded = quote_plus(query)
    ebay = f"https://www.ebay.com/sch/i.html?_nkw={encoded}&LH_Sold=1&LH_Complete=1"
    point130 = f"https://130point.com/sales/?q={encoded}"
    return ebay, point130


def analyze_image(path: Path) -> CardResult:
    tokens = tokenize_stem(path.stem)

    joueur = guess_player(tokens)
    annee = find_year(tokens)
    marque = find_brand(tokens)
    serie = guess_series(tokens, marque)
    insert = guess_insert(tokens)
    numero_carte = find_card_number(tokens)
    equipe = find_team(tokens)

    requete = build_price_query(joueur, annee, marque, serie, numero_carte, equipe)
    lien_ebay, lien_130point = build_links(requete)

    return CardResult(
        fichier_photo=path.name,
        joueur=joueur,
        annee=annee,
        marque=marque,
        serie=serie,
        insert=insert,
        numero_carte=numero_carte,
        equipe=equipe,
        requete_prix=requete,
        lien_ebay=lien_ebay,
        lien_130point=lien_130point,
        prix_bas="à vérifier",
        prix_moyen="à vérifier",
        prix_haut="à vérifier",
        statut="analyse basique (prix à vérifier)",
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
        # Ligne vide de démonstration pour aider l'utilisateur à voir la structure.
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
